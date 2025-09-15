import os
import numpy as np
import pydicom

class DataManager:
    def get_slice(self, idx):
        idx = int(idx)
        b64list = []
        for original_images in self.original_images_list:
            if idx < len(original_images):
                b64list.append(self.get_png_b64(original_images[idx]))
            else:
                arr = np.zeros_like(original_images[0])
                b64list.append(self.get_png_b64(arr))
        return b64list

    def get_single_slice(self, series_idx, idx):
        series_idx = int(series_idx)
        idx = int(idx)
        original_images = self.original_images_list[series_idx]
        if idx < len(original_images):
            return self.get_png_b64(original_images[idx])
        else:
            arr = np.zeros_like(original_images[0])
            return self.get_png_b64(arr)

    # Python側API: ROI統計計算
    def get_roi_stats(self, series_idx, slice_idx, x, y, w, h):
        series_idx = int(series_idx)
        slice_idx = int(slice_idx)
        x = int(x)
        y = int(y)
        w = int(w)
        h = int(h)
        # 正規化前の元データ（HU値）を使用
        arr = self.original_images_list[series_idx][slice_idx]
        roi = arr[y:y+h, x:x+w]
        roi_flat = roi.flatten()  # 1次元配列に変換
        mean = float(np.mean(roi_flat)) if roi_flat.size > 0 else 0.0
        std = float(np.std(roi_flat)) if roi_flat.size > 0 else 0.0
        return {'mean': round(mean, 8), 'std': round(std, 8)}

    # Python側API: メタデータ取得
    def get_metadata(self, series_idx, slice_idx):
        series_idx = int(series_idx)
        slice_idx = int(slice_idx)
        if series_idx < len(self.all_subfolders) and slice_idx < len(self.images_list[series_idx]):
            try:
                # 現在のファイル名からフォルダを特定
                current_filename = self.file_names_list[series_idx][slice_idx]
                current_folder = None
                for folder in self.all_subfolders[series_idx]:
                    test_images, _, test_files = self.load_single_folder(folder)
                    if test_files and current_filename in test_files:
                        current_folder = folder
                        break
                if current_folder is None:
                    current_folder = self.all_subfolders[series_idx][0]  # フォールバック
                
                ds = pydicom.dcmread(os.path.join(current_folder, current_filename), force=True)
                return str(ds)
            except Exception as e:
                return f"メタデータの読み込みに失敗しました: {e}"
        return "メタデータがありません"

    # Python側API: ファイル名取得
    def get_filename(self, series_idx, slice_idx):
        series_idx = int(series_idx)
        slice_idx = int(slice_idx)
        if series_idx < len(self.file_names_list) and slice_idx < len(self.file_names_list[series_idx]):
            return self.file_names_list[series_idx][slice_idx]
        return ""
    
    # Python側API: フォルダリスト取得
    def get_folder_list(self, series_idx):
        series_idx = int(series_idx)
        if series_idx < len(self.all_subfolders):
            return [os.path.basename(folder) for folder in self.all_subfolders[series_idx]]
        return []
    
    # Python側API: フォルダ切り替え
    def switch_folder(self, series_idx, folder_idx):
        series_idx = int(series_idx)
        folder_idx = int(folder_idx)
        if series_idx < len(self.all_subfolders) and folder_idx < len(self.all_subfolders[series_idx]):
            new_folder = self.all_subfolders[series_idx][folder_idx]
            images, original_images, file_names = self.load_single_folder(new_folder)
            self.images_list[series_idx] = images
            self.original_images_list[series_idx] = original_images
            self.file_names_list[series_idx] = file_names
            self.series_max_idx_list[series_idx] = len(images) - 1
            # フォルダ切替時も諧調揃えONなら全画像再描画のためTrueを返す
            return {
                'success': True,
                'max_idx': len(images) - 1,
                'match_contrast_enabled': self.match_contrast_enabled
            }
        return {'success': False}
    
    # Python側API: 現在のフォルダ名取得
    def get_current_folder_name(self, series_idx):
        series_idx = int(series_idx)
        if series_idx < len(self.all_subfolders):
            # 現在のフォルダを特定するために、最初のファイル名を比較
            current_first_file = self.file_names_list[series_idx][0] if self.file_names_list[series_idx] else ""
            for folder in self.all_subfolders[series_idx]:
                test_images, _, test_files = self.load_single_folder(folder)
                if test_files and test_files[0] == current_first_file:
                    return os.path.basename(folder)
            # 見つからない場合は最初のフォルダを返す
            if self.all_subfolders[series_idx]:
                return os.path.basename(self.all_subfolders[series_idx][0])
        return "Unknown"
    
    # Python側API: フォルダタイプ取得
    def get_folder_type(self, series_idx):
        series_idx = int(series_idx)
        if series_idx < len(self.folder_types):
            return self.folder_types[series_idx]
        return "folder1"
