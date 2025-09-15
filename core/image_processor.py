import base64
import io
import os
from PIL import Image
import numpy as np
from datetime import datetime

class ImageProcessor:
    def get_min_ct_window(self):
        # 各画像のmin, max, 幅を計算し、最小幅の画像のmin, maxを基準にする
        min_base = None
        min_range = None
        min_width = None
        for series in self.original_images_list:
            for arr in series:
                arr_min = np.min(arr)
                arr_max = np.max(arr)
                width = arr_max - arr_min
                if min_width is None or width < min_width:
                    min_width = width
                    min_base = arr_min
                    min_range = width
        # min_base: 最小幅画像のmin, min_range: その幅
        # 幅が0の場合は1にする（ゼロ除算防止）
        if min_range == 0:
            min_range = 1
        return min_base, min_range

    def get_png_b64(self, arr):
        # arrは必ず生CT値（float, pixel_array+RescaleIntercept）であること
        arr_min = float(np.min(arr))
        arr_max = float(np.max(arr))
        arr_width = arr_max - arr_min
        if getattr(self, 'match_contrast_enabled', False):
            base, ct_range = self.get_min_ct_window()
            arr_disp = np.clip((arr - base) / ct_range, 0, 1)
            arr_disp_uint8 = (arr_disp * 255.0).astype(np.uint8)
            print(f"[諧調揃えON] 基準min: {base}, 基準max: {base+ct_range}, 幅: {ct_range} | スライスmin: {arr_min}, max: {arr_max}, 幅: {arr_width} | 正規化後min: {arr_disp.min()}, max: {arr_disp.max()}")
            arr = arr_disp_uint8
        else:
            arr_norm = (arr - arr_min) / (arr_width + 1e-8)
            arr_uint8 = (arr_norm * 255.0).astype(np.uint8)
            print(f"[諧調揃えOFF] スライスmin: {arr_min}, max: {arr_max}, 幅: {arr_width} | 正規化後min: {arr_norm.min()}, max: {arr_norm.max()}")
            arr = arr_uint8
        img = Image.fromarray(arr)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
        return b64

    def get_display_images_for_download(self, current_slices):
        """現在表示されている画素値（諧調調整済み）をPNGとして保存"""
        try:
            from PIL import Image
            import io
            import base64
            import numpy as np
            from datetime import datetime
            import os
            
            print("[DEBUG] ダウンロード処理開始")
            
            # 現在表示されている画像のbase64データを取得
            display_images = []
            slice_names = []
            
            print(f"[DEBUG] original_images_listの長さ: {len(self.original_images_list)}")
            print(f"[DEBUG] 受け取ったcurrent_slices: {current_slices}")
            
            for i, original_images in enumerate(self.original_images_list):
                print(f"[DEBUG] シリーズ {i} の処理開始")
                
                # 現在のスライスインデックスを取得
                current_slice_idx = current_slices[i] if i < len(current_slices) else 0
                
                print(f"[DEBUG] 現在のスライスインデックス: {current_slice_idx}")
                
                # 生CT値を取得
                arr = original_images[current_slice_idx]
                print(f"[DEBUG] 配列の形状: {arr.shape}, 型: {arr.dtype}")
                
                # 諧調調整を適用して表示用画素値に変換
                arr_min = float(np.min(arr))
                arr_max = float(np.max(arr))
                arr_width = arr_max - arr_min
                
                print(f"[DEBUG] CT値範囲: min={arr_min}, max={arr_max}, width={arr_width}")
                
                if getattr(self, 'match_contrast_enabled', False):
                    base, ct_range = self.get_min_ct_window()
                    arr_disp = np.clip((arr - base) / ct_range, 0, 1)
                    arr_disp_uint8 = (arr_disp * 255.0).astype(np.uint8)
                    print(f"[DEBUG] 諧調揃えON: base={base}, ct_range={ct_range}")
                else:
                    arr_norm = (arr - arr_min) / (arr_width + 1e-8)
                    arr_disp_uint8 = (arr_norm * 255.0).astype(np.uint8)
                    print(f"[DEBUG] 諧調揃えOFF")
                
                print(f"[DEBUG] 表示用画素値範囲: min={arr_disp_uint8.min()}, max={arr_disp_uint8.max()}")
                
                # PNGに変換
                img = Image.fromarray(arr_disp_uint8)
                display_images.append(img)
                
                # スライス名を生成
                slice_name = f"series{i+1}_slice{current_slice_idx+1}"
                slice_names.append(slice_name)
                print(f"[DEBUG] スライス名: {slice_name}")
            
            print(f"[DEBUG] 表示画像数: {len(display_images)}")
            
            # 複数画像を横に並べて結合
            if len(display_images) == 1:
                combined_image = display_images[0]
                print("[DEBUG] 単一画像のため結合なし")
            else:
                # 画像の高さを統一（最大の高さに合わせる）
                max_height = max(img.height for img in display_images)
                total_width = sum(img.width for img in display_images) + (len(display_images) - 1) * 10  # 10px間隔
                
                print(f"[DEBUG] 結合画像サイズ: {total_width}x{max_height}")
                
                # 新しい画像を作成
                combined_image = Image.new('L', (total_width, max_height), 255)
                
                # 画像を横に並べて配置
                x_offset = 0
                for img in display_images:
                    # 画像を中央揃えで配置
                    y_offset = (max_height - img.height) // 2
                    combined_image.paste(img, (x_offset, y_offset))
                    x_offset += img.width + 10  # 10px間隔
                
                print("[DEBUG] 画像結合完了")
            
            # 個々の画像を保存
            individual_files = []
            downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            for i, img in enumerate(display_images):
                individual_filename = f"{slice_names[i]}-{timestamp}.png"
                individual_file_path = os.path.join(downloads_dir, individual_filename)
                img.save(individual_file_path, format='PNG')
                individual_files.append(individual_filename)
                print(f"[DEBUG] 個別画像保存完了: {individual_file_path}")
            
            # 結合画像のファイル名を生成
            combined_filename = f"{'-'.join(slice_names)}-{timestamp}.png"
            combined_file_path = os.path.join(downloads_dir, combined_filename)
            
            # 結合画像をPNGとして保存
            combined_image.save(combined_file_path, format='PNG')
            
            print(f"[DEBUG] 結合画像保存完了: {combined_file_path}")
            
            return {
                'success': True,
                'file_path': combined_file_path,
                'filename': combined_filename,
                'individual_files': individual_files
            }
            
        except Exception as e:
            print(f"[DEBUG] エラー発生: {str(e)}")
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'error': str(e)
            }

    def get_display_images_with_roi_for_download(self, roi_coords_list, current_slices):
        """現在表示されている画素値（諧調調整済み）+ ROIをPNGとして保存"""
        try:
            from PIL import Image, ImageDraw
            import io
            import base64
            import numpy as np
            from datetime import datetime
            import os
            
            print("[DEBUG] ROI含むダウンロード処理開始")
            
            # 現在表示されている画像のbase64データを取得
            display_images = []
            slice_names = []
            
            print(f"[DEBUG] original_images_listの長さ: {len(self.original_images_list)}")
            print(f"[DEBUG] ROI座標リスト: {roi_coords_list}")
            print(f"[DEBUG] 受け取ったcurrent_slices: {current_slices}")
            
            for i, original_images in enumerate(self.original_images_list):
                print(f"[DEBUG] シリーズ {i} の処理開始")
                
                # 現在のスライスインデックスを取得
                current_slice_idx = current_slices[i] if i < len(current_slices) else 0
                
                print(f"[DEBUG] 現在のスライスインデックス: {current_slice_idx}")
                
                # 生CT値を取得
                arr = original_images[current_slice_idx]
                print(f"[DEBUG] 配列の形状: {arr.shape}, 型: {arr.dtype}")
                
                # 諧調調整を適用して表示用画素値に変換
                arr_min = float(np.min(arr))
                arr_max = float(np.max(arr))
                arr_width = arr_max - arr_min
                
                print(f"[DEBUG] CT値範囲: min={arr_min}, max={arr_max}, width={arr_width}")
                
                if getattr(self, 'match_contrast_enabled', False):
                    base, ct_range = self.get_min_ct_window()
                    arr_disp = np.clip((arr - base) / ct_range, 0, 1)
                    arr_disp_uint8 = (arr_disp * 255.0).astype(np.uint8)
                    print(f"[DEBUG] 諧調揃えON: base={base}, ct_range={ct_range}")
                else:
                    arr_norm = (arr - arr_min) / (arr_width + 1e-8)
                    arr_disp_uint8 = (arr_norm * 255.0).astype(np.uint8)
                    print(f"[DEBUG] 諧調揃えOFF")
                
                print(f"[DEBUG] 表示用画素値範囲: min={arr_disp_uint8.min()}, max={arr_disp_uint8.max()}")
                
                # PNGに変換
                img = Image.fromarray(arr_disp_uint8)
                
                # ROIを描画
                if i < len(roi_coords_list) and roi_coords_list[i] is not None:
                    roi_coords = roi_coords_list[i]
                    print(f"[DEBUG] ROI座標: {roi_coords}")
                    
                    # ROIサイズを取得（JavaScriptから送信された値を使用）
                    roi_width = roi_coords.get('width', 10)
                    roi_height = roi_coords.get('height', 10)
                    
                    # グレースケール画像をRGBに変換（ROIを色付きで描画するため）
                    img_rgb = img.convert('RGB')
                    
                    # ROIを描画
                    draw = ImageDraw.Draw(img_rgb)
                    x = roi_coords['x']
                    y = roi_coords['y']
                    w = roi_width
                    h = roi_height
                    
                    # 色を取得（デフォルトは赤）
                    roi_color = roi_coords.get('color', '#ff0000')
                    
                    # 16進数カラーコードをRGBタプルに変換
                    if roi_color.startswith('#'):
                        roi_color = roi_color[1:]  # #を除去
                        r = int(roi_color[0:2], 16)
                        g = int(roi_color[2:4], 16)
                        b = int(roi_color[4:6], 16)
                        color_tuple = (r, g, b)
                    else:
                        color_tuple = (255, 0, 0)  # デフォルトは赤
                    
                    # 指定された色で枠線を描画
                    draw.rectangle([x, y, x + w, y + h], outline=color_tuple, width=2)
                    print(f"[DEBUG] ROI描画完了: ({x}, {y}, {w}, {h}), 色: {color_tuple}")
                    
                    # RGB画像をそのまま使用
                    img = img_rgb
                
                display_images.append(img)
                
                # スライス名を生成
                slice_name = f"series{i+1}_slice{current_slice_idx+1}"
                slice_names.append(slice_name)
                print(f"[DEBUG] スライス名: {slice_name}")
            
            print(f"[DEBUG] 表示画像数: {len(display_images)}")
            
            # 複数画像を横に並べて結合
            if len(display_images) == 1:
                combined_image = display_images[0]
                print("[DEBUG] 単一画像のため結合なし")
            else:
                # 画像の高さを統一（最大の高さに合わせる）
                max_height = max(img.height for img in display_images)
                total_width = sum(img.width for img in display_images) + (len(display_images) - 1) * 10  # 10px間隔
                
                print(f"[DEBUG] 結合画像サイズ: {total_width}x{max_height}")
                
                # 新しい画像を作成（RGBまたはグレースケール）
                # ROIがある場合はRGB、ない場合はグレースケール
                has_roi = any(i < len(roi_coords_list) and roi_coords_list[i] is not None for i in range(len(display_images)))
                if has_roi:
                    combined_image = Image.new('RGB', (total_width, max_height), (255, 255, 255))
                else:
                    combined_image = Image.new('L', (total_width, max_height), 255)
                
                # 画像を横に並べて配置
                x_offset = 0
                for img in display_images:
                    # 画像を中央揃えで配置
                    y_offset = (max_height - img.height) // 2
                    combined_image.paste(img, (x_offset, y_offset))
                    x_offset += img.width + 10  # 10px間隔
                
                print("[DEBUG] 画像結合完了")
            
            # 個々の画像を保存
            individual_files = []
            downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            for i, img in enumerate(display_images):
                individual_filename = f"{slice_names[i]}-ROI-{timestamp}.png"
                individual_file_path = os.path.join(downloads_dir, individual_filename)
                img.save(individual_file_path, format='PNG')
                individual_files.append(individual_filename)
                print(f"[DEBUG] ROI含む個別画像保存完了: {individual_file_path}")
            
            # 結合画像のファイル名を生成
            combined_filename = f"{'-'.join(slice_names)}-ROI-{timestamp}.png"
            combined_file_path = os.path.join(downloads_dir, combined_filename)
            
            # 結合画像をPNGとして保存
            combined_image.save(combined_file_path, format='PNG')
            
            print(f"[DEBUG] ROI含む結合画像保存完了: {combined_file_path}")
            
            return {
                'success': True,
                'file_path': combined_file_path,
                'filename': combined_filename,
                'individual_files': individual_files
            }
            
        except Exception as e:
            print(f"[DEBUG] エラー発生: {str(e)}")
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'error': str(e)
            }