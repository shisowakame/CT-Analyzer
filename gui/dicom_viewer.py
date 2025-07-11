import os
import glob
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider
from matplotlib import gridspec
import pydicom as dicom

class DicomViewer:
    def __init__(self, parent=None):
        self.parent = parent
        self.fig = None
        self.axes = []
        self.image_tables = []
        self.dicom_data = []  # [フォルダ][スライス][H][W]
        self.slider = None
        self.current_slice = 0
        self.min_file_num = 0

    def load_folders(self, folder_paths, image_type='dcm'):
        """
        複数のDICOMフォルダを読み込み、NumPy配列として保持する。
        """
        file_num_list = []
        dcm_path_list = []
        for dcm_folder in folder_paths:
            if not os.path.isdir(dcm_folder):
                print(f"ディレクトリではありません: {dcm_folder}")
                continue
            dicom_files = glob.glob(os.path.join(dcm_folder, f'*.{image_type}'))
            sorted_files = self.sort_dicom_files(dicom_files)
            dcm_path_list.append(sorted_files)
            file_num_list.append(len(sorted_files))
        if not file_num_list:
            print("有効なDICOMフォルダがありません")
            return
        self.min_file_num = min(file_num_list)
        all_images = []
        for dcm_paths in dcm_path_list:
            images = [self.dicom2ndarray(dcm_path) for dcm_path in dcm_paths[:self.min_file_num] if self.dicom2ndarray(dcm_path) is not None]
            all_images.append(images)
        self.dicom_data = np.array(all_images)
        self.current_slice = self.min_file_num // 2

    def dicom2ndarray(self, dicom_file):
        try:
            ref = dicom.dcmread(dicom_file, force=True)
            img = ref.pixel_array
            return img
        except Exception as e:
            print(f"ファイル読み込みエラー: {e}")
            return None

    def sort_dicom_files(self, dicom_files):
        try:
            dicom_files.sort(key=lambda x: dicom.dcmread(x, force=True).ImagePositionPatient[2])
        except Exception as e:
            print(f"DICOMソートエラー: {e}")
        return dicom_files

    def setup_figure(self):
        if self.dicom_data is None or len(self.dicom_data) == 0:
            print("画像データがありません")
            return
        dir_num = len(self.dicom_data)
        self.fig = plt.figure()
        gs = gridspec.GridSpec(2, dir_num, height_ratios=[20, 1])
        self.axes = []
        self.image_tables = []
        for i, images in enumerate(self.dicom_data):
            ax = self.fig.add_subplot(gs[0, i])
            ax.axis('off')
            self.axes.append(ax)
            image_table = ax.imshow(images[self.current_slice], cmap='gray')
            self.image_tables.append(image_table)
        ax_slicer = self.fig.add_subplot(gs[1, :])
        self.slider = Slider(ax=ax_slicer, label="position", valinit=self.current_slice, valmin=0, valmax=self.min_file_num - 1, valfmt='%d', orientation='horizontal')
        self.slider.on_changed(self.update_images)

    def update_images(self, val):
        idx = int(self.slider.val)
        for images, image_table in zip(self.dicom_data, self.image_tables):
            image_table.set_data(images[idx])
        self.fig.canvas.draw_idle()
        self.current_slice = idx

    def show(self):
        self.setup_figure()
        plt.show()

# --- テスト用 ---
if __name__ == '__main__':
    import argparse
    def single_dicom_viewer_arguments(args_list=None):
        parser = argparse.ArgumentParser()
        parser.add_argument('img_folders', type=str, nargs='*', help='dcmファイルがまとめられているフォルダを入力')
        parser.add_argument('--image_type', '-it', type=str, default='dcm', help='対象とする画像の拡張子を指定')
        args = parser.parse_args(args_list)
        return args
    args = single_dicom_viewer_arguments()
    viewer = DicomViewer()
    viewer.load_folders(args.img_folders, image_type=args.image_type)
    viewer.show() 