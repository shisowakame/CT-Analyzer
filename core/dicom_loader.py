import os
import glob
import numpy as np
import pydicom

class DicomLoader:
    def load_all_folders(self, folders):
        all_images = []
        all_original_images = []
        all_file_names = []
        all_subfolders = []
        folder_types = []
        
        for folder in folders:
            subfolders = []
            for item in os.listdir(folder):
                item_path = os.path.join(folder, item)
                if os.path.isdir(item_path):
                    dicom_files = glob.glob(os.path.join(item_path, '*.dcm'))
                    if dicom_files:
                        subfolders.append(item_path)
            
            if len(subfolders) > 1:
                folder_types.append('folder2')
                all_subfolders.append(subfolders)
                current_subfolder = subfolders[0]
            else:
                folder_types.append('folder1')
                all_subfolders.append([folder])
                current_subfolder = folder
            
            images, original_images, file_names = self.load_single_folder(current_subfolder)
            all_images.append(images)
            all_original_images.append(original_images)
            all_file_names.append(file_names)
        
        self.all_subfolders = all_subfolders
        self.folder_types = folder_types
        return all_images, all_original_images, all_file_names
    
    def load_single_folder(self, folder):
        dicom_files = glob.glob(os.path.join(folder, '*.dcm'))
        try:
            dicom_files.sort(key=lambda x: pydicom.dcmread(x, force=True).ImagePositionPatient[2])
        except Exception as e:
            print(f"DICOMソートエラー: {e}")
            dicom_files.sort()
        
        images = []
        original_images = []
        file_names = []
        for f in dicom_files:
            try:
                ds = pydicom.dcmread(f, force=True)
                arr = ds.pixel_array
                original_arr = arr.astype(np.float32) + ds.RescaleIntercept
                original_images.append(original_arr)
                arr = self.normalize(arr)
                images.append(arr)
                file_names.append(os.path.basename(f))
            except Exception as e:
                print(f"読み込み失敗: {f} {e}")
        
        return images, original_images, file_names

    def normalize(self, arr):
        arr = arr.astype(np.float32)
        arr = (arr - arr.min()) / (arr.max() - arr.min() + 1e-8) * 255.0
        return arr.astype(np.uint8)

    