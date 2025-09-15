from core.dicom_loader import DicomLoader
from core.image_processor import ImageProcessor
from core.data_manager import DataManager
from core.exporter import Exporter
from gui.web_controller import WebController


class DicomWebApi(DicomLoader, ImageProcessor, DataManager, WebController, Exporter):
    def __init__(self, dicom_folders):
        self.dicom_folders = dicom_folders
        self.images_list, self.original_images_list, self.file_names_list = self.load_all_folders(dicom_folders)
        self.series_count = len(self.images_list)
        self.series_max_idx_list = [len(images) - 1 for images in self.images_list]
        self.global_max_idx = min(self.series_max_idx_list) if self.series_count > 0 else 0
        self.match_contrast_enabled = False
from core.dicom_loader import DicomLoader
from core.image_processor import ImageProcessor
from core.data_manager import DataManager
from gui.web_controller import WebController
from core.exporter import Exporter


class DicomWebApi(DicomLoader, ImageProcessor, DataManager, WebController, Exporter):
    def __init__(self, dicom_folders):
        self.dicom_folders = dicom_folders
        self.images_list, self.original_images_list, self.file_names_list = self.load_all_folders(dicom_folders)
        self.series_count = len(self.images_list)
        self.series_max_idx_list = [len(images) - 1 for images in self.images_list]
        self.global_max_idx = min(self.series_max_idx_list) if self.series_count > 0 else 0
        # 諧調揃え状態
        self.match_contrast_enabled = False