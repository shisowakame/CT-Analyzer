import os
import glob
import base64
import io
import numpy as np
import pydicom
from PIL import Image
import webview

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>DICOM Web Viewer (Multi-Series)</title>
    <style>
        body { font-family: sans-serif; margin: 20px; }
        #grid { display: grid; grid-template-columns: repeat({col_num}, 1fr); gap: 10px; max-width: 95vw; }
        .series-block { display: flex; flex-direction: column; align-items: center; }
        .dicom-img { max-width: 45vw; max-height: 40vh; border: 1px solid #888; }
        .slider { width: 25vw; }
        #global-slider-block { margin-top: 20px; text-align: center; }
        #global-slider { width: 50vw; }
    </style>
</head>
<body>
    <h2>DICOM Web Viewer (Multi-Series)</h2>
    <div id="grid">
        {series_blocks}
    </div>
    <div id="global-slider-block">
        <input type="range" id="global-slider" min="0" max="{global_max_idx}" value="0" />
        <span id="global-slice-label">Slice: 0/{global_max_idx}</span>
    </div>
    <script>
        const seriesCount = {series_count};
        const seriesMaxIdxList = {series_max_idx_list};
        let currentSlices = Array(seriesCount).fill(0);
        const imgs = [], sliders = [], labels = [];
        for (let i = 0; i < seriesCount; i++) {
            imgs.push(document.getElementById('dicom-img-' + i));
            sliders.push(document.getElementById('slider-' + i));
            labels.push(document.getElementById('slice-label-' + i));
        }
        for (let i = 0; i < seriesCount; i++) {
            sliders[i].addEventListener('input', async function() {
                const idx = sliders[i].value;
                currentSlices[i] = parseInt(idx);
                labels[i].textContent = `Slice: ${idx}/${seriesMaxIdxList[i]}`;
                const b64 = await window.pywebview.api.get_single_slice(i, idx);
                imgs[i].src = 'data:image/png;base64,' + b64;
            });
        }
        const globalSlider = document.getElementById('global-slider');
        const globalLabel = document.getElementById('global-slice-label');
        globalSlider.addEventListener('input', async function() {
            const idx = globalSlider.value;
            globalLabel.textContent = `Slice: ${idx}/${globalSlider.max}`;
            for (let i = 0; i < seriesCount; i++) {
                currentSlices[i] = parseInt(idx);
                sliders[i].value = idx;
                labels[i].textContent = `Slice: ${idx}/${seriesMaxIdxList[i]}`;
            }
            const b64list = await window.pywebview.api.get_slice(idx);
            for (let i = 0; i < seriesCount; i++) {
                imgs[i].src = 'data:image/png;base64,' + b64list[i];
            }
        });
    </script>
</body>
</html>
'''

class DicomWebApi:
    def __init__(self, dicom_folders):
        self.images_list = self.load_all_folders(dicom_folders)
        self.series_count = len(self.images_list)
        self.series_max_idx_list = [len(images) - 1 for images in self.images_list]
        self.global_max_idx = min(self.series_max_idx_list) if self.series_count > 0 else 0

    def load_all_folders(self, folders):
        all_images = []
        for folder in folders:
            dicom_files = sorted(glob.glob(os.path.join(folder, '*.dcm')))
            images = []
            for f in dicom_files:
                try:
                    ds = pydicom.dcmread(f, force=True)
                    arr = ds.pixel_array
                    arr = self.normalize(arr)
                    images.append(arr)
                except Exception as e:
                    print(f"読み込み失敗: {f} {e}")
            all_images.append(images)
        return all_images

    def normalize(self, arr):
        arr = arr.astype(np.float32)
        arr = (arr - arr.min()) / (arr.max() - arr.min() + 1e-8) * 255.0
        return arr.astype(np.uint8)

    def get_png_b64(self, arr):
        img = Image.fromarray(arr)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
        return b64

    def get_init_html(self):
        b64list = [self.get_png_b64(images[0]) for images in self.images_list]
        series_blocks = '\n        '.join([
            f'<div class="series-block">'
            f'<img class="dicom-img" id="dicom-img-{i}" src="data:image/png;base64,{b64}" />'
            f'<input type="range" class="slider" id="slider-{i}" min="0" max="{self.series_max_idx_list[i]}" value="0" />'
            f'<span id="slice-label-{i}">Slice: 0/{self.series_max_idx_list[i]}</span>'
            f'</div>'
            for i, b64 in enumerate(b64list)
        ])
        col_num = min(self.series_count, 4) if self.series_count > 1 else 1
        return HTML_TEMPLATE.format(
            series_blocks=series_blocks,
            global_max_idx=self.global_max_idx,
            series_count=self.series_count,
            series_max_idx_list=self.series_max_idx_list,
            col_num=col_num
        )

    def get_slice(self, idx):
        idx = int(idx)
        b64list = []
        for images in self.images_list:
            if idx < len(images):
                b64list.append(self.get_png_b64(images[idx]))
            else:
                arr = np.zeros_like(images[0])
                b64list.append(self.get_png_b64(arr))
        return b64list

    def get_single_slice(self, series_idx, idx):
        series_idx = int(series_idx)
        idx = int(idx)
        images = self.images_list[series_idx]
        if idx < len(images):
            return self.get_png_b64(images[idx])
        else:
            arr = np.zeros_like(images[0])
            return self.get_png_b64(arr)

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('Usage: python dicom_webview.py <dicom_folder1> [<dicom_folder2> ...]')
        exit(1)
    folders = sys.argv[1:]
    api = DicomWebApi(folders)
    html = api.get_init_html()
    window = webview.create_window('DICOM Web Viewer (Multi-Series)', html=html, js_api=api, width=1200, height=900)
    webview.start(debug=True) 