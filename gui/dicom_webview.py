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
    <title>DICOM Web Viewer</title>
    <style>
        body {{ font-family: sans-serif; margin: 20px; }}
        #dicom-img {{ max-width: 90vw; max-height: 80vh; border: 1px solid #888; }}
        #slider {{ width: 80vw; }}
    </style>
</head>
<body>
    <h2>DICOM Web Viewer</h2>
    <img id="dicom-img" src="data:image/png;base64,{img_b64}" />
    <br>
    <input type="range" id="slider" min="0" max="{max_idx}" value="{init_idx}" />
    <span id="slice-label">Slice: {init_idx}/{max_idx}</span>
    <script>
        const slider = document.getElementById('slider');
        const img = document.getElementById('dicom-img');
        const label = document.getElementById('slice-label');
        slider.addEventListener('input', async function() {{
            const idx = slider.value;
            label.textContent = `Slice: ${{idx}}/${{slider.max}}`;
            const b64 = await window.pywebview.api.get_slice(idx);
            img.src = 'data:image/png;base64,' + b64;
        }});
    </script>
</body>
</html>
'''

class DicomWebApi:
    def __init__(self, dicom_folder):
        self.images = self.load_dicom_images(dicom_folder)
        self.max_idx = len(self.images) - 1

    def load_dicom_images(self, folder):
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
        return images

    def normalize(self, arr):
        arr = arr.astype(np.float32)
        arr = (arr - arr.min()) / (arr.max() - arr.min() + 1e-8) * 255.0
        return arr.astype(np.uint8)

    def get_png_b64(self, idx):
        arr = self.images[int(idx)]
        img = Image.fromarray(arr)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
        return b64

    def get_init_html(self):
        b64 = self.get_png_b64(0)
        return HTML_TEMPLATE.format(img_b64=b64, max_idx=self.max_idx, init_idx=0)

    def get_slice(self, idx):
        return self.get_png_b64(idx)

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('Usage: python3 dicom_webview.py <dicom_folder>')
        exit(1)
    folder = sys.argv[1]
    api = DicomWebApi(folder)
    html = api.get_init_html()
    window = webview.create_window('DICOM Web Viewer', html=html, js_api=api, width=1000, height=800)
    webview.start(debug=True) 