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
        body {{ font-family: sans-serif; margin: 20px; }}
        #toolbar {{ margin-bottom: 16px; padding: 8px 0; background: #f0f0f0; border-radius: 6px; display: flex; align-items: center; gap: 16px; }}
        #toolbar label {{ margin-right: 4px; }}
        #toolbar input[type='number'] {{ width: 48px; }}
        #toolbar button.preset {{ margin-left: 2px; }}
        #grid {{ display: grid; grid-template-columns: repeat({col_num}, 1fr); gap: 10px; }}
        .series-block {{ display: flex; flex-direction: column; align-items: center; position: relative; }}
        .dicom-img {{ display: block; position: relative; z-index: 1; }}
        .roi-canvas {{ position: absolute; left: 0; top: 0; z-index: 2; pointer-events: auto; }}
        .slider {{ width: 25vw; }}
        #global-slider-block {{ margin-top: 20px; text-align: center; }}
        #global-slider {{ width: 50vw; }}
        .info-panel {{ margin-top: 5px; font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px; min-width: 200px; }}
        .dicom-img, .roi-canvas {{
            max-width: 300px;
            max-height: 300px;
        }}
    </style>
</head>
<body>
    <div id="toolbar">
        <div style="display: flex; align-items: center; gap: 8px;">
            <span>モード:</span>
            <div style="display: flex; align-items: center; background: #e0e0e0; border-radius: 20px; padding: 2px; position: relative;">
                <button id="sync-mode-btn" style="border: none; background: #4CAF50; color: white; padding: 6px 12px; border-radius: 18px; cursor: pointer; font-size: 12px; transition: all 0.3s;">同期モード</button>
                <button id="indep-mode-btn" style="border: none; background: transparent; color: #666; padding: 6px 12px; border-radius: 18px; cursor: pointer; font-size: 12px; transition: all 0.3s;">独立モード</button>
            </div>
        </div>
        <label>ROIサイズ:</label>
        <input type="number" id="roi-width" min="3" max="200" value="10"> ×
        <input type="number" id="roi-height" min="3" max="200" value="10">
        <span>プリセット:</span>
        <button class="preset" data-w="5" data-h="5">5×5</button>
        <button class="preset" data-w="10" data-h="10">10×10</button>
        <button class="preset" data-w="20" data-h="20">20×20</button>
        <button class="preset" data-w="50" data-h="50">50×50</button>
    </div>
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
        const imgs = [], sliders = [], labels = [], canvases = [], infoPanels = [];
        for (let i = 0; i < seriesCount; i++) {{
            imgs.push(document.getElementById('dicom-img-' + i));
            sliders.push(document.getElementById('slider-' + i));
            labels.push(document.getElementById('slice-label-' + i));
            canvases.push(document.getElementById('roi-canvas-' + i));
            infoPanels.push(document.getElementById('info-panel-' + i));
        }}
        // ROIサイズ管理
        let roiW = 10, roiH = 10;
        document.getElementById('roi-width').addEventListener('change', function() {{
            let v = parseInt(this.value);
            if (isNaN(v) || v < 3) v = 3;
            if (v > 200) v = 200;
            roiW = v;
            this.value = v;
        }});
        document.getElementById('roi-height').addEventListener('change', function() {{
            let v = parseInt(this.value);
            if (isNaN(v) || v < 3) v = 3;
            if (v > 200) v = 200;
            roiH = v;
            this.value = v;
        }});
        document.querySelectorAll('button.preset').forEach(btn => {{
            btn.addEventListener('click', function() {{
                roiW = parseInt(this.dataset.w);
                roiH = parseInt(this.dataset.h);
                document.getElementById('roi-width').value = roiW;
                document.getElementById('roi-height').value = roiH;
            }});
        }});
        // モード管理
        let syncMode = true;
        let roiCoords = Array(seriesCount).fill(null); // [{{x: x, y: y}}] or null
        
        function updateModeButtons() {{
            const syncBtn = document.getElementById('sync-mode-btn');
            const indepBtn = document.getElementById('indep-mode-btn');
            if (syncMode) {{
                syncBtn.style.background = '#4CAF50';
                syncBtn.style.color = 'white';
                indepBtn.style.background = 'transparent';
                indepBtn.style.color = '#666';
            }} else {{
                syncBtn.style.background = 'transparent';
                syncBtn.style.color = '#666';
                indepBtn.style.background = '#4CAF50';
                indepBtn.style.color = 'white';
            }}
        }}
        
        function switchMode(newMode) {{
            if (syncMode === newMode) return;
            
            // モード切替時の座標保持オプション
            let keep = confirm('モード切替時にROI座標を保持しますか？（OK:保持/キャンセル:リセット）');
            syncMode = newMode;
            updateModeButtons();
            
            if (!keep) {{
                roiCoords = Array(seriesCount).fill(null);
                redrawAllROIs();
            }} else if (syncMode) {{
                // 独立→同期: 先頭画像のROIを全画像にコピー
                if (roiCoords[0]) {{
                    for (let i = 1; i < seriesCount; i++) roiCoords[i] = {{x: roiCoords[0].x, y: roiCoords[0].y}};
                    redrawAllROIs();
                }}
            }}
        }}
        
        document.getElementById('sync-mode-btn').addEventListener('click', function() {{
            switchMode(true);
        }});
        
        document.getElementById('indep-mode-btn').addEventListener('click', function() {{
            switchMode(false);
        }});
        
        // 初期状態の設定
        updateModeButtons();
        // ROI描画・操作
        function redrawAllROIs() {{
            for (let i = 0; i < seriesCount; i++) drawROI(i);
        }}
        function drawROI(idx) {{
            const canvas = canvases[idx];
            const img = imgs[idx];
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            if (!roiCoords[idx]) return;
            // imgのスケール補正
            const scaleX = canvas.width / img.naturalWidth;
            const scaleY = canvas.height / img.naturalHeight;
            const x = roiCoords[idx].x * scaleX;
            const y = roiCoords[idx].y * scaleY;
            const w = roiW * scaleX;
            const h = roiH * scaleY;
            ctx.strokeStyle = 'red';
            ctx.lineWidth = 2;
            ctx.setLineDash([4,2]);
            ctx.globalAlpha = 0.7;
            ctx.strokeRect(x, y, w, h);
            ctx.globalAlpha = 0.2;
            ctx.fillStyle = 'red';
            ctx.fillRect(x, y, w, h);
            ctx.globalAlpha = 1.0;
        }}
        // canvasイベント
        for (let i = 0; i < seriesCount; i++) {{
            const canvas = canvases[i];
            const img = imgs[i];
            function resizeCanvas() {{
                canvas.width = img.width;
                canvas.height = img.height;
                canvas.style.left = img.offsetLeft + 'px';
                canvas.style.top = img.offsetTop + 'px';
            }}
            img.onload = resizeCanvas;
            window.addEventListener('resize', resizeCanvas);
            resizeCanvas();
            canvas.style.pointerEvents = 'auto';
            canvas.addEventListener('mousedown', function(e) {{
                const rect = canvas.getBoundingClientRect();
                // imgのスケール補正
                const scaleX = img.naturalWidth / rect.width;
                const scaleY = img.naturalHeight / rect.height;
                let x = Math.round((e.clientX - rect.left) * scaleX);
                let y = Math.round((e.clientY - rect.top) * scaleY);
                // 画像範囲外ならエラーポップアップを表示
                if (x < 0 || y < 0 || x + roiW > img.naturalWidth || y + roiH > img.naturalHeight) {{
                    showErrorPopup(canvas, e.clientX - rect.left, e.clientY - rect.top, 'ROIが画像範囲外です');
                    return;
                }}
                if (syncMode) {{
                    for (let j = 0; j < seriesCount; j++) roiCoords[j] = {{x: x, y: y}};
                    redrawAllROIs();
                    updateAllStats();
                }} else {{
                    roiCoords[i] = {{x: x, y: y}};
                    drawROI(i);
                    updateStats(i);
                }}
            }});
            // ROIエラーポップアップ関数
            function showErrorPopup(canvas, x, y, msg) {{
                let popup = document.createElement('div');
                popup.textContent = msg;
                popup.style.position = 'absolute';
                popup.style.left = (x + canvas.offsetLeft) + 'px';
                popup.style.top = (y + canvas.offsetTop - 24) + 'px';
                popup.style.background = 'rgba(255,0,0,0.85)';
                popup.style.color = 'white';
                popup.style.fontSize = '12px';
                popup.style.padding = '2px 8px';
                popup.style.borderRadius = '4px';
                popup.style.pointerEvents = 'none';
                popup.style.zIndex = 1000;
                popup.style.boxShadow = '0 2px 6px rgba(0,0,0,0.2)';
                canvas.parentNode.appendChild(popup);
                setTimeout(() => {{
                    if (popup.parentNode) popup.parentNode.removeChild(popup);
                }}, 1000);
            }}
        }}
        // ROI統計更新
        async function updateStats(idx) {{
            if (!roiCoords[idx]) return;
            const x = roiCoords[idx].x, y = roiCoords[idx].y;
            const stats = await window.pywebview.api.get_roi_stats(idx, currentSlices[idx], x, y, roiW, roiH);
            const img = imgs[idx];
            infoPanels[idx].innerHTML = `画像サイズ: ${{img.naturalWidth}}x${{img.naturalHeight}}<br>ROI: [${{x}},${{y}}] ${{roiW}}x${{roiH}}<br>平均: ${{stats.mean}}<br>標準偏差: ${{stats.std}}`;
        }}
        function updateAllStats() {{
            for (let i = 0; i < seriesCount; i++) updateStats(i);
        }}
        // スライダーで画像切り替え時にcanvasクリア&ROI再描画
        for (let i = 0; i < seriesCount; i++) {{
            sliders[i].addEventListener('input', async function() {{
                const idx = sliders[i].value;
                currentSlices[i] = parseInt(idx);
                labels[i].textContent = `Slice: ${{idx}}/${{seriesMaxIdxList[i]}}`;
                const b64 = await window.pywebview.api.get_single_slice(i, idx);
                imgs[i].src = 'data:image/png;base64,' + b64;
                // 画像切り替え後にcanvasリサイズ・ROI再描画・統計更新
                setTimeout(() => {{
                    const img = imgs[i];
                    const canvas = canvases[i];
                    canvas.width = img.width;
                    canvas.height = img.height;
                    canvas.style.left = img.offsetLeft + 'px';
                    canvas.style.top = img.offsetTop + 'px';
                    drawROI(i);
                    updateStats(i);
                }}, 50);
            }});
        }}
        // グローバルスライダーも同様
        const globalSlider = document.getElementById('global-slider');
        const globalLabel = document.getElementById('global-slice-label');
        globalSlider.addEventListener('input', async function() {{
            const idx = globalSlider.value;
            globalLabel.textContent = `Slice: ${{idx}}/${{globalSlider.max}}`;
            for (let i = 0; i < seriesCount; i++) {{
                currentSlices[i] = parseInt(idx);
                sliders[i].value = idx;
                labels[i].textContent = `Slice: ${{idx}}/${{seriesMaxIdxList[i]}}`;
                const img = imgs[i];
                const canvas = canvases[i];
                img.onload = function() {{
                    canvas.width = img.width;
                    canvas.height = img.height;
                    canvas.style.left = img.offsetLeft + 'px';
                    canvas.style.top = img.offsetTop + 'px';
                    drawROI(i);
                    updateStats(i);
                }};
            }}
            const b64list = await window.pywebview.api.get_slice(idx);
            for (let i = 0; i < seriesCount; i++) {{
                imgs[i].src = 'data:image/png;base64,' + b64list[i];
            }}
        }});
    </script>
</body>
</html>
'''

class DicomWebApi:
    def __init__(self, dicom_folders):
        self.dicom_folders = dicom_folders
        self.images_list, self.original_images_list = self.load_all_folders(dicom_folders)
        self.series_count = len(self.images_list)
        self.series_max_idx_list = [len(images) - 1 for images in self.images_list]
        self.global_max_idx = min(self.series_max_idx_list) if self.series_count > 0 else 0

    def load_all_folders(self, folders):
        all_images = []
        all_original_images = []
        for folder in folders:
            dicom_files = glob.glob(os.path.join(folder, '*.dcm'))
            # DICOMファイルをImagePositionPatient[2]でソート
            try:
                dicom_files.sort(key=lambda x: pydicom.dcmread(x, force=True).ImagePositionPatient[2])
            except Exception as e:
                print(f"DICOMソートエラー: {e}")
                # ソートに失敗した場合はファイル名でソート
                dicom_files.sort()
            images = []
            original_images = []
            for f in dicom_files:
                try:
                    ds = pydicom.dcmread(f, force=True)
                    arr = ds.pixel_array
                    # 元データを保存（HU値）
                    original_arr = arr.astype(np.float32) + ds.RescaleIntercept
                    original_images.append(original_arr)
                    # 表示用に正規化
                    arr = self.normalize(arr)
                    images.append(arr)
                except Exception as e:
                    print(f"読み込み失敗: {f} {e}")
            all_images.append(images)
            all_original_images.append(original_images)
        return all_images, all_original_images

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
        # 画像サイズ取得（最初の画像のshapeを使う）
        img_shapes = [images[0].shape if len(images) > 0 else (512, 512) for images in self.images_list]
        # フォルダ名を取得
        folder_names = [os.path.basename(folder) for folder in self.dicom_folders]
        series_blocks = '\n        '.join([
            f'<div class="series-block" style="position:relative;">'
            f'<div style="margin-bottom: 8px; font-weight: bold;">{folder_names[i]}</div>'
            f'<img class="dicom-img" id="dicom-img-{i}" src="data:image/png;base64,{b64}" width="{img_shapes[i][1]}" height="{img_shapes[i][0]}" style="display:block; position:relative; z-index:1;" />'
            f'<canvas class="roi-canvas" id="roi-canvas-{i}" width="{img_shapes[i][1]}" height="{img_shapes[i][0]}" style="position:absolute; left:0; top:0; z-index:2; pointer-events:auto;"></canvas>'
            f'<div class="info-panel" id="info-panel-{i}"></div>'
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