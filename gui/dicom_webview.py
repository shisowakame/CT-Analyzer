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
        <button id="reset-roi-btn" style="border: none; background: #f44336; color: white; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-left: 8px;">ROIリセット</button>
        <button id="show-history-btn" style="border: none; background: #607d8b; color: white; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-left: 8px;">履歴表示</button>
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
{history_popup}
    <script>
        const seriesCount = {series_count};
        const seriesMaxIdxList = {series_max_idx_list};
        let currentSlices = Array(seriesCount).fill(0);
        const imgs = [], sliders = [], labels = [], canvases = [], infoPanels = [], filenames = [];
        for (let i = 0; i < seriesCount; i++) {{
            imgs.push(document.getElementById('dicom-img-' + i));
            sliders.push(document.getElementById('slider-' + i));
            labels.push(document.getElementById('slice-label-' + i));
            canvases.push(document.getElementById('roi-canvas-' + i));
            infoPanels.push(document.getElementById('info-panel-' + i));
            filenames.push(document.getElementById('filename-' + i));
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
            
            syncMode = newMode;
            updateModeButtons();
            
            if (syncMode) {{
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
        
        // ROIリセットボタン
        document.getElementById('reset-roi-btn').addEventListener('click', function() {{
            roiCoords = Array(seriesCount).fill(null);
            redrawAllROIs();
            // 統計情報もクリア
            for (let i = 0; i < seriesCount; i++) {{
                infoPanels[i].innerHTML = '';
            }}
        }});
        
        // 初期状態の設定
        updateModeButtons();
        
        // フォルダ選択機能
        async function showFolderSelector(seriesIdx) {{
            // フォルダ2の場合のみ選択機能を有効にする
            const folderType = await window.pywebview.api.get_folder_type(seriesIdx);
            if (folderType !== 'folder2') return;
            
            const selector = document.getElementById('folder-selector-' + seriesIdx);
            if (selector.style.display === 'none') {{
                // 他のセレクターを閉じる
                for (let i = 0; i < seriesCount; i++) {{
                    const otherSelector = document.getElementById('folder-selector-' + i);
                    if (otherSelector) otherSelector.style.display = 'none';
                }}
                // フォルダリストを表示
                const folderList = await window.pywebview.api.get_folder_list(seriesIdx);
                selector.innerHTML = '';
                folderList.forEach((folder, idx) => {{
                    const div = document.createElement('div');
                    div.textContent = folder;
                    div.style.padding = '4px 8px';
                    div.style.cursor = 'pointer';
                    div.style.borderBottom = '1px solid #eee';
                    div.onmouseover = function() {{ this.style.backgroundColor = '#f0f0f0'; }};
                    div.onmouseout = function() {{ this.style.backgroundColor = 'white'; }};
                    div.onclick = function() {{ 
                        selectFolder(seriesIdx, idx);
                        selector.style.display = 'none';
                    }};
                    selector.appendChild(div);
                }});
                selector.style.display = 'block';
            }} else {{
                selector.style.display = 'none';
            }}
        }}
        
        async function selectFolder(seriesIdx, folderIdx) {{
            const result = await window.pywebview.api.switch_folder(seriesIdx, folderIdx);
            if (result.success) {{
                // 画像を更新
                const b64 = await window.pywebview.api.get_single_slice(seriesIdx, 0);
                imgs[seriesIdx].src = 'data:image/png;base64,' + b64;
                
                // スライダーをリセット
                sliders[seriesIdx].value = 0;
                sliders[seriesIdx].max = result.max_idx;
                currentSlices[seriesIdx] = 0;
                labels[seriesIdx].textContent = `Slice: 0/${{result.max_idx}}`;
                
                // ファイル名を更新
                const filename = await window.pywebview.api.get_filename(seriesIdx, 0);
                filenames[seriesIdx].textContent = filename;
                
                // ROIはリセットしない（保持）
                drawROI(seriesIdx);
                updateStats(seriesIdx);
                
                // フォルダ名を更新
                const folderName = await window.pywebview.api.get_current_folder_name(seriesIdx);
                const folderElement = document.querySelector(`[onclick="showFolderSelector(${{seriesIdx}})"]`);
                folderElement.textContent = folderName + ' ▼';
            }}
        }}
        
        // クリック以外でセレクターを閉じる
        document.addEventListener('click', function(e) {{
            if (!e.target.closest('.series-block')) {{
                for (let i = 0; i < seriesCount; i++) {{
                    document.getElementById('folder-selector-' + i).style.display = 'none';
                }}
            }}
        }});
        
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
                const filename = await window.pywebview.api.get_filename(i, idx);
                filenames[i].textContent = filename;
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
                const filename = await window.pywebview.api.get_filename(i, idx);
                filenames[i].textContent = filename;
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

        // 履歴機能
        let historyData = [];
        const showHistoryBtn = document.getElementById('show-history-btn');
        const historyPopup = document.getElementById('history-popup');
        const historyContent = document.getElementById('history-content');
        const closeHistoryBtn = document.getElementById('close-history-btn');
        const saveHistoryBtn = document.getElementById('save-history-btn');
        const resetHistoryBtn = document.getElementById('reset-history-btn');
        const historyTableBlock = document.getElementById('history-table-block');
        
        showHistoryBtn.addEventListener('click', function() {{
          historyPopup.style.display = 'flex';
          renderHistoryTable();
        }});
        
        closeHistoryBtn.addEventListener('click', function() {{
          historyPopup.style.display = 'none';
        }});
        
        historyPopup.addEventListener('click', function(e) {{
          if (e.target === historyPopup) historyPopup.style.display = 'none';
        }});
        
        function getCurrentStats() {{
          let row = [];
          for (let i = 0; i < seriesCount; i++) {{
            let mean = '';
            let std = '';
            if (roiCoords[i]) {{
              const info = infoPanels[i].innerText;
              const m = info.match(/平均: ([\d\.-]+)/);
              const s = info.match(/標準偏差: ([\d\.-]+)/);
              mean = m ? m[1] : '';
              std = s ? s[1] : '';
            }}
            row.push({{mean, std}});
          }}
          return row;
        }}
        
        function renderHistoryTable() {{
          let html = '<table style="font-size:11px; border-collapse:collapse; width:100%; min-width:400px;">';
          // ヘッダ行
          html += '<tr>';
          html += '<th style="border-bottom:1px solid #bbb; padding:2px 4px; width:40px;">番号</th>';
          for (let i = 0; i < seriesCount; i++) {{
            html += `<th colspan="2" style="border-bottom:1px solid #bbb; padding:2px 4px;">${{document.querySelector(`[onclick="showFolderSelector(${{i}})"]`)?.textContent.replace(' ▼','') || 'Series ' + i}}</th>`;
          }}
          html += '<th style="border-bottom:1px solid #bbb; padding:2px 4px; width:60px;">操作</th>';
          html += '</tr>';
          // サブヘッダ行
          html += '<tr>';
          html += '<th style="border-bottom:1px solid #bbb; padding:2px 4px;"></th>';
          for (let i = 0; i < seriesCount; i++) {{
            html += '<th style="border-bottom:1px solid #bbb; padding:2px 4px;">平均</th><th style="border-bottom:1px solid #bbb; padding:2px 4px;">標準偏差</th>';
          }}
          html += '<th style="border-bottom:1px solid #bbb; padding:2px 4px;"></th>';
          html += '</tr>';
          // データ行
          for (let r = 0; r < historyData.length; r++) {{
            html += '<tr>';
            html += `<td style="border-bottom:1px solid #eee; padding:2px 4px; text-align:center; font-weight:bold;">${{r + 1}}</td>`;
            for (let i = 0; i < seriesCount; i++) {{
              html += `<td style="border-bottom:1px solid #eee; padding:2px 4px; text-align:right;">${{historyData[r][i]?.mean || ''}}</td><td style="border-bottom:1px solid #eee; padding:2px 4px; text-align:right;">${{historyData[r][i]?.std || ''}}</td>`;
            }}
            html += `<td style="border-bottom:1px solid #eee; padding:2px 4px; text-align:center;"><button onclick="deleteHistoryRow(${{r}})" style="background:#f44336; color:white; border:none; border-radius:2px; padding:1px 4px; font-size:10px; cursor:pointer;">削除</button></td>`;
            html += '</tr>';
          }}
          // 平均行
          if (historyData.length > 0) {{
            html += '<tr style="background:#f4f4f4;">';
            html += '<td style="font-weight:bold; text-align:center;">平均</td>';
            for (let i = 0; i < seriesCount; i++) {{
              let meanSum = 0, stdSum = 0, cnt = 0;
              for (let r = 0; r < historyData.length; r++) {{
                if (historyData[r][i]?.mean) {{ meanSum += parseFloat(historyData[r][i].mean); cnt++; }}
              }}
              for (let r = 0; r < historyData.length; r++) {{
                if (historyData[r][i]?.std) {{ stdSum += parseFloat(historyData[r][i].std); }}
              }}
              html += `<td style="font-weight:bold; text-align:right;">${{cnt ? (meanSum/cnt).toFixed(4) : ''}}</td><td style="font-weight:bold; text-align:right;">${{cnt ? (stdSum/cnt).toFixed(4) : ''}}</td>`;
            }}
            html += '<td style="text-align:center;"></td>';
            html += '</tr>';
          }}
          html += '</table>';
          historyTableBlock.innerHTML = html;
        }}
        
        function deleteHistoryRow(rowIndex) {{
          if (rowIndex >= 0 && rowIndex < historyData.length) {{
            historyData.splice(rowIndex, 1);
            renderHistoryTable();
          }}
        }}
        
        saveHistoryBtn.addEventListener('click', function() {{
          const row = getCurrentStats();
          historyData.push(row);
          renderHistoryTable();
        }});
        
        resetHistoryBtn.addEventListener('click', function() {{
          historyData = [];
          renderHistoryTable();
        }});
        
        window.addEventListener('keydown', function(e) {{
          if ((e.ctrlKey || e.metaKey) && e.key === 's') {{
            if (historyPopup.style.display === 'flex') {{
              e.preventDefault();
              const row = getCurrentStats();
              historyData.push(row);
              renderHistoryTable();
            }}
          }}
        }});
    </script>
</body>
</html>
'''

class DicomWebApi:
    def __init__(self, dicom_folders):
        self.dicom_folders = dicom_folders
        self.images_list, self.original_images_list, self.file_names_list = self.load_all_folders(dicom_folders)
        self.series_count = len(self.images_list)
        self.series_max_idx_list = [len(images) - 1 for images in self.images_list]
        self.global_max_idx = min(self.series_max_idx_list) if self.series_count > 0 else 0

    def load_all_folders(self, folders):
        all_images = []
        all_original_images = []
        all_file_names = []
        all_subfolders = []
        folder_types = []  # 'folder2' or 'folder1'
        
        for folder in folders:
            # フォルダ2内のフォルダ1を検索
            subfolders = []
            for item in os.listdir(folder):
                item_path = os.path.join(folder, item)
                if os.path.isdir(item_path):
                    # DICOMファイルが含まれているかチェック
                    dicom_files = glob.glob(os.path.join(item_path, '*.dcm'))
                    if dicom_files:
                        subfolders.append(item_path)
            
            if len(subfolders) > 1:
                # フォルダ2（複数のフォルダ1を含む）
                folder_types.append('folder2')
                all_subfolders.append(subfolders)
                # 最初のフォルダ1を読み込み
                current_subfolder = subfolders[0]
            else:
                # フォルダ1（直接指定または単一フォルダ）
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
        # DICOMファイルをImagePositionPatient[2]でソート
        try:
            dicom_files.sort(key=lambda x: pydicom.dcmread(x, force=True).ImagePositionPatient[2])
        except Exception as e:
            print(f"DICOMソートエラー: {e}")
            # ソートに失敗した場合はファイル名でソート
            dicom_files.sort()
        
        images = []
        original_images = []
        file_names = []
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
                # ファイル名を保存
                file_names.append(os.path.basename(f))
            except Exception as e:
                print(f"読み込み失敗: {f} {e}")
        
        return images, original_images, file_names

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
        # フォルダ名を取得（現在表示中のフォルダ1の名前）
        folder_names = []
        for i in range(len(self.dicom_folders)):
            if i < len(self.all_subfolders) and self.all_subfolders[i]:
                # 現在のフォルダ1を特定
                current_first_file = self.file_names_list[i][0] if self.file_names_list[i] else ""
                for folder in self.all_subfolders[i]:
                    test_images, _, test_files = self.load_single_folder(folder)
                    if test_files and test_files[0] == current_first_file:
                        folder_names.append(os.path.basename(folder))
                        break
                else:
                    folder_names.append(os.path.basename(self.dicom_folders[i]))
            else:
                folder_names.append(os.path.basename(self.dicom_folders[i]))
        
        series_blocks = '\n        '.join([
            f'<div class="series-block" style="position:relative;">'
            + (f'<div style="margin-bottom: 8px; font-weight: bold; cursor: pointer; padding: 4px; border-radius: 4px; background: #f0f0f0;" onclick="showFolderSelector({i})">{folder_names[i]} ▼</div>'
               f'<div id="folder-selector-{i}" style="display: none; position: absolute; top: 30px; left: 0; background: white; border: 1px solid #ccc; border-radius: 4px; padding: 8px; z-index: 1000; max-height: 200px; overflow-y: auto;"></div>' if self.folder_types[i] == 'folder2' else f'<div style="margin-bottom: 8px; font-weight: bold;">{folder_names[i]}</div>')
            + f'<img class="dicom-img" id="dicom-img-{i}" src="data:image/png;base64,{b64}" width="{img_shapes[i][1]}" height="{img_shapes[i][0]}" style="display:block; position:relative; z-index:1;" />'
            f'<canvas class="roi-canvas" id="roi-canvas-{i}" width="{img_shapes[i][1]}" height="{img_shapes[i][0]}" style="position:absolute; left:0; top:0; z-index:2; pointer-events:auto;"></canvas>'
            f'<div class="info-panel" id="info-panel-{i}"></div>'
            f'<input type="range" class="slider" id="slider-{i}" min="0" max="{self.series_max_idx_list[i]}" value="0" />'
            f'<div style="display: flex; gap: 5px; margin-top: 5px;"><div style="font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px;"><span id="slice-label-{i}" style="font-weight: bold;">Slice: 0/{self.series_max_idx_list[i]}</span></div><div style="font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px;"><span id="filename-{i}">{self.file_names_list[i][0] if len(self.file_names_list[i]) > 0 else ""}</span></div></div>'
            f'</div>'
            for i, b64 in enumerate(b64list)
        ])
        col_num = min(self.series_count, 4) if self.series_count > 1 else 1
        
        # 履歴ポップアップHTML
        history_popup_html = '''
    <div id="history-popup" style="display:none; position:fixed; top:0; left:0; width:100vw; height:100vh; background:rgba(0,0,0,0.25); z-index:2000; align-items:center; justify-content:center;">
      <div id="history-content" style="background:white; border-radius:8px; padding:24px 18px 18px 18px; min-width:480px; min-height:200px; max-width:90vw; max-height:80vh; overflow:auto; box-shadow:0 4px 24px rgba(0,0,0,0.18); position:relative;">
        <button id="close-history-btn" style="position:absolute; top:8px; right:8px; background:#eee; border:none; border-radius:4px; font-size:14px; cursor:pointer;">×</button>
        <h3 style="font-size:14px; margin:0 0 8px 0;">ROI統計履歴</h3>
        <div id="history-table-block"></div>
        <div style="margin-top:10px; display:flex; gap:8px;">
          <button id="save-history-btn" style="font-size:12px; padding:4px 10px; border-radius:4px; border:none; background:#2196f3; color:white; cursor:pointer;">保存</button>
          <button id="reset-history-btn" style="font-size:12px; padding:4px 10px; border-radius:4px; border:none; background:#f44336; color:white; cursor:pointer;">リセット</button>
        </div>
        <div style="font-size:11px; color:#888; margin-top:6px;">Ctrl+Sでも保存できます</div>
      </div>
    </div>'''
        
        # HTMLテンプレートに履歴ポップアップを動的に挿入
        html_content = HTML_TEMPLATE.format(
            series_blocks=series_blocks,
            global_max_idx=self.global_max_idx,
            series_count=self.series_count,
            series_max_idx_list=self.series_max_idx_list,
            col_num=col_num,
            history_popup=history_popup_html
        )
        
        return html_content

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
            
            return {
                'success': True,
                'max_idx': len(images) - 1
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