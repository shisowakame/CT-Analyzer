import os
import glob
import base64
import io
import numpy as np
import pydicom
from PIL import Image
import webview
import pandas as pd
import datetime # datetimeモジュールを追加


HTML_TEMPLATE = r'''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>DICOM Web Viewer (Multi-Series)</title>
    <style>
        body {{ font-family: sans-serif; margin: 20px; }}
        .main-container {{ display: flex; gap: 20px; align-items: flex-start; width: 100vw; box-sizing: border-box; }}
        .left-panel {{ flex: 0 0 66.66vw; width: 66.66vw; }}
        .right-panel {{ flex: 0 0 33.33vw; min-width: 300px; margin-top: 0px; width: 33.33vw; }}
        #toolbar {{ margin-bottom: 16px; padding: 8px 0; background: #f0f0f0; border-radius: 6px; display: flex; align-items: center; gap: 10px; }}
        #toolbar label {{ margin-right: 4px; }}
        #toolbar input[type='number'] {{ width: 48px; }}
        #toolbar button.preset {{ margin-left: 2px; }}
        #grid {{
            width: 66.66vw;
            margin-left: 25px;
            display: grid;
            grid-template-columns: repeat({col_num}, 1fr);
            gap: 20px;
        }}
        .series-block {{ width: 100%; }}
        .dicom-square {{ aspect-ratio: 1/1; width: 100%; position: relative; }}
        .dicom-img, .roi-canvas {{
            position: absolute;
            top: 0; left: 0;
            width: 100%;
            height: 100%;
            object-fit: contain;
            max-width: none;
            max-height: none;
            margin: 0;
            padding: 0;
            border: none;
            display: block;
        }}
        .series-block {{ display: flex; flex-direction: column; align-items: center; position: relative; }}
        .dicom-img {{ display: block; position: relative; z-index: 1; }}
        .roi-canvas {{ position: absolute; left: 0; top: 0; z-index: 2; pointer-events: auto; }}
        .slider {{ width: 20vw; }}
        #global-slider-block {{ margin-top: 20px; text-align: center; }}
        #global-slider {{ width: 40vw; }}
        .info-panel {{ margin-top: 5px; font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px; min-width: 200px; }}
        
        /* 履歴パネル用スタイル */
        .history-panel {{ background: #f9f9f9; border: 1px solid #ddd; border-radius: 6px; padding: 12px; }}
        .history-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }}
        .history-header h3 {{ margin: 0; font-size: 14px; color: #333; }}
        .history-controls {{ display: flex; gap: 8px; }}
        .history-content {{ }}
        .history-table-block {{ max-height: 70vh; overflow-y: auto; }}
        .history-table {{ font-size: 11px; border-collapse: collapse; width: 100%; }}
        .history-table th, .history-table td {{ border-bottom: 1px solid #eee; padding: 2px 4px; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .history-table th {{ border-bottom: 1px solid #bbb; background: #f5f5f5; }}
        .delete-btn {{ background: #f44336; color: white; border: none; border-radius: 2px; padding: 1px 4px; font-size: 10px; cursor: pointer; }}
        .save-btn {{ background: #2196f3; color: white; border: none; border-radius: 4px; padding: 4px 10px; font-size: 12px; cursor: pointer; }}
        .reset-btn {{ background: #f44336; color: white; border: none; border-radius: 4px; padding: 4px 10px; font-size: 12px; cursor: pointer; }}
        .export-excel-btn {{ background: #1a7340; color: white; border: none; border-radius: 4px; padding: 4px 10px; font-size: 12px; cursor: pointer; }} /* 新しいスタイル */
        .info-text {{ font-size: 11px; color: #888; margin-top: 8px; }}
        h2 {{ margin-top: 0; font-size: clamp(1rem, 1.7vw, 1.2rem); }}
        #toolbar span:first-child {{ margin-left: 10px; }}
        .right-panel {{ flex: 0 0 30vw; min-width: 300px; margin-top: 0px; width: 30vw; }}
        #history-avg-row-block {{ max-width: 30vw; width: 100%; overflow-x: auto; }}
        #history-avg-row-block > table {{ max-width: 30vw; width: 100%; box-sizing: border-box; }}
        #grid {{
            width: var(--dicom-grid-width);
            margin-left: 25px;
            display: grid;
            grid-template-columns: repeat({col_num}, 1fr);
            gap: 20px;
        }}
        #toolbar {{
            width: var(--dicom-grid-width);
            margin-left: 25px;
            margin-bottom: 16px;
            padding: 8px 0;
            background: #f0f0f0;
            border-radius: 6px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
    </style>
</head>
<body>
    <div class="main-container" style="display: flex; gap: 20px; align-items: flex-start;">
        <div class="left-panel">
            <h2>DICOM Web Viewer (Multi-Series)</h2>
            <div id="toolbar">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span>モード:</span>
                    <div style="display: flex; align-items: center; background: #e0e0e0; border-radius: 20px; padding: 2px; position: relative;">
                        <button id="sync-mode-btn" style="border: none; background: #4CAF50; color: white; padding: 6px 12px; border-radius: 18px; cursor: pointer; font-size: 12px; transition: all 0.3s;">同期</button>
                        <button id="indep-mode-btn" style="border: none; background: transparent; color: #666; padding: 6px 12px; border-radius: 18px; cursor: pointer; font-size: 12px; transition: all 0.3s;">独立</button>
                    </div>
                </div>
                <button id="reset-roi-btn" style="border: none; background: #f44336; color: white; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-left: 8px;">リセット</button>
                <label>サイズ:</label>
                <input type="number" id="roi-width" min="3" max="200" value="10"> ×
                <input type="number" id="roi-height" min="3" max="200" value="10">
                <span>プリセット:</span>
                <button class="preset" data-w="5" data-h="5">5×5</button>
                <button class="preset" data-w="10" data-h="10">10×10</button>
                <button class="preset" data-w="20" data-h="20">20×20</button>
                <button class="preset" data-w="50" data-h="50">50×50</button>
            </div>
            <div id="grid">
                {series_blocks}
            </div>
            <div id="global-slider-block">
                <input type="range" id="global-slider" min="0" max="{global_max_idx}" value="0" />
                <span id="global-slice-label">Slice: 1/{global_max_idx_plus1}</span>
            </div>
        </div>
        <div class="right-panel">
            <div class="history-panel">
                <div class="history-header">
                    <h3>ROI統計履歴</h3>
                    <div class="history-controls">
                        <button id="save-btn" class="save-btn">保存</button>
                        <button id="reset-btn" class="reset-btn">リセット</button>
                        <button id="export-excel-btn" class="export-excel-btn">Excelエクスポート</button>
                    </div>
                </div>
                <div class="history-content">
                    <div id="history-table-block" class="history-table-block"></div>
                    <div id="history-avg-row-block"></div>
                </div>
                <div class="info-text">Ctrl+Sでも保存できます</div>
            </div>
        </div>
    </div>
    
    <script>
        const seriesCount = {series_count};
        const seriesMaxIdxList = {series_max_idx_list};
        let currentSlices = Array(seriesCount).fill(0);
        const imgs = [], sliders = [], labels = [], canvases = [], infoPanels = [], filenames = [];
        // フォルダ1のベース名リスト
        const seriesFolderBaseNames = {series_folder_base_names};
        
        // 履歴データ管理
        let historyData = [];
        const historyTableBlock = document.getElementById('history-table-block');
        const saveBtn = document.getElementById('save-btn');
        const resetBtn = document.getElementById('reset-btn');
        const exportExcelBtn = document.getElementById('export-excel-btn');
        
        for (let i = 0; i < seriesCount; i++) {{
            imgs.push(document.getElementById('dicom-img-' + i));
            sliders.push(document.getElementById('slider-' + i));
            labels.push(document.getElementById('slice-label-' + i));
            canvases.push(document.getElementById('roi-canvas-' + i));
            infoPanels.push(document.getElementById('info-panel-' + i));
            filenames.push(document.getElementById('filename-' + i));
        }}
        
        // 履歴テーブル描画関数
        function renderHistoryTable() {{
            let html = '<table class="history-table">';
            // ヘッダ行
            html += '<tr>';
            html += '<th style="width: 40px;"> </th>';
            for (let i = 0; i < seriesCount; i++) {{
                html += '<th colspan="3">Folder' + (i + 1) + '</th>';
            }}
            html += '<th style="width: 60px;"> </th>';
            html += '</tr>';
            // サブヘッダ行
            html += '<tr>';
            html += '<th></th>';
            for (let i = 0; i < seriesCount; i++) {{
                html += '<th>平均</th><th>標準偏差</th><th>Info</th>';
            }}
            html += '<th></th>';
            html += '</tr>';
            // データ行
            for (let r = 0; r < historyData.length; r++) {{
                html += '<tr>';
                html += '<td style="text-align: center; font-weight: bold;">' + (r + 1) + '</td>';
                for (let i = 0; i < seriesCount; i++) {{
                    const meanVal = historyData[r][i] && historyData[r][i].mean ? historyData[r][i].mean : '';
                    const stdVal = historyData[r][i] && historyData[r][i].std ? historyData[r][i].std : '';
                    const infoVal = historyData[r][i] && historyData[r][i].info ? historyData[r][i].info : '';
                    html += '<td>' + meanVal + '</td><td>' + stdVal + '</td><td>' + infoVal + '</td>';
                }}
                html += '<td style="text-align: center;"><button onclick="deleteHistoryRow(' + r + ')" class="delete-btn">×</button></td>';
                html += '</tr>';
            }}
            html += '</table>';
            historyTableBlock.innerHTML = html;

            // 平均行分離（Info列は平均行から除外）
            let avgHtml = '';
            if (historyData.length > 0) {{
                avgHtml += '<table class="history-table" style="background: #f4f4f4;">';
                avgHtml += '<tr>';
                avgHtml += '<th style="width: 40px;">　</th>';
                for (let i = 0; i < seriesCount; i++) {{
                    avgHtml += '<th colspan="2">Folder' + (i + 1) + '</th>';
                }}
                avgHtml += '<th style="width: 60px;">　</th>';
                avgHtml += '</tr>';
                avgHtml += '<tr>';
                avgHtml += '<th></th>';
                for (let i = 0; i < seriesCount; i++) {{
                    avgHtml += '<th>平均</th><th>標準偏差</th>';
                }}
                avgHtml += '<th></th>';
                avgHtml += '</tr>';
                avgHtml += '<tr style="background: #f4f4f4;">';
                avgHtml += '<td style="font-weight: bold; text-align: center;">平均</td>';
                for (let i = 0; i < seriesCount; i++) {{
                    let meanSum = 0, stdSum = 0, cnt = 0;
                    for (let r = 0; r < historyData.length; r++) {{
                        if (historyData[r][i] && historyData[r][i].mean) {{
                            meanSum += parseFloat(historyData[r][i].mean);
                            cnt++;
                        }}
                    }}
                    for (let r = 0; r < historyData.length; r++) {{
                        if (historyData[r][i] && historyData[r][i].std) {{
                            stdSum += parseFloat(historyData[r][i].std);
                        }}
                    }}
                    const avgMean = cnt ? (meanSum/cnt).toFixed(4) : '';
                    const avgStd = cnt ? (stdSum/cnt).toFixed(4) : '';
                    avgHtml += '<td style="font-weight: bold;">' + avgMean + '</td><td style="font-weight: bold;">' + avgStd + '</td>';
                }}
                avgHtml += '<td></td>';
                avgHtml += '</tr>';
                avgHtml += '</table>';
            }}
            document.getElementById('history-avg-row-block').innerHTML = avgHtml;
        }}
        
        // 履歴行削除関数
        function deleteHistoryRow(rowIndex) {{
            if (rowIndex >= 0 && rowIndex < historyData.length) {{
                historyData.splice(rowIndex, 1);
                renderHistoryTable();
            }}
        }}
        
        // 現在の統計情報を取得
        function getCurrentStats() {{
            let row = [];
            for (let i = 0; i < seriesCount; i++) {{
                let mean = '';
                let std = '';
                let info = '';
                if (roiCoords[i]) {{
                    const infoPanel = infoPanels[i];
                    const infoText = infoPanel.innerText;
                    const m = infoText.match(/平均: ([\d\.-]+)/);
                    const s = infoText.match(/標準偏差: ([\d\.-]+)/);
                    mean = m ? m[1] : '';
                    std = s ? s[1] : '';
                    // フォルダ名
                    let folderName = '';
                    const folderElem = document.querySelector('[onclick="showFolderSelector(' + i + ')"]');
                    if (folderElem) {{
                        folderName = folderElem.textContent.replace(' ▼','');
                    }}
                    // フォルダ名が空欄の場合はseriesFolderBaseNamesから補完
                    if (!folderName) {{
                        folderName = seriesFolderBaseNames[i];
                    }}
                    // ファイル名
                    let fileName = filenames[i] ? filenames[i].textContent : '';
                    // 座標
                    let coord = '';
                    if (roiCoords[i]) {{
                        coord = '(' + roiCoords[i].x + ',' + roiCoords[i].y + ')';
                    }}
                    info = folderName + '/' + fileName + '/' + coord;
                }}
                row.push({{mean: mean, std: std, info: info}});
            }}
            return row;
        }}
        
        // 履歴関連イベントリスナー
        saveBtn.addEventListener('click', function() {{
            const row = getCurrentStats();
            historyData.push(row);
            renderHistoryTable();
        }});
        
        resetBtn.addEventListener('click', function() {{
            historyData = [];
            renderHistoryTable();
        }});

        // Excelエクスポートボタンのイベントリスナー
        exportExcelBtn.addEventListener('click', async function() {{
            if (historyData.length === 0) {{
                alert('エクスポートする履歴データがありません。');
                return;
            }}
            try {{
                const result = await window.pywebview.api.export_history_to_excel(historyData);
                if (result.success) {{
                    alert('履歴がExcelファイルとして保存されました:\\n' + result.filePath);
                }} else {{
                    alert('Excelファイルの保存に失敗しました:\\n' + result.message);
                }}
            }} catch (error) {{
                console.error('Excelエクスポートエラー:', error);
                alert('Excelファイルの保存中にエラーが発生しました。');
            }}
        }});
        
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
        let roiCoords = Array(seriesCount).fill(null);
        
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
            for (let i = 0; i < seriesCount; i++) {{
                infoPanels[i].innerHTML = '';
            }}
        }});
        
        updateModeButtons();
        
        // フォルダ選択機能
        async function showFolderSelector(seriesIdx) {{
            const folderType = await window.pywebview.api.get_folder_type(seriesIdx);
            if (folderType !== 'folder2') return;
            
            const selector = document.getElementById('folder-selector-' + seriesIdx);
            if (selector.style.display === 'none') {{
                for (let i = 0; i < seriesCount; i++) {{
                    const otherSelector = document.getElementById('folder-selector-' + i);
                    if (otherSelector) otherSelector.style.display = 'none';
                }}
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
                const b64 = await window.pywebview.api.get_single_slice(seriesIdx, 0);
                imgs[seriesIdx].src = 'data:image/png;base64,' + b64;
                
                sliders[seriesIdx].value = 0;
                sliders[seriesIdx].max = result.max_idx;
                currentSlices[seriesIdx] = 0;
                labels[seriesIdx].textContent = 'Slice: 1/' + (result.max_idx + 1);
                
                const filename = await window.pywebview.api.get_filename(seriesIdx, 0);
                filenames[seriesIdx].textContent = filename;
                
                drawROI(seriesIdx);
                updateStats(seriesIdx);
                
                const folderName = await window.pywebview.api.get_current_folder_name(seriesIdx);
                const folderElement = document.querySelector('[onclick="showFolderSelector(' + seriesIdx + ')"]');
                folderElement.textContent = folderName + ' ▼';
                
                // 全体シークバーの分母を更新
                const globalMaxIdx = Math.min(...Array.from({{length: seriesCount}}, (_, i) => parseInt(sliders[i].max)));
                const oldGlobalMax = parseInt(globalSlider.max);
                globalSlider.max = globalMaxIdx;
                
                // 最小値に変更があった場合のみ位置をリセット
                if (globalMaxIdx < oldGlobalMax) {{
                    globalSlider.value = 0;
                    globalLabel.textContent = 'Slice: 1/' + (globalMaxIdx + 1);
                }} else {{
                    // 位置はそのまま、ラベルのみ更新
                    const currentValue = parseInt(globalSlider.value);
                    globalLabel.textContent = 'Slice: ' + (currentValue + 1) + '/' + (globalMaxIdx + 1);
                }}
            }}
        }}
        
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
                const scaleX = img.naturalWidth / rect.width;
                const scaleY = img.naturalHeight / rect.height;
                let x = Math.round((e.clientX - rect.left) * scaleX);
                let y = Math.round((e.clientY - rect.top) * scaleY);
                
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
            infoPanels[idx].innerHTML = '画像サイズ: ' + img.naturalWidth + 'x' + img.naturalHeight + '<br>ROI: [' + 
                '<input type="number" id="roi-x-' + idx + '" value="' + x + '" style="width: 50px; text-align: center;" onchange="updateROIFromInput(' + idx + ')">,' + 
                '<input type="number" id="roi-y-' + idx + '" value="' + y + '" style="width: 50px; text-align: center;" onchange="updateROIFromInput(' + idx + ')">] ' + roiW + 'x' + roiH + 
                '<br>平均: ' + stats.mean + '<br>標準偏差: ' + stats.std + 
                '<span style="position: absolute; bottom: 2px; right: 2px; cursor: pointer; font-size: 12px; color: #666;" onclick="showMetadata(' + idx + ')" title="メタデータ表示">📋</span>';
        }}
        
        // メタデータ表示
        async function showMetadata(idx) {{
            try {{
                const metadata = await window.pywebview.api.get_metadata(idx, currentSlices[idx]);
                const popup = document.createElement('div');
                popup.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; border: 2px solid #ccc; border-radius: 8px; padding: 20px; max-width: 90vw; max-height: 90vh; overflow: hidden; z-index: 1000; box-shadow: 0 4px 20px rgba(0,0,0,0.3); display: flex; flex-direction: column;';
                
                const headerHtml = `
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; flex-shrink: 0;">
                        <h3 style="margin: 0;">DICOMメタデータ</h3>
                        <div style="display: flex; gap: 8px;">
                            <button onclick="this.parentElement.parentElement.parentElement.remove()" style="background: #f44336; color: white; border: none; border-radius: 4px; padding: 4px 8px; cursor: pointer;">×</button>
                        </div>
                    </div>
                    <div style="margin-bottom: 10px; flex-shrink: 0;">
                        <input type="text" id="metadata-search" placeholder="メタデータを検索..." style="width: 100%; padding: 6px; border: 1px solid #ccc; border-radius: 4px; font-size: 12px;">
                    </div>
                `;
                
                const contentHtml = `
                    <div id="metadata-content" style="flex: 1; overflow: auto; border: 1px solid #eee; border-radius: 4px; padding: 10px; background: #f9f9f9;">
                        <pre id="metadata-text" style="font-family: monospace; font-size: 12px; margin: 0; white-space: pre-wrap;">` + metadata + `</pre>
                    </div>
                `;
                
                popup.innerHTML = headerHtml + contentHtml;
                document.body.appendChild(popup);
                
                // 検索機能
                const searchInput = document.getElementById('metadata-search');
                const metadataText = document.getElementById('metadata-text');
                const originalText = metadataText.textContent;
                
                searchInput.addEventListener('input', function() {{
                    const searchTerm = this.value.toLowerCase();
                    if (searchTerm === '') {{
                        metadataText.textContent = originalText;
                        return;
                    }}
                    
                    const lines = originalText.split('\\n');
                    const filteredLines = lines.filter(line => 
                        line.toLowerCase().includes(searchTerm)
                    );
                    metadataText.textContent = filteredLines.join('\\n');
                }});
                
            }} catch (error) {{
                alert('メタデータの取得に失敗しました: ' + error);
            }}
        }}
        
        // 座標入力からROI更新
        function updateROIFromInput(idx) {{
            const xInput = document.getElementById('roi-x-' + idx);
            const yInput = document.getElementById('roi-y-' + idx);
            const x = parseInt(xInput.value);
            const y = parseInt(yInput.value);
            const img = imgs[idx];
            
            if (isNaN(x) || isNaN(y)) return;
            if (x < 0 || y < 0 || x + roiW > img.naturalWidth || y + roiH > img.naturalHeight) {{
                alert('ROIが画像範囲外です');
                return;
            }}
            
            if (syncMode) {{
                for (let j = 0; j < seriesCount; j++) roiCoords[j] = {{x: x, y: y}};
                redrawAllROIs();
                updateAllStats();
            }} else {{
                roiCoords[idx] = {{x: x, y: y}};
                drawROI(idx);
                updateStats(idx);
            }}
        }}
        
        function updateAllStats() {{
            for (let i = 0; i < seriesCount; i++) updateStats(i);
        }}
        
        // スライダーで画像切り替え
        for (let i = 0; i < seriesCount; i++) {{
            sliders[i].addEventListener('input', async function() {{
                const idx = sliders[i].value;
                currentSlices[i] = parseInt(idx);
                labels[i].textContent = 'Slice: ' + (parseInt(idx) + 1) + '/' + (parseInt(sliders[i].max) + 1);
                const filename = await window.pywebview.api.get_filename(i, idx);
                filenames[i].textContent = filename;
                const b64 = await window.pywebview.api.get_single_slice(i, idx);
                imgs[i].src = 'data:image/png;base64,' + b64;
                
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
        
        // グローバルスライダー
        const globalSlider = document.getElementById('global-slider');
        const globalLabel = document.getElementById('global-slice-label');
        globalSlider.addEventListener('input', async function() {{
            const idx = globalSlider.value;
            globalLabel.textContent = 'Slice: ' + (parseInt(idx) + 1) + '/' + (parseInt(globalSlider.max) + 1);
            for (let i = 0; i < seriesCount; i++) {{
                currentSlices[i] = parseInt(idx);
                sliders[i].value = idx;
                labels[i].textContent = 'Slice: ' + (parseInt(idx) + 1) + '/' + (parseInt(sliders[i].max) + 1);
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
        
        // キーボードショートカット
        window.addEventListener('keydown', function(e) {{
            if ((e.ctrlKey || e.metaKey) && e.key === 's') {{
                e.preventDefault();
                const row = getCurrentStats();
                historyData.push(row);
                renderHistoryTable();
            }}
        }});
        
        // 初期履歴テーブル表示
        renderHistoryTable();
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
            + f'<div class="dicom-square">'
            + f'<img class="dicom-img" id="dicom-img-{i}" src="data:image/png;base64,{b64}" />'
            + f'<canvas class="roi-canvas" id="roi-canvas-{i}"></canvas>'
            + f'</div>'
            + f'<div class="info-panel" id="info-panel-{i}"></div>'
            + f'<input type="range" class="slider" id="slider-{i}" min="0" max="{self.series_max_idx_list[i]}" value="0" />'
            + f'<div style="display: flex; gap: 5px; margin-top: 5px;"><div style="font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px;"><span id="slice-label-{i}" style="font-weight: bold;">Slice: 1/{self.series_max_idx_list[i] + 1}</span></div><div style="font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px;"><span id="filename-{i}">{self.file_names_list[i][0] if len(self.file_names_list[i]) > 0 else ""}</span></div></div>'
            + f'</div>'
            for i, b64 in enumerate(b64list)
        ])
        col_num = min(self.series_count, 4) if self.series_count > 1 else 1
        
        # フォルダ1のベース名リストを作成
        series_folder_base_names = [os.path.basename(f) for f in self.dicom_folders]
        html_content = HTML_TEMPLATE.format(
            series_blocks=series_blocks,
            global_max_idx=self.global_max_idx,
            global_max_idx_plus1=self.global_max_idx + 1,
            series_count=self.series_count,
            series_max_idx_list=self.series_max_idx_list,
            col_num=col_num,
            series_folder_base_names=series_folder_base_names
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

    def export_history_to_excel(self, history_data):
        try:
            # DataFrameの準備
            columns = ['No.']
            for i in range(self.series_count):
                columns.extend([f'Folder{i+1} Mean', f'Folder{i+1} Std Dev', f'Folder{i+1} Info'])
            
            data_rows = []
            for r_idx, row_data in enumerate(history_data):
                flat_row = [r_idx + 1]
                for series_stat in row_data:
                    flat_row.extend([series_stat.get('mean', ''), series_stat.get('std', ''), series_stat.get('info', '')])
                data_rows.append(flat_row)
            
            df = pd.DataFrame(data_rows, columns=columns)
            # info列以外をfloat型に変換
            for col in df.columns:
                if col.startswith('Folder') and not col.endswith('Info'):
                    df[col] = pd.to_numeric(df[col], errors='coerce')

            # --- 平均行の追加 ---
            # 2行空白行を挿入
            empty_row = [''] * len(columns)
            df_blank = pd.DataFrame([empty_row, empty_row], columns=columns)
            # 平均タイトル行（info列を除外）
            mean_title_no_info = [col for col in columns if not col.endswith('Info')]
            df_mean_title = pd.DataFrame([mean_title_no_info], columns=mean_title_no_info)
            # 平均値行（info列を除外）
            mean_values_no_info = ['']
            for i in range(self.series_count):
                mean_col = f'Folder{i+1} Mean'
                std_col = f'Folder{i+1} Std Dev'
                mean_val = df[mean_col].mean() if not df[mean_col].isnull().all() else ''
                std_val = df[std_col].mean() if not df[std_col].isnull().all() else ''
                mean_values_no_info.extend([
                    round(mean_val, 4) if mean_val != '' else '',
                    round(std_val, 4) if std_val != '' else ''
                ])
            df_mean_values = pd.DataFrame([mean_values_no_info], columns=mean_title_no_info)
            # 結合
            df_final = pd.concat([df, df_blank, df_mean_title, df_mean_values], ignore_index=True)
            
            # 現在の日時を取得し、ファイル名を生成
            now = datetime.datetime.now()
            # YYYYMMDD_HHMMSS の形式でフォーマット
            timestamp = now.strftime("%Y%m%d_%H%M%S") 
            file_name = f"ROI-history-{timestamp}.xlsx"
            
            # スクリプトが実行されているディレクトリに保存
            #script_dir = os.path.dirname(os.path.abspath(__file__))
            #file_path = os.path.join(script_dir, file_name)
            
            # ダウンロードフォルダに保存
            downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
            file_path = os.path.join(downloads_dir, file_name)

            df_final.to_excel(file_path, index=False, engine='openpyxl')
            
            return {'success': True, 'filePath': file_path}
        except Exception as e:
            # エラーの詳細をメッセージに含める
            return {'success': False, 'message': f"ファイルの保存中にエラーが発生しました: {str(e)}"}

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('Usage: python dicom_webview.py <dicom_folder1> [<dicom_folder2> ...]')
        exit(1)
    folders = sys.argv[1:]
    api = DicomWebApi(folders)
    html = api.get_init_html()
    window = webview.create_window('DICOM Web Viewer (Multi-Series)', html=html, js_api=api, width=1200, height=900)
    webview.start(debug=False)