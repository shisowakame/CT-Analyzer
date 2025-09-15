const seriesCount = {series_count};
const seriesMaxIdxList = {series_max_idx_list};
let currentSlices = Array(seriesCount).fill(0);
const imgs = [], sliders = [], labels = [], canvases = [], infoPanels = [], filenames = [];
// フォルダ1のベース名リスト
const seriesFolderBaseNames = {series_folder_base_names};

// CSS変数を設定
const colNum = Math.min(seriesCount, 4) > 1 ? Math.min(seriesCount, 4) : 1;
document.documentElement.style.setProperty('--col-num', colNum);
document.documentElement.style.setProperty('--dicom-grid-width', '66.66vw');

// 履歴データ管理
let historyData = [];
const historyTableBlock = document.getElementById('history-table-block');
const saveBtn = document.getElementById('save-btn');
const resetBtn = document.getElementById('reset-btn');
const exportExcelBtn = document.getElementById('export-excel-btn');

// ROI色管理
let roiColor = '#ff0000'; // デフォルトは赤

// ROI座標管理（getCurrentStats関数で使用するため先に定義）
let roiCoords = Array(seriesCount).fill(null);

// ROIサイズ管理（getCurrentStats関数で使用するため先に定義）
let roiW = 10, roiH = 10;

for (let i = 0; i < seriesCount; i++) {
    imgs.push(document.getElementById('dicom-img-' + i));
    sliders.push(document.getElementById('slider-' + i));
    labels.push(document.getElementById('slice-label-' + i));
    canvases.push(document.getElementById('roi-canvas-' + i));
    infoPanels.push(document.getElementById('info-panel-' + i));
    filenames.push(document.getElementById('filename-' + i));
}

// 履歴テーブル描画関数
function renderHistoryTable() {
    let html = '<table class="history-table">';
    // ヘッダ行
    html += '<tr>';
    html += '<th style="width: 40px;"> </th>';
    for (let i = 0; i < seriesCount; i++) {
        html += '<th colspan="3">Folder' + (i + 1) + '</th>';
    }
    html += '<th style="width: 60px;"> </th>';
    html += '</tr>';
    // サブヘッダ行
    html += '<tr>';
    html += '<th></th>';
    for (let i = 0; i < seriesCount; i++) {
        html += '<th>平均</th><th>標準偏差</th><th>Info</th>';
    }
    html += '<th></th>';
    html += '</tr>';
    // データ行
    for (let r = 0; r < historyData.length; r++) {
        html += '<tr>';
        html += '<td style="text-align: center; font-weight: bold;">' + (r + 1) + '</td>';
        for (let i = 0; i < seriesCount; i++) {
            const meanVal = historyData[r][i] && historyData[r][i].mean ? historyData[r][i].mean : '';
            const stdVal = historyData[r][i] && historyData[r][i].std ? historyData[r][i].std : '';
            const infoVal = historyData[r][i] && historyData[r][i].info ? historyData[r][i].info : '';
            // infoセルは▽ボタンのみ表示、押すとポップアップ
            html += '<td>' + meanVal + '</td><td>' + stdVal + '</td>' +
                '<td style="text-align:center;"><button class="info-popup-btn" data-info="' + encodeURIComponent(infoVal) + '" style="background:none;border:none;cursor:pointer;font-size:16px;">'
                + '<i class="fa-solid fa-circle-info" style="color:#1976d2;"></i>'
                + '</button></td>';
        }
        html += '<td style="text-align: center;"><button onclick="deleteHistoryRow(' + r + ')" class="delete-btn">×</button></td>';
        html += '</tr>';
    }
    html += '</table>';
    historyTableBlock.innerHTML = html;

    // 平均行分離（Info列は平均行から除外）
    let avgHtml = '';
    if (historyData.length > 0) {
        avgHtml += '<table class="history-table" style="background: #f4f4f4;">';
        avgHtml += '<tr>';
        avgHtml += '<th style="width: 40px;">　</th>';
        for (let i = 0; i < seriesCount; i++) {
            avgHtml += '<th colspan="2">Folder' + (i + 1) + '</th>';
        }
        avgHtml += '<th style="width: 60px;">　</th>';
        avgHtml += '</tr>';
        avgHtml += '<tr>';
        avgHtml += '<th></th>';
        for (let i = 0; i < seriesCount; i++) {
            avgHtml += '<th>平均</th><th>標準偏差</th>';
        }
        avgHtml += '<th></th>';
        avgHtml += '</tr>';
        avgHtml += '<tr style="background: #f4f4f4;">';
        avgHtml += '<td style="font-weight: bold; text-align: center;">平均</td>';
        for (let i = 0; i < seriesCount; i++) {
            let meanSum = 0, stdSum = 0, cnt = 0;
            for (let r = 0; r < historyData.length; r++) {
                if (historyData[r][i] && historyData[r][i].mean) {
                    meanSum += parseFloat(historyData[r][i].mean);
                    cnt++;
                }
            }
            for (let r = 0; r < historyData.length; r++) {
                if (historyData[r][i] && historyData[r][i].std) {
                    stdSum += parseFloat(historyData[r][i].std);
                }
            }
            const avgMean = cnt ? (meanSum/cnt).toFixed(4) : '';
            const avgStd = cnt ? (stdSum/cnt).toFixed(4) : '';
            avgHtml += '<td style="font-weight: bold;">' + avgMean + '</td><td style="font-weight: bold;">' + avgStd + '</td>';
        }
        avgHtml += '<td></td>';
        avgHtml += '</tr>';
        avgHtml += '</table>';
    }
    document.getElementById('history-avg-row-block').innerHTML = avgHtml;

    // info-popup-btn のクリックイベントリスナーを再設定
    addInfoPopupEventListeners();
}

// 新しい情報ポップアップ表示関数
function showInfoPopup(infoText, event) {
    const popup = document.createElement('div');
    const rect = event.target.getBoundingClientRect();
    popup.style.cssText = `
        position: fixed;
        top: ${rect.top - 10}px;
        left: ${rect.left}px;
        transform: translate(-100%, -100%);
        background: white;
        border: 2px solid #ccc;
        border-radius: 8px;
        padding: 20px;
        max-width: 400px;
        max-height: 80vh;
        overflow-y: auto;
        z-index: 1000;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        display: flex;
        flex-direction: column;
    `;

    const headerHtml = `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; flex-shrink: 0;">
            <h3 style="margin: 0;">ROI情報</h3>
            <button onclick="this.parentElement.parentElement.remove()" style="background: #f44336; color: white; border: none; border-radius: 4px; padding: 4px 8px; cursor: pointer;">×</button>
        </div>
    `;

    const contentHtml = `
        <div style="flex: 1; overflow: auto; border: 1px solid #eee; border-radius: 4px; padding: 10px; background: #f9f9f9;">
            <pre style="font-family: monospace; font-size: 12px; margin: 0; white-space: pre-wrap;">` + infoText.replace('/ROIの基準座標:', '<br>/ROIの基準座標:') + `</pre>
        </div>
    `;

    popup.innerHTML = headerHtml + contentHtml;
    document.body.appendChild(popup);
}

// イベント委譲を使って info-popup-btn にイベントリスナーを追加
function addInfoPopupEventListeners() {
    historyTableBlock.querySelectorAll('.info-popup-btn').forEach(button => {
        button.onclick = function(e) {
            e.stopPropagation(); // 親要素へのイベント伝播を停止
            const info = decodeURIComponent(this.dataset.info);
            showInfoPopup(info, e);
        };
    });
}

// 履歴行削除関数
function deleteHistoryRow(rowIndex) {
    if (rowIndex >= 0 && rowIndex < historyData.length) {
        historyData.splice(rowIndex, 1);
        renderHistoryTable();
    }
}

// 現在の統計情報を取得
function getCurrentStats() {
    let row = [];
    for (let i = 0; i < seriesCount; i++) {
        let mean = '';
        let std = '';
        let info = '';
        if (roiCoords[i]) {
            const infoPanel = infoPanels[i];
            const infoText = infoPanel.innerText;
            const m = infoText.match(/平均: ([\d.-]+)/);
            const s = infoText.match(/標準偏差: ([\d.-]+)/);
            mean = m ? m[1] : '';
            std = s ? s[1] : '';
            // フォルダ名
            let folderName = '';
            const folderElem = document.querySelector('[onclick="showFolderSelector(' + i + ')"]');
            if (folderElem) {
                folderName = folderElem.textContent.replace(' ▼','');
            }
            // フォルダ名が空欄の場合はseriesFolderBaseNamesから補完
            if (!folderName) {
                folderName = seriesFolderBaseNames[i];
            }
            // ファイル名
            let fileName = filenames[i] ? filenames[i].textContent : '';
            // 座標
            let coord = '';
            if (roiCoords[i]) {
                coord = '(' + roiCoords[i].x + ',' + roiCoords[i].y + ')';
            }
            // ROIサイズ
            let roiSize = roiW + 'x' + roiH;
            info = folderName + '/' + fileName + '/ROIの基準座標:' + coord + '/ROIsize:' + roiSize;
        }
        row.push({mean: mean, std: std, info: info});
    }
    return row;
}

// 履歴関連イベントリスナー
saveBtn.addEventListener('click', function() {
    const row = getCurrentStats();
    historyData.push(row);
    renderHistoryTable();
});

resetBtn.addEventListener('click', function() {
    historyData = [];
    renderHistoryTable();
});

// Excelエクスポートボタンのイベントリスナー
exportExcelBtn.addEventListener('click', async function() {
    if (historyData.length === 0) {
        alert('エクスポートする履歴データがありません。');
        return;
    }
    try {
        const result = await window.pywebview.api.export_history_to_excel(historyData);
        if (result.success) {
            alert('履歴がExcelファイルとして保存されました:\n' + result.filePath);
        } else {
            alert('Excelファイルの保存に失敗しました:\n' + result.message);
        }
    } catch (error) {
        console.error('Excelエクスポートエラー:', error);
        alert('Excelファイルの保存中にエラーが発生しました。');
    }
});

// ROIサイズ管理
document.getElementById('roi-width').addEventListener('change', function() {
    let v = parseInt(this.value);
    if (isNaN(v) || v < 3) v = 3;
    if (v > 200) v = 200;
    roiW = v;
    this.value = v;
});
document.getElementById('roi-height').addEventListener('change', function() {
    let v = parseInt(this.value);
    if (isNaN(v) || v < 3) v = 3;
    if (v > 200) v = 200;
    roiH = v;
    this.value = v;
});
document.querySelectorAll('button.preset').forEach(btn => {
    btn.addEventListener('click', function() {
        roiW = parseInt(this.dataset.w);
        roiH = parseInt(this.dataset.h);
        document.getElementById('roi-width').value = roiW;
        document.getElementById('roi-height').value = roiH;
    });
});

// モード管理
let syncMode = true;

function updateModeButtons() {
    const syncBtn = document.getElementById('sync-mode-btn');
    const indepBtn = document.getElementById('indep-mode-btn');
    if (syncMode) {
        syncBtn.style.background = '#4CAF50';
        syncBtn.style.color = 'white';
        indepBtn.style.background = 'transparent';
        indepBtn.style.color = '#666';
    } else {
        syncBtn.style.background = 'transparent';
        syncBtn.style.color = '#666';
        indepBtn.style.background = '#4CAF50';
        indepBtn.style.color = 'white';
    }
}

function switchMode(newMode) {
    if (syncMode === newMode) return;

    syncMode = newMode;
    updateModeButtons();

    // 同期モードに変更した際は、自動的にROIを同期しない
    // ユーザーが実際にROIをクリックした際に同期される
}

document.getElementById('sync-mode-btn').addEventListener('click', function() {
    switchMode(true);
});

document.getElementById('indep-mode-btn').addEventListener('click', function() {
    switchMode(false);
});

// ROIリセットボタン
document.getElementById('reset-roi-btn').addEventListener('click', function() {
    roiCoords = Array(seriesCount).fill(null);
    redrawAllROIs();
    for (let i = 0; i < seriesCount; i++) {
        infoPanels[i].innerHTML = '';
    }
});

// 諧調揃えボタン状態
let matchContrastEnabled = false;
const matchContrastOnBtn = document.getElementById('match-contrast-on-btn');
const matchContrastOffBtn = document.getElementById('match-contrast-off-btn');

// ONボタンのイベントリスナー
matchContrastOnBtn.addEventListener('click', async function() {
    if (!matchContrastEnabled) {
        matchContrastEnabled = true;
        updateMatchContrastButtons();
        // Python側API呼び出し
        if (window.pywebview && window.pywebview.api && window.pywebview.api.set_match_contrast_enabled) {
            await window.pywebview.api.set_match_contrast_enabled(matchContrastEnabled);
        }
        // 画像再描画
        if (typeof redrawAllImages === 'function') {
            redrawAllImages();
        }
    }
});

// OFFボタンのイベントリスナー
matchContrastOffBtn.addEventListener('click', async function() {
    if (matchContrastEnabled) {
        matchContrastEnabled = false;
        updateMatchContrastButtons();
        // Python側API呼び出し
        if (window.pywebview && window.pywebview.api && window.pywebview.api.set_match_contrast_enabled) {
            await window.pywebview.api.set_match_contrast_enabled(matchContrastEnabled);
        }
        // 画像再描画
        if (typeof redrawAllImages === 'function') {
            redrawAllImages();
        }
    }
});

// ボタン状態更新関数
function updateMatchContrastButtons() {
    if (matchContrastEnabled) {
        matchContrastOnBtn.style.background = '#ff9800';
        matchContrastOnBtn.style.color = 'white';
        matchContrastOffBtn.style.background = 'transparent';
        matchContrastOffBtn.style.color = '#666';
    } else {
        matchContrastOnBtn.style.background = 'transparent';
        matchContrastOnBtn.style.color = '#666';
        matchContrastOffBtn.style.background = '#ff9800';
        matchContrastOffBtn.style.color = 'white';
    }
}

    // 表示画像ダウンロード機能
    const downloadDisplayBtn = document.getElementById('download-display-btn');
    downloadDisplayBtn.addEventListener('click', async function() {
        console.log('[DEBUG] ダウンロードボタンがクリックされました');
        try {
            console.log('[DEBUG] pywebview API確認:', !!window.pywebview, !!window.pywebview?.api, !!window.pywebview?.api?.get_display_images_for_download);
            
            if (window.pywebview && window.pywebview.api && window.pywebview.api.get_display_images_for_download) {
                console.log('[DEBUG] Python API呼び出し開始');
                console.log('[DEBUG] 現在のスライスインデックス:', currentSlices);
                const result = await window.pywebview.api.get_display_images_for_download(currentSlices);
                console.log('[DEBUG] Python API結果:', result);
                
                if (result.success) {
                    console.log('[DEBUG] ファイル保存成功:', result.file_path);
                    let message = '画像を保存しました:\n';
                    message += '結合画像: ' + result.filename + '\n';
                    if (result.individual_files && result.individual_files.length > 0) {
                        message += '個別画像:\n';
                        result.individual_files.forEach(filename => {
                            message += '  - ' + filename + '\n';
                        });
                    }
                    message += '保存先: ' + result.file_path.replace(/\\/g, '/').split('/').slice(-2).join('/');
                    alert(message);
                } else {
                    console.error('ダウンロードエラー:', result.error);
                    alert('画像の保存に失敗しました: ' + result.error);
                }
            } else {
                console.error('[DEBUG] pywebview APIが見つかりません');
                alert('pywebview APIが見つかりません');
            }
        } catch (error) {
            console.error('ダウンロード処理エラー:', error);
            alert('ダウンロード処理中にエラーが発生しました: ' + error.message);
        }
    });

    // ROI含む表示画像ダウンロード機能
    const downloadDisplayRoiBtn = document.getElementById('download-display-roi-btn');
    downloadDisplayRoiBtn.addEventListener('click', async function() {
        console.log('[DEBUG] ROI含むダウンロードボタンがクリックされました');
        try {
            console.log('[DEBUG] pywebview API確認:', !!window.pywebview, !!window.pywebview?.api, !!window.pywebview?.api?.get_display_images_with_roi_for_download);
            
            if (window.pywebview && window.pywebview.api && window.pywebview.api.get_display_images_with_roi_for_download) {
                console.log('[DEBUG] ROI座標を取得中...');
                
                // 現在のROI座標、色、サイズを取得
                const roiCoordsList = [];
                for (let i = 0; i < seriesCount; i++) {
                    if (roiCoords[i]) {
                        roiCoordsList.push({
                            x: roiCoords[i].x,
                            y: roiCoords[i].y,
                            color: roiColor,
                            width: roiW,
                            height: roiH
                        });
                    } else {
                        roiCoordsList.push(null);
                    }
                }
                
                console.log('[DEBUG] ROI座標リスト:', roiCoordsList);
                console.log('[DEBUG] 現在のスライスインデックス:', currentSlices);
                console.log('[DEBUG] Python API呼び出し開始');
                
                const result = await window.pywebview.api.get_display_images_with_roi_for_download(roiCoordsList, currentSlices);
                console.log('[DEBUG] Python API結果:', result);
                
                if (result.success) {
                    console.log('[DEBUG] ROI含むファイル保存成功:', result.file_path);
                    let message = 'ROI含む画像を保存しました:\n';
                    message += '結合画像: ' + result.filename + '\n';
                    if (result.individual_files && result.individual_files.length > 0) {
                        message += '個別画像:\n';
                        result.individual_files.forEach(filename => {
                            message += '  - ' + filename + '\n';
                        });
                    }
                    message += '保存先: ' + result.file_path.replace(/\\/g, '/').split('/').slice(-2).join('/');
                    alert(message);
                } else {
                    console.error('ROI含むダウンロードエラー:', result.error);
                    alert('ROI含む画像の保存に失敗しました: ' + result.error);
                }
            } else {
                console.error('[DEBUG] pywebview APIが見つかりません');
                alert('pywebview APIが見つかりません');
            }
        } catch (error) {
            console.error('ROI含むダウンロード処理エラー:', error);
            alert('ROI含むダウンロード処理中にエラーが発生しました: ' + error.message);
        }
    });

updateModeButtons();

// ROI色変更機能
const roiColorPicker = document.getElementById('roi-color-picker');

// ROI色変更のイベントリスナー
roiColorPicker.addEventListener('change', function() {
    roiColor = this.value;
    
    // 既存のROIを新しい色で再描画
    redrawAllROIs();
});

// フォルダ選択機能
async function showFolderSelector(seriesIdx) {
    const folderType = await window.pywebview.api.get_folder_type(seriesIdx);
    if (folderType !== 'folder2') return;

    const selector = document.getElementById('folder-selector-' + seriesIdx);
    if (selector.style.display === 'none') {
        const folderList = await window.pywebview.api.get_folder_list(seriesIdx);
        selector.innerHTML = '';
        folderList.forEach((folder, idx) => {
            const div = document.createElement('div');
            div.textContent = folder;
            div.style.padding = '4px 8px';
            div.style.cursor = 'pointer';
            div.style.borderBottom = '1px solid #eee';
            div.onmouseover = function() { this.style.backgroundColor = '#f0f0f0'; };
            div.onmouseout = function() { this.style.backgroundColor = 'white'; };
            div.onclick = function(e) {
                e.stopPropagation(); // イベント伝播を停止
                selectFolder(seriesIdx, idx);
                selector.style.display = 'none';
            };
            selector.appendChild(div);
        });
        selector.style.display = 'block';
    } else {
        selector.style.display = 'none';
    }
}

async function selectFolder(seriesIdx, folderIdx) {
    const result = await window.pywebview.api.switch_folder(seriesIdx, folderIdx);
    if (result.success) {
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
        const globalMaxIdx = Math.min(...Array.from({length: seriesCount}, (_, i) => parseInt(sliders[i].max)));
        const oldGlobalMax = parseInt(globalSlider.max);
        globalSlider.max = globalMaxIdx;
        if (globalMaxIdx < oldGlobalMax) {
            globalSlider.value = 0;
            globalLabel.textContent = 'Slice: 1/' + (globalMaxIdx + 1);
        } else {
            const currentValue = parseInt(globalSlider.value);
            globalLabel.textContent = 'Slice: ' + (currentValue + 1) + '/' + (globalMaxIdx + 1);
        }
        // 諧調揃えONなら全画像再描画
        if (result.match_contrast_enabled) {
            redrawAllImages();
        }
    }
}

document.addEventListener('click', function(e) {
    // フォルダ選択ボタンまたはフォルダ選択リストの外をクリックした場合
    if (!e.target.closest('.series-block') || 
        (e.target.closest('.series-block') && 
         !e.target.closest('[onclick*="showFolderSelector"]') && 
         !e.target.closest('[id*="folder-selector-"]'))) {
        for (let i = 0; i < seriesCount; i++) {
            const selector = document.getElementById('folder-selector-' + i);
            if (selector) {
                selector.style.display = 'none';
            }
        }
    }
});

// ROI描画・操作
function redrawAllROIs() {
    for (let i = 0; i < seriesCount; i++) drawROI(i);
}

function drawROI(idx) {
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
    ctx.strokeStyle = roiColor;
    ctx.lineWidth = 2;
    ctx.setLineDash([4,2]);
    ctx.globalAlpha = 0.7;
    ctx.strokeRect(x, y, w, h);
    ctx.globalAlpha = 0.2;
    ctx.fillStyle = roiColor;
    ctx.fillRect(x, y, w, h);
    ctx.globalAlpha = 1.0;
}

// canvasイベント
for (let i = 0; i < seriesCount; i++) {
    const canvas = canvases[i];
    const img = imgs[i];
    function resizeCanvas() {
        canvas.width = img.width;
        canvas.height = img.height;
        canvas.style.left = img.offsetLeft + 'px';
        canvas.style.top = img.offsetTop + 'px';
    }
    img.onload = resizeCanvas;
    window.addEventListener('resize', resizeCanvas);
    resizeCanvas();
    canvas.style.pointerEvents = 'auto';
    canvas.addEventListener('mousedown', function(e) {
        const rect = canvas.getBoundingClientRect();
        const scaleX = img.naturalWidth / rect.width;
        const scaleY = img.naturalHeight / rect.height;
        let x = Math.round((e.clientX - rect.left) * scaleX);
        let y = Math.round((e.clientY - rect.top) * scaleY);

        if (x < 0 || y < 0 || x + roiW > img.naturalWidth || y + roiH > img.naturalHeight) {
            showErrorPopup(canvas, e.clientX - rect.left, e.clientY - rect.top, 'ROIが画像範囲外です');
            return;
        }
        if (syncMode) {
            for (let j = 0; j < seriesCount; j++) roiCoords[j] = {x: x, y: y};
            redrawAllROIs();
            updateAllStats();
        } else {
            roiCoords[i] = {x: x, y: y};
            drawROI(i);
            updateStats(i);
        }
    });

    function showErrorPopup(canvas, x, y, msg) {
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
        setTimeout(() => {
            if (popup.parentNode) popup.parentNode.removeChild(popup);
        }, 1000);
    }
}

// ROI統計更新
async function updateStats(idx) {
    if (!roiCoords[idx]) return;
    const x = roiCoords[idx].x, y = roiCoords[idx].y;
    const stats = await window.pywebview.api.get_roi_stats(idx, currentSlices[idx], x, y, roiW, roiH);
    const img = imgs[idx];
    infoPanels[idx].innerHTML = '画像サイズ: ' + img.naturalWidth + 'x' + img.naturalHeight + '<br>ROI: [' +
        '<input type="number" id="roi-x-' + idx + '" value="' + x + '" style="width: 50px; text-align: center;" onchange="updateROIFromInput(' + idx + ')">,' +
        '<input type="number" id="roi-y-' + idx + '" value="' + y + '" style="width: 50px; text-align: center;" onchange="updateROIFromInput(' + idx + ')">] ' + roiW + 'x' + roiH +
        '<br>平均: ' + stats.mean + '<br>標準偏差: ' + stats.std;
}

// メタデータ表示
async function showMetadata(idx) {
    try {
        const metadata = await window.pywebview.api.get_metadata(idx, currentSlices[idx]);
        const popup = document.createElement('div');
        popup.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; border: 2px solid #ccc; border-radius: 8px; padding: 20px; max-width: 90vw; max-height: 90vh; overflow: hidden; z-index: 1000; box-shadow: 0 4px 20px rgba(0,0,0,0.3); display: flex; flex-direction: column;';

        const headerHtml = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; flex-shrink: 0;">
                <h3 style="margin: 0;">DICOMメタデータ</h3>
                <div style="display: flex; gap: 8px;">
                    <button onclick="this.parentElement.parentElement.parentElement.remove()" style="background: #f44336; color: white; border: none; border-radius: 4px; padding: 4px 8px; cursor: pointer;"><i class="fa-solid fa-xmark"></i></button>
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
        const originalText = metadataText.innerText;

        searchInput.addEventListener('input', function() {
            const searchTerm = this.value.toLowerCase();
            if (searchTerm === '') {
                metadataText.innerText = originalText;
                return;
            }
            // 改行コードを正規化してsplit
            const lines = originalText.replace(/\r\n|\r/g, '\n').split('\n');
            const filteredLines = lines.filter(line =>
                line.toLowerCase().includes(searchTerm)
            );
            metadataText.innerText = filteredLines.join('\n');
        });

    } catch (error) {
        alert('メタデータの取得に失敗しました: ' + error);
    }
}

// 座標入力からROI更新
function updateROIFromInput(idx) {
    const xInput = document.getElementById('roi-x-' + idx);
    const yInput = document.getElementById('roi-y-' + idx);
    const x = parseInt(xInput.value);
    const y = parseInt(yInput.value);
    const img = imgs[idx];

    if (isNaN(x) || isNaN(y)) return;
    if (x < 0 || y < 0 || x + roiW > img.naturalWidth || y + roiH > img.naturalHeight) {
        alert('ROIが画像範囲外です');
        return;
    }

    if (syncMode) {
        for (let j = 0; j < seriesCount; j++) roiCoords[j] = {x: x, y: y};
        redrawAllROIs();
        updateAllStats();
    } else {
        roiCoords[idx] = {x: x, y: y};
        drawROI(idx);
        updateStats(idx);
    }
}

function updateAllStats() {
    for (let i = 0; i < seriesCount; i++) updateStats(i);
}

// スライダーで画像切り替え
for (let i = 0; i < seriesCount; i++) {
    sliders[i].addEventListener('input', async function() {
        const idx = sliders[i].value;
        currentSlices[i] = parseInt(idx);
        labels[i].textContent = 'Slice: ' + (parseInt(idx) + 1) + '/' + (parseInt(sliders[i].max) + 1);
        const filename = await window.pywebview.api.get_filename(i, idx);
        filenames[i].textContent = filename;
        const b64 = await window.pywebview.api.get_single_slice(i, idx);
        imgs[i].src = 'data:image/png;base64,' + b64;

        setTimeout(() => {
            const img = imgs[i];
            const canvas = canvases[i];
            canvas.width = img.width;
            canvas.height = img.height;
            canvas.style.left = img.offsetLeft + 'px';
            canvas.style.top = img.offsetTop + 'px';
            drawROI(i);
            updateStats(i);
        }, 50);
    });
}

// グローバルスライダー
const globalSlider = document.getElementById('global-slider');
const globalLabel = document.getElementById('global-slice-label');
globalSlider.addEventListener('input', async function() {
    const idx = globalSlider.value;
    globalLabel.textContent = 'Slice: ' + (parseInt(idx) + 1) + '/' + (parseInt(globalSlider.max) + 1);
    for (let i = 0; i < seriesCount; i++) {
        currentSlices[i] = parseInt(idx);
        sliders[i].value = idx;
        labels[i].textContent = 'Slice: ' + (parseInt(idx) + 1) + '/' + (parseInt(sliders[i].max) + 1);
        const filename = await window.pywebview.api.get_filename(i, idx);
        filenames[i].textContent = filename;
        const img = imgs[i];
        const canvas = canvases[i];
        img.onload = function() {
            canvas.width = img.width;
            canvas.height = img.height;
            canvas.style.left = img.offsetLeft + 'px';
            canvas.style.top = img.offsetTop + 'px';
            drawROI(i);
            updateStats(i);
        };
    }
    const b64list = await window.pywebview.api.get_slice(idx);
    for (let i = 0; i < seriesCount; i++) {
        imgs[i].src = 'data:image/png;base64,' + b64list[i];
    }
});

// キーボードショートカット
window.addEventListener('keydown', function(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        const row = getCurrentStats();
        historyData.push(row);
        renderHistoryTable();
    }
});

// redrawAllImages: 諧調揃えON/OFFやフォルダ切替時に全画像を再描画
function redrawAllImages() {
    for (let i = 0; i < seriesCount; i++) {
        (async function(idx) {
            const sliceIdx = currentSlices[idx];
            const b64 = await window.pywebview.api.get_single_slice(idx, sliceIdx);
            imgs[idx].src = 'data:image/png;base64,' + b64;
            setTimeout(() => {
                const img = imgs[idx];
                const canvas = canvases[idx];
                canvas.width = img.width;
                canvas.height = img.height;
                canvas.style.left = img.offsetLeft + 'px';
                canvas.style.top = img.offsetTop + 'px';
                drawROI(idx);
                updateStats(idx);
            }, 50);
        })(i);
    }
}

// 初期履歴テーブル表示
renderHistoryTable();
