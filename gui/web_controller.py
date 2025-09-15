import os

class WebController:
    def get_init_html(self, html_template):
        # b64list = [self.get_png_b64(images[0]) for images in self.images_list]
        b64list = [self.get_png_b64(original_images[0]) for original_images in self.original_images_list]
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
            + f'<div style="display: flex; gap: 5px; margin-top: 5px; align-items: center;">'
            + f'<div style="font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px; display: flex; align-items: center;">'
            + f'<span id="filename-{i}" style="display: inline-block; vertical-align: middle;">{self.file_names_list[i][0] if len(self.file_names_list[i]) > 0 else ""}</span>'
            + f'<button onclick="showMetadata({i})" title="メタデータ表示" style="background: none; border: none; padding: 0 0 0 6px; margin: 0; cursor: pointer; vertical-align: middle; height: 100%; display: flex; align-items: center;">'
            + f'<i class="fa-solid fa-circle-info" style="color: #1976d2; font-size: 1.2em;"></i>'
            + f'</button>'
            + f'</div>'
            + f'<div style="font-size: 12px; color: #222; background: #f4f4f4; border-radius: 4px; padding: 4px 8px;"><span id="slice-label-{i}" style="font-weight: bold;">Slice: 1/{self.series_max_idx_list[i] + 1}</span></div>'
            + f'</div>'
            + f'</div>'
            for i, b64 in enumerate(b64list)
        ])
        col_num = min(self.series_count, 4) if self.series_count > 1 else 1
        
        # フォルダ1のベース名リストを作成
        series_folder_base_names = [os.path.basename(f) for f in self.dicom_folders]
        html_content = html_template.format(
            series_blocks=series_blocks,
            global_max_idx=self.global_max_idx,
            global_max_idx_plus1=self.global_max_idx + 1,
            series_count=self.series_count,
            series_max_idx_list=self.series_max_idx_list,
            col_num=col_num,
            series_folder_base_names=series_folder_base_names
        )
        
        return html_content
    # JSから呼び出すAPI: 諧調揃え状態のON/OFF切替
    def set_match_contrast_enabled(self, enabled):
        self.match_contrast_enabled = bool(enabled)
        return {'success': True, 'enabled': self.match_contrast_enabled}
