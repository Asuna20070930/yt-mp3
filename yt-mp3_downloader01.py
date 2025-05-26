import os
import re
import json
import subprocess
from google.colab import drive, auth
import platform
import glob
import time
from mutagen.mp3 import MP3
from mutagen.id3 import ID3, TALB, TPE1, TIT2, TCON, TDRC
from datetime import datetime

import gspread
from google.auth import default
import pandas as pd

drive.mount('/content/drive', force_remount=True)

base_output_dir = "/content/drive/My Drive/MUSIC"
os.makedirs(base_output_dir, exist_ok=True)

category_folders = {
    "中文歌": os.path.join(base_output_dir, "中文歌"),
    "日文歌": os.path.join(base_output_dir, "日文歌"), 
    "英文歌": os.path.join(base_output_dir, "英文歌"),
    "純音樂": os.path.join(base_output_dir, "純音樂")
}

for folder_path in category_folders.values():
    os.makedirs(folder_path, exist_ok=True)

spreadsheet_name = "音樂資料庫"

gc = None
spreadsheet = None
worksheet = None

cache_dir = "/content/yt_dlp_cache"
os.makedirs(cache_dir, exist_ok=True)

def get_file_size(file_path):
    try:
        size_bytes = os.path.getsize(file_path)
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.2f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"
    except Exception as e:
        print(f"獲取檔案大小 {os.path.basename(file_path)} 時發生錯誤: {str(e)}")
        return "未知大小"

def get_mp3_metadata(file_path):
    try:
        audio = MP3(file_path)
        try:
            tags = ID3(file_path)
            title = str(tags.get('TIT2', '未知標題')).replace('TIT2(text=["', '').replace('"])', '')
            artist = str(tags.get('TPE1', '未知藝人')).replace('TPE1(text=["', '').replace('"])', '')
            album = str(tags.get('TALB', '未知專輯')).replace('TALB(text=["', '').replace('"])', '')
        except:
            title = os.path.basename(file_path).replace('.mp3', '')
            artist = '未知藝人'
            album = '未知專輯'

        duration = audio.info.length
        return {
            'title': title if title != '未知標題' else os.path.basename(file_path).replace('.mp3', ''),
            'artist': artist,
            'album': album,
            'duration': format_duration_seconds(duration)
        }
    except Exception as e:
        print(f"讀取MP3檔案 {os.path.basename(file_path)} 元數據時發生錯誤: {str(e)}")
        return {
            'title': os.path.basename(file_path).replace('.mp3', ''),
            'artist': '未知藝人',
            'album': '未知專輯',
            'duration': '未知時長'
        }

def format_duration_seconds(seconds):
    if seconds is None:
        return "未知時長"
    try:
        seconds = int(seconds)
        minutes, seconds = divmod(seconds, 60)
        hours, minutes = divmod(minutes, 60)
        if hours > 0:
            return f"{hours}:{minutes:02d}:{seconds:02d}"
        else:
            return f"{minutes:02d}:{seconds:02d}"
    except ValueError:
        return "格式錯誤"

def initialize_google_sheet():
    global gc, spreadsheet, worksheet, spreadsheet_name
    try:
        auth.authenticate_user()
        creds, _ = default()
        gc = gspread.authorize(creds)

        try:
            spreadsheet = gc.open(spreadsheet_name)
            print(f"已連接到現有的試算表: {spreadsheet_name}")
        except gspread.exceptions.SpreadsheetNotFound:
            spreadsheet = gc.create(spreadsheet_name)
            print(f"找不到試算表 '{spreadsheet_name}'，已自動創建新的試算表。")

        required_sheets = ["下載記錄", "中文歌", "日文歌", "英文歌", "純音樂"]
        existing_sheets = [ws.title for ws in spreadsheet.worksheets()]

        for sheet_name in required_sheets:
            if sheet_name not in existing_sheets:
                spreadsheet.add_worksheet(title=sheet_name, rows=1, cols=10)
                print(f"創建了新工作表: {sheet_name}")

            ws = spreadsheet.worksheet(sheet_name)
            current_headers = []
            if ws.row_count > 0:
                current_headers = ws.row_values(1)
 
            expected_headers = ["序號", "日期時間", "檔案名稱", "YouTube網址", "歌曲標題", "藝術家", "專輯", "時長", "文件大小", "類別"]

            if not current_headers or current_headers != expected_headers:
                ws.update('A1', [expected_headers], value_input_option='USER_ENTERED')
                print(f"已在工作表 '{sheet_name}' 中設定/更新標題行。")

                requests = []

                # 調整欄位寬度，包含新的「類別」欄位
                column_pixel_widths = [
                    ('A', 30), ('B', 150), ('C', 100), ('D', 300),
                    ('E', 300), ('F', 100), ('G', 100), ('H', 50), ('I', 700), ('J', 70)
                ]

                for col_letter, width_px in column_pixel_widths:
                    col_index = gspread.utils.a1_to_rowcol(col_letter + '1')[1] - 1
                    requests.append({
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": ws.id,
                                "dimension": "COLUMNS",
                                "startIndex": col_index,
                                "endIndex": col_index + 1
                            },
                            "properties": {
                                "pixelSize": width_px
                            },
                            "fields": "pixelSize"
                        }
                    })

                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1000,  
                            "startColumnIndex": 0,
                            "endColumnIndex": 10  # 更新為10欄
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "LEFT"
                            }
                        },
                        "fields": "userEnteredFormat.horizontalAlignment"
                    }
                })

                if requests:
                    try:
                        spreadsheet.batch_update({"requests": requests})
                        print(f"已調整工作表 '{sheet_name}' 的欄位寬度並設置所有欄位靠左對齊。")
                    except Exception as e_width:
                        print(f"調整工作表 '{sheet_name}' 欄位格式時發生錯誤: {e_width}")

        worksheet = spreadsheet.worksheet("下載記錄")
        print("預設使用「下載記錄」工作表")

        return True

    except Exception as e:
        print(f"初始化 Google Sheet 時發生錯誤: {str(e)}")
        gc = spreadsheet = worksheet = None
        return False

def select_song_category():
    """讓用戶選擇歌曲的類別"""
    print("\n請選擇歌曲類別:")
    print("1. 中文歌")
    print("2. 日文歌")
    print("3. 英文歌")
    print("4. 純音樂(OST)")
    print("0. 不分類")

    while True:
        choice = input("請選擇類別 (0-4): ").strip()
        if choice == "1":
            return "中文歌"
        elif choice == "2":
            return "日文歌"
        elif choice == "3":
            return "英文歌"
        elif choice == "4":
            return "純音樂"
        elif choice == "0":
            return None
        else:
            print("無效的選擇，請重新輸入。")

def get_output_directory(category):
    """根據類別返回對應的輸出目錄"""
    if category and category in category_folders:
        return category_folders[category]
    else:
        return base_output_dir

def add_record_to_google_sheet(filename, youtube_url, file_path=None, metadata=None, category=None):
    global worksheet, spreadsheet
    if not spreadsheet:
        print("Google Sheet 尚未初始化。無法新增記錄。")
        return False

    try:
        main_worksheet = spreadsheet.worksheet("下載記錄")
        all_values = main_worksheet.get_all_values()

        serial_number = 1
        if all_values:
             serial_number = len(all_values)
             expected_headers = ["序號", "日期時間", "檔案名稱", "YouTube網址", "歌曲標題", "藝術家", "專輯", "時長", "文件大小", "類別"]
             if all_values[0] != expected_headers:
                 serial_number = len(all_values) + 1

        if not all_values:
             serial_number = 1
             expected_headers = ["序號", "日期時間", "檔案名稱", "YouTube網址", "歌曲標題", "藝術家", "專輯", "時長", "文件大小", "類別"]
             try:
                main_worksheet.update('A1', [expected_headers], value_input_option='USER_ENTERED')
                print("偵測到工作表為空或標題行遺失，已自動補上標題行。")
             except Exception as e_header:
                print(f"嘗試補上標題行時發生錯誤: {e_header}")
                return False

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        file_size_str = "N/A"
        if file_path and os.path.exists(file_path):
            file_size_str = get_file_size(file_path)

        title = artist = album = duration_str = "未知"
        if metadata:
            title = metadata.get('title', '未知標題')
            artist = metadata.get('artist', '未知藝人')
            album = metadata.get('album', '未知專輯')
            duration_str = metadata.get('duration', '未知時長')
        elif file_path and os.path.exists(file_path) and filename.lower().endswith(".mp3"):
            temp_metadata = get_mp3_metadata(file_path)
            if temp_metadata:
                title = temp_metadata.get('title', '未知標題')
                artist = temp_metadata.get('artist', '未知藝人')
                album = temp_metadata.get('album', '未知專輯')
                duration_str = temp_metadata.get('duration', '未知時長')

        # 新增類別欄位
        new_row = [
            str(serial_number),
            now,
            filename,
            youtube_url,
            title,
            artist,
            album,
            duration_str,
            file_size_str,
            category if category else "未分類"
        ]

        main_worksheet.append_row(new_row, value_input_option='USER_ENTERED')
        print(f"記錄已新增至主要工作表「下載記錄」")

        # 如果有選擇類別，也加到對應的分類工作表中
        if category:
            try:
                category_worksheet = spreadsheet.worksheet(category)
                category_values = category_worksheet.get_all_values()

                category_serial = 1
                if category_values:
                    category_serial = len(category_values)
                    if category_values[0][0] == "序號":
                        category_serial = len(category_values)
                    else:
                        category_serial = len(category_values) + 1

                new_row[0] = str(category_serial)

                category_worksheet.append_row(new_row, value_input_option='USER_ENTERED')
                print(f"記錄也已新增至分類工作表「{category}」")
            except Exception as e:
                print(f"添加到分類工作表時發生錯誤: {str(e)}")

        try:
            for ws_name in ["下載記錄"] + ([category] if category else []):
                ws = spreadsheet.worksheet(ws_name)
                all_ws_values = ws.get_all_values()
                new_row_index = len(all_ws_values)

                requests = [{
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": new_row_index - 1,
                            "endRowIndex": new_row_index,
                            "startColumnIndex": 0,
                            "endColumnIndex": 10  # 更新為10欄
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "LEFT"
                            }
                        },
                        "fields": "userEnteredFormat.horizontalAlignment"
                    }
                }]

                spreadsheet.batch_update({"requests": requests})
        except Exception as e:
            print(f"設置新行格式時發生錯誤: {e}")

        return True
    except Exception as e:
        print(f"新增記錄到 Google Sheet 時發生錯誤: {str(e)}")
        return False

def setup_cookies():
    print("正在設置cookies以避免YouTube限制...")

    system = platform.system().lower()

    browsers = []
    if system == 'windows':
        browsers = ['chrome', 'edge', 'firefox', 'opera', 'brave', 'vivaldi', 'safari']
    elif system == 'darwin':
        browsers = ['chrome', 'firefox', 'safari', 'edge', 'opera', 'brave', 'vivaldi']
    else:
        browsers = ['chrome', 'firefox', 'opera', 'brave', 'edge', 'vivaldi']

    print("\n解決YouTube 429錯誤的方法:")
    print("1. 使用代理IP（減緩請求速率）")
    print("2. 手動設置User-Agent（模擬不同瀏覽器）")
    print("3. 手動輸入YouTube cookies（最有效但需要用戶操作）")

    choice = input("請選擇處理方式 (1/2/3): ")
    cookies_file = f"{cache_dir}/youtube_cookies.txt"
    user_agent = ""

    if choice == "1":
        return "--socket-timeout 30 --sleep-interval 5 --max-sleep-interval 10 --retries 10"
    elif choice == "2":
        ua_options = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.5 Safari/605.1.15",
            "Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/117.0"
        ]
        print("\n請選擇User-Agent:")
        for i, ua in enumerate(ua_options, 1):
            print(f"{i}. {ua}")

        ua_choice = input("請選擇 (1/2/3): ")
        try:
            user_agent = ua_options[int(ua_choice) - 1]
            return f"--user-agent \"{user_agent}\""
        except:
            print("無效選擇，使用默認User-Agent")
            return ""
    elif choice == "3":
        print("\n== 如何提供YouTube cookies ==")
        print("1. 在瀏覽器中登入YouTube")
        print("2. 安裝Cookie Editor瀏覽器擴展")
        print("3. 前往YouTube網站，開啟Cookie Editor")
        print("4. 點擊「Export」按鈕，選擇「Export as Netscape HTTP Cookie File」")
        print("5. 將內容複製並貼上到下方")
        print("\n請貼上從Cookie Editor擴展匯出的cookies內容（貼上後按Enter再按Ctrl+D結束輸入）:")

        cookies_content = []
        while True:
            try:
                line = input()
                cookies_content.append(line)
            except EOFError:
                break

        with open(cookies_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(cookies_content))

        if os.path.exists(cookies_file) and os.path.getsize(cookies_file) > 0:
            print("成功保存cookies!")
            return f"--cookies \"{cookies_file}\""
        else:
            print("cookies似乎為空，將嘗試不使用cookies")
            return ""
    else:
        print("無效選擇，不使用特殊參數")
        return ""

def setup_download_interval():
    """設定下載間隔時間"""
    print("\n=== 設定下載間隔 ===")
    print("為了避免 YouTube 429 錯誤，建議設定適當的下載間隔")

    min_interval = int(input("請設定最小下載間隔（秒）[預設30]: ") or "30")
    max_interval = int(input("請設定最大下載間隔（秒）[預設60]: ") or "60")

    if max_interval < min_interval:
        max_interval = min_interval
        print(f"最大間隔已調整為 {max_interval} 秒（不能小於最小間隔）")

    print(f"已設定下載間隔：{min_interval} - {max_interval} 秒")
    return min_interval, max_interval

def test_youtube_connection(extra_params):
    print("測試與YouTube的連接...")

    test_cmd = f'yt-dlp --dump-json "ytsearch1:test" {extra_params}'
    result = subprocess.run(test_cmd, shell=True, capture_output=True, text=True)

    if result.returncode == 0:
        print("✅ YouTube連接正常!")
        return True
    else:
        if "429" in result.stderr:
            print("❌ 仍然出現429錯誤，可能需要使用更有效的cookies或等待一段時間再試")
        elif "Unable to download webpage" in result.stderr:
            print("❌ 無法連接到YouTube，請檢查網絡連接")
        else:
            print(f"❌ 連接測試失敗: {result.stderr}")
        return False

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

def format_view_count(view_count):
    if view_count >= 1000000000:
        return f"{view_count/1000000000:.1f}B"
    elif view_count >= 1000000:
        return f"{view_count/1000000:.1f}M"
    elif view_count >= 1000:
        return f"{view_count/1000:.1f}K"
    else:
        return str(view_count)

def format_duration(seconds):
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)

    if hours > 0:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    else:
        return f"{minutes}:{seconds:02d}"

def check_duplicate_and_handle(filename, metadata, file_path, category=None):
    """
    檢查試算表中是否有重複檔案，並根據特定規則處理
    """
    global worksheet, spreadsheet
    if not worksheet:
        print("Google Sheet 尚未初始化，無法檢查重複。")
        return True, None

    try:
        all_data = worksheet.get_all_values()
        if len(all_data) <= 1:
            return True, None

        headers = all_data[0]
        data_rows = all_data[1:]

        filename_idx = headers.index("檔案名稱") if "檔案名稱" in headers else -1
        title_idx = headers.index("歌曲標題") if "歌曲標題" in headers else -1
        duration_idx = headers.index("時長") if "時長" in headers else -1
        filesize_idx = headers.index("文件大小") if "文件大小" in headers else -1

        if filename_idx == -1:
            print("找不到檔案名稱列，無法檢查重複。")
            return True, None

        duplicate_info = None
        exact_match = False
        name_match_only = False

        current_title = metadata.get('title', '')
        current_duration = metadata.get('duration', '')
        current_size = get_file_size(file_path)
        current_size_bytes = os.path.getsize(file_path)

        all_matches = []

        for i, row in enumerate(data_rows):
            if filename_idx < len(row) and row[filename_idx] == filename:
                row_info = {
                    "row_index": i + 2,
                    "filename": row[filename_idx],
                    "title": row[title_idx] if title_idx != -1 and title_idx < len(row) else "",
                    "duration": row[duration_idx] if duration_idx != -1 and duration_idx < len(row) else "",
                    "filesize": row[filesize_idx] if filesize_idx != -1 and filesize_idx < len(row) else ""
                }

                if (title_idx != -1 and row_info["title"] == current_title and
                    duration_idx != -1 and row_info["duration"] == current_duration):
                    exact_match = True
                    duplicate_info = row_info
                    all_matches.append(row_info)
                else:
                    name_match_only = True
                    if not duplicate_info:
                        duplicate_info = row_info
                    all_matches.append(row_info)

        if exact_match and duplicate_info:
            print(f"\n找到完全匹配的重複檔案（檔案名稱、標題、時長相同）!")
            print(f"現有檔案: {duplicate_info['filename']}, 大小: {duplicate_info['filesize']}")
            print(f"新下載檔案: {filename}, 大小: {current_size}")

            def parse_size(size_str):
                units = {"B": 1, "KB": 1024, "MB": 1024**2, "GB": 1024**3}
                size_str = size_str.strip()
                try:
                    if size_str.endswith("B"):
                        value = float(size_str.split()[0])
                        return value
                    else:
                        for unit, multiplier in units.items():
                            if unit in size_str:
                                value = float(size_str.split()[0])
                                return value * multiplier
                    return 0
                except:
                    return 0

            existing_size_bytes = parse_size(duplicate_info['filesize'])

            if current_size_bytes > existing_size_bytes:
                print("新檔案較大，將保留新檔案並更新記錄。")
                return True, [m['row_index'] for m in all_matches]
            else:
                choice = input("現有檔案較大或相同大小，是否仍要覆蓋? (y/n): ").lower()
                if choice == 'y':
                    print("用戶選擇覆蓋，將更新記錄。")
                    return True, [m['row_index'] for m in all_matches]
                else:
                    print("將刪除新下載的檔案。")
                    try:
                        os.remove(file_path)
                        print(f"已刪除新下載的檔案: {file_path}")
                        return False, None
                    except Exception as e:
                        print(f"刪除檔案失敗: {str(e)}")
                        return False, None

        elif name_match_only and duplicate_info:
            print(f"\n找到檔案名稱相同但內容不同的檔案!")
            print(f"現有檔案: {duplicate_info['filename']}, 標題: {duplicate_info['title']}, 時長: {duplicate_info['duration']}")
            print(f"新檔案: {filename}, 標題: {current_title}, 時長: {current_duration}")

            choice = input("是否覆蓋現有檔案? (y/n): ").lower()
            if choice == 'y':
                print("用戶選擇覆蓋，將更新記錄。")
                return True, [m['row_index'] for m in all_matches]
            else:
                print("用戶選擇不覆蓋，將創建新版本。")
                output_dir = get_output_directory(category)
                versions_dir = os.path.join(output_dir, f"{os.path.splitext(filename)[0]}_versions")
                os.makedirs(versions_dir, exist_ok=True)

                original_file = os.path.join(output_dir, duplicate_info['filename'])

                if os.path.exists(original_file):
                    new_original_path = os.path.join(versions_dir, f"original_{duplicate_info['filename']}")
                    try:
                        import shutil
                        shutil.copy2(original_file, new_original_path)
                        print(f"已複製原始檔案到: {new_original_path}")
                    except Exception as e:
                        print(f"複製原始檔案失敗: {str(e)}")
                else:
                    print(f"無法找到原始檔案: {original_file}，只會移動新檔案")

                new_file_path = os.path.join(versions_dir, f"new_{filename}")
                try:
                    import shutil
                    shutil.copy2(file_path, new_file_path)
                    print(f"已複製新檔案到: {new_file_path}")
                    return True, None
                except Exception as e:
                    print(f"複製新檔案失敗: {str(e)}")
                    return True, None

        return True, None

    except Exception as e:
        print(f"檢查重複時發生錯誤: {str(e)}")
        return True, None

def update_existing_record(row_index, filename, youtube_url, file_path=None, metadata=None, category=None):
    """更新已存在的記錄"""
    global worksheet, spreadsheet
    if not spreadsheet:
        print("Google Sheet 尚未初始化。無法更新記錄。")
        return False

    try:
        main_worksheet = spreadsheet.worksheet("下載記錄")

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        file_size_str = "N/A"
        if file_path and os.path.exists(file_path):
            file_size_str = get_file_size(file_path)

        title = artist = album = duration_str = "未知"
        if metadata:
            title = metadata.get('title', '未知標題')
            artist = metadata.get('artist', '未知藝人')
            album = metadata.get('album', '未知專輯')
            duration_str = metadata.get('duration', '未知時長')
        elif file_path and os.path.exists(file_path) and filename.lower().endswith(".mp3"):
            temp_metadata = get_mp3_metadata(file_path)
            if temp_metadata:
                title = temp_metadata.get('title', '未知標題')
                artist = temp_metadata.get('artist', '未知藝人')
                album = temp_metadata.get('album', '未知專輯')
                duration_str = temp_metadata.get('duration', '未知時長')

        # 修改：更新範圍包含類別欄位
        range_name = f'B{row_index}:J{row_index}'
        values = [[
            now,
            filename,
            youtube_url,
            title,
            artist,
            album,
            duration_str,
            file_size_str,
            category if category else "未分類"
        ]]

        main_worksheet.update(range_name, values, value_input_option='USER_ENTERED')
        print(f"已更新「下載記錄」工作表中的記錄第 {row_index} 行")

        # 修改：使用傳入的類別參數而非重新選擇
        if category:
            try:
                category_worksheet = spreadsheet.worksheet(category)
                all_category_records = category_worksheet.get_all_values()

                found_in_category = False
                category_row_index = -1

                for i, row in enumerate(all_category_records[1:], 2):
                    if len(row) >= 3 and row[2] == filename:
                        found_in_category = True
                        category_row_index = i
                        break

                if found_in_category:
                    range_name = f'B{category_row_index}:J{category_row_index}'
                    category_worksheet.update(range_name, values, value_input_option='USER_ENTERED')
                    print(f"已更新「{category}」工作表中的記錄第 {category_row_index} 行")
                else:
                    category_serial = len(all_category_records)
                    if all_category_records[0][0] == "序號":
                        category_serial = len(all_category_records)
                    else:
                        category_serial = len(all_category_records) + 1

                    new_row = [str(category_serial)] + values[0]
                    category_worksheet.append_row(new_row, value_input_option='USER_ENTERED')
                    print(f"記錄已新增至分類工作表「{category}」")

                    try:
                        new_row_index = len(all_category_records) + 1

                        requests = [{
                            "repeatCell": {
                                "range": {
                                    "sheetId": category_worksheet.id,
                                    "startRowIndex": new_row_index - 1,
                                    "endRowIndex": new_row_index,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 10  # 更新為10欄
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "horizontalAlignment": "LEFT"
                                    }
                                },
                                "fields": "userEnteredFormat.horizontalAlignment"
                            }
                        }]

                        spreadsheet.batch_update({"requests": requests})
                    except Exception as e:
                        print(f"設置新行格式時發生錯誤: {e}")
            except Exception as e:
                print(f"處理分類記錄時發生錯誤: {str(e)}")

        return True
    except Exception as e:
        print(f"更新記錄到 Google Sheet 時發生錯誤: {str(e)}")
        return False

def find_similar_files(metadata, current_file, output_dir):
    """
    查找與當前下載檔案的標題和時長都相同的檔案
    返回相似檔案的路徑和元數據列表
    """
    if not metadata or 'title' not in metadata or 'duration' not in metadata:
        return []

    similar_files = []
    target_title = metadata['title']
    target_duration = metadata['duration']

    def convert_duration_to_seconds(duration_str):
        parts = duration_str.split(':')
        if len(parts) == 2: 
            return int(parts[0]) * 60 + int(parts[1])
        elif len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        return 0

    target_duration_seconds = convert_duration_to_seconds(target_duration)
    
    for mp3_file in glob.glob(f"{output_dir}/*.mp3"):
        if mp3_file == current_file:
            continue

        file_metadata = get_mp3_metadata(mp3_file)
        if not file_metadata:
            continue

        if 'title' in file_metadata and file_metadata['title'] == target_title:
            if 'duration' in file_metadata:
                file_duration_seconds = convert_duration_to_seconds(file_metadata['duration'])
                if abs(file_duration_seconds - target_duration_seconds) <= 1:
                    similar_files.append((mp3_file, file_metadata))

    return similar_files

def download_as_mp3(youtube_url, extra_params=""):
    try:
        print(f"正在處理: {youtube_url}")

        category = select_song_category()
        output_dir = get_output_directory(category)

        output_template = f"{output_dir}/%(title)s.%(ext)s"

        command = f'yt-dlp {extra_params} --no-playlist -x --audio-format mp3 --audio-quality 0 --add-metadata --embed-metadata --no-embed-thumbnail --no-write-thumbnail -o "{output_template}" "{youtube_url}"'

        print("正在下載...")
        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"下載失敗: {result.stderr}")

            if "429" in result.stderr:
                print("\n遇到429錯誤 (Too Many Requests)，這意味著YouTube認為我們的下載行為是自動程序。")
                print("請嘗試以下解決方案:")
                print("1. 等待一段時間後再試")
                print("2. 使用cookies選項重新運行此程序")
                print("3. 使用VPN或代理改變IP地址")
                return False
            return False

        files = glob.glob(f"{output_dir}/*.mp3")
        if files:
            latest_file = max(files, key=os.path.getctime)
            filename = os.path.basename(latest_file)

            print("\n下載完成!")
            print(f"檔案名稱: {filename}")

            file_size = get_file_size(latest_file)
            print(f"文件大小: {file_size}")

            metadata = get_mp3_metadata(latest_file)
            if metadata:
                print(f"標題: {metadata['title']}")
                print(f"演出者: {metadata['artist']}")
                if metadata['album'] != '未知專輯':
                    print(f"專輯: {metadata['album']}")
                print(f"時長: {metadata['duration']}")

            similar_files = find_similar_files(metadata, latest_file, output_dir)
            if similar_files:
                print("\n⚠️ 發現有標題和時長都相同的歌曲存在！可能是重複下載。")
                print("\n=== 現有相似檔案 ===")
                for i, (file_path, file_meta) in enumerate(similar_files):
                    print(f"\n[檔案 {i+1}]")
                    print(f"檔名: {os.path.basename(file_path)}")
                    print(f"標題: {file_meta['title']}")
                    print(f"演出者: {file_meta['artist']}")
                    print(f"時長: {file_meta['duration']}")
                    if file_meta['album'] != '未知專輯':
                        print(f"專輯: {file_meta['album']}")
                    print(f"檔案大小: {get_file_size(file_path)}")
                    print(f"路徑: {file_path}")

                print("\n=== 剛下載的檔案 ===")
                print(f"檔名: {filename}")
                print(f"標題: {metadata['title']}")
                print(f"演出者: {metadata['artist']}")
                print(f"時長: {metadata['duration']}")
                if metadata['album'] != '未知專輯':
                    print(f"專輯: {metadata['album']}")
                print(f"檔案大小: {file_size}")
                print(f"路徑: {latest_file}")

                action = input("\n請選擇操作：\n1. 保留剛下載的檔案\n2. 刪除剛下載的檔案\n3. 保留全部\n請輸入選項 (1-3): ").strip()

                if action == "2":
                    try:
                        os.remove(latest_file)
                        print(f"已刪除剛下載的檔案: {filename}")
                        return True
                    except Exception as e:
                        print(f"刪除檔案時發生錯誤: {str(e)}")
                elif action == "1":
                    print("將保留剛下載的檔案，繼續處理...")
                    delete_old = input("是否要刪除之前的相似檔案? (y/n, 預設為n): ").lower().strip()
                    if delete_old == 'y':
                        for file_path, _ in similar_files:
                            try:
                                os.remove(file_path)
                                print(f"已刪除舊檔案: {os.path.basename(file_path)}")
                            except Exception as e:
                                print(f"刪除舊檔案時發生錯誤: {str(e)}")
                else:
                    print("將保留所有檔案，繼續處理...")

            new_name = input("請輸入新檔名（直接按Enter保持原檔名，無需.mp3副檔名）: ")
            if new_name:
                new_filepath = f"{output_dir}/{sanitize_filename(new_name)}.mp3"
                if os.path.exists(new_filepath) and new_filepath != latest_file:
                    print(f"\n⚠️ 警告：檔案「{os.path.basename(new_filepath)}」已存在!")
                    overwrite = input("是否覆蓋現有檔案? (y/n, 預設為n): ").lower().strip()
                    if overwrite == 'y':
                        try:
                            os.remove(new_filepath)
                            os.rename(latest_file, new_filepath)
                            filename = os.path.basename(new_filepath)
                            latest_file = new_filepath
                            print(f"已覆蓋並重新命名為: {filename}")
                        except Exception as e:
                            print(f"覆蓋檔案時發生錯誤: {str(e)}")
                    else:
                        print("將進入手動命名流程...")
                        while True:
                            manual_name = input("請輸入一個不重複的新檔名（無需.mp3副檔名）: ")
                            if not manual_name:
                                print("檔名不能為空，請重新輸入。")
                                continue

                            manual_filepath = f"{output_dir}/{sanitize_filename(manual_name)}.mp3"
                            if os.path.exists(manual_filepath) and manual_filepath != latest_file:
                                print(f"檔案「{os.path.basename(manual_filepath)}」也已存在，請再試一次。")
                            else:
                                try:
                                    os.rename(latest_file, manual_filepath)
                                    filename = os.path.basename(manual_filepath)
                                    latest_file = manual_filepath
                                    print(f"已重新命名為: {filename}")
                                    break
                                except Exception as e:
                                    print(f"重新命名檔案時發生錯誤: {str(e)}")
                                    break
                else:
                    try:
                        os.rename(latest_file, new_filepath)
                        filename = os.path.basename(new_filepath)
                        latest_file = new_filepath
                        print(f"已重新命名為: {filename}")
                    except Exception as e:
                        print(f"重新命名檔案時發生錯誤: {str(e)}")

            metadata = get_mp3_metadata(latest_file)

            should_continue, row_to_update = check_duplicate_and_handle(filename, metadata, latest_file, category)

            if should_continue:
                if row_to_update:
                    update_existing_record(row_to_update, filename, youtube_url, latest_file, metadata, category)
                else:
                    add_record_to_google_sheet(filename, youtube_url, latest_file, metadata, category)
                return True
            else:
                print("由於重複檢查結果，不添加新記錄。")
                return False
        else:
            print("找不到下載的檔案。")
            return False

    except Exception as e:
        print(f"下載時發生錯誤: {str(e)}")
        return False

def select_from_search_results(song_name, extra_params=""):
    try:
        print(f"正在搜尋: {song_name}")
        search_query = f"ytsearch15:{song_name}"
        command = f'yt-dlp {extra_params} --dump-json "{search_query}"'

        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"搜尋失敗: {result.stderr}")
            return None

        videos = []
        for line in result.stdout.strip().split('\n'):
            if line:
                try:
                    video_info = json.loads(line)
                    duration = video_info.get('duration', 0)

                    videos.append({
                        'title': video_info.get('title', '未知標題'),
                        'url': video_info.get('webpage_url', ''),
                        'channel': video_info.get('channel', '未知頻道'),
                        'view_count': video_info.get('view_count', 0),
                        'view_count_text': format_view_count(video_info.get('view_count', 0)),
                        'duration': duration,
                        'duration_text': format_duration(duration)
                    })
                except json.JSONDecodeError:
                    continue

        if not videos:
            print("找不到相關影片!")
            return None

        print("\n請從以下結果中選擇一個影片:")
        for i, video in enumerate(videos[:10], 1):
            print(f"{i}. {video['title']} - {video['channel']} ({video['view_count_text']} 觀看次數, {video['duration_text']})")

        selection = input("\n請輸入編號選擇影片 (輸入0取消): ")
        try:
            selection = int(selection)
            if selection > 0 and selection <= len(videos):
                selected_video = videos[selection-1]
                print(f"\n您選擇了: {selected_video['title']}")
                return selected_video['url']
            else:
                print("已取消選擇")
                return None
        except ValueError:
            print("無效的輸入，已取消選擇")
            return None

    except Exception as e:
        print(f"處理搜尋結果時發生錯誤: {str(e)}")
        return None

def download_by_url(extra_params=""):
    while True:
        youtube_url = input("請輸入 YouTube 影片網址 (輸入 '0' 退出): ")
        if youtube_url.lower() == '0':
            break

        if "youtube.com" in youtube_url or "youtu.be" in youtube_url:
            download_as_mp3(youtube_url, extra_params)
        else:
            print("請輸入有效的 YouTube 網址!")

    print("程式已結束")

def batch_download_urls(extra_params=""):
    urls = []
    print("請輸入多個 YouTube 網址 (每行一個，輸入空行結束):")

    while True:
        url = input()
        if not url:
            break
        if "youtube.com" in url or "youtu.be" in url:
            urls.append(url)
        else:
            print(f"警告: '{url}' 不像是 YouTube 網址，已略過")

    if not urls:
        print("沒有輸入有效的YouTube網址，操作已取消")
        return

    print(f"\n共有 {len(urls)} 個影片等待下載")

    confirm = input(f"確定要開始下載這 {len(urls)} 個影片嗎? (y/n, 預設 y): ").lower()
    if confirm and confirm != 'y' and confirm != '':
        print("操作已取消")
        return

    apply_rate_limit = input("是否啟用下載間隔以避免429錯誤? (y/n, 預設 y): ").lower() != 'n'
    rate_limit_args = "--limit-rate 500K --sleep-interval 10" if apply_rate_limit else ""
    if rate_limit_args and extra_params:
        extra_params = f"{extra_params} {rate_limit_args}"
    elif rate_limit_args:
        extra_params = rate_limit_args

    print(f"\n開始下載 {len(urls)} 個影片...")
    success_count = 0

    for i, url in enumerate(urls, 1):
        print(f"\n處理第 {i}/{len(urls)} 個影片:")

        if download_as_mp3(url, extra_params):
            success_count += 1
            print(f"進度：{success_count}/{len(urls)} 完成")
        else:
            print(f"下載失敗：{i}/{len(urls)}")

        if i < len(urls) and apply_rate_limit:
            pause_sec = 5
            print(f"等待 {pause_sec} 秒後下載下一個影片...")
            time.sleep(pause_sec)

    print(f"\n下載完成! 成功: {success_count}/{len(urls)}")

def download_song_with_manual_selection(extra_params=""):
    while True:
        song_name = input("請輸入歌曲名稱 (輸入 '0' 退出): ")
        if song_name == '0':
            break

        youtube_url = select_from_search_results(song_name, extra_params)
        if youtube_url:
            download_as_mp3(youtube_url, extra_params)
        else:
            print("搜尋失敗或已取消選擇，請重新輸入歌曲名稱")

    print("程式已結束")

print("進階 YouTube 音樂下載器 (優化版) - 解決429錯誤")
print("檔案將儲存至:", output_dir)

print("\n== 防止YouTube 429錯誤設置 ==")
extra_params = setup_cookies()

if extra_params:
    test_youtube_connection(extra_params)

print("\n== 初始化 Google Sheets ==")
if not initialize_google_sheet():
    print("警告：無法初始化 Google Sheets，下載記錄可能無法保存。")

print("\n== 請選擇下載模式 ==")
print("1. 輸入 YouTube 網址下載")
print("2. 批次下載多個 YouTube 網址")
print("3. 輸入歌曲名稱下載 (手動選擇)")

choice = input("請選擇模式 (1/2/3): ")
if choice == "1":
    download_by_url(extra_params)
elif choice == "2":
    batch_download_urls(extra_params)
elif choice == "3":
    download_song_with_manual_selection(extra_params)
else:
    print("無效的選擇，默認使用 YouTube 網址下載模式")
    download_by_url(extra_params)
