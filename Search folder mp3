import os
import glob
import pandas as pd
from mutagen.mp3 import MP3
from mutagen.id3 import ID3
from datetime import datetime
from google.colab import drive
import gspread
from google.colab import auth
from google.auth import default

# 掛載 Google Drive
drive.mount('/content/drive')

# 設定路徑
music_dir = "/content/drive/My Drive/MUSIC"
spreadsheet_name = "音樂資料庫"  # 您的 Google Sheet 名稱

# 定義類別資料夾
category_folders = {
    "中文歌": os.path.join(music_dir, "中文歌"),
    "日文歌": os.path.join(music_dir, "日文歌"),
    "英文歌": os.path.join(music_dir, "英文歌"),
    "純音樂": os.path.join(music_dir, "純音樂")
}

# 需要安裝的套件
!pip install gspread google-auth

# 獲取文件大小的函數
def get_file_size(file_path):
    """獲取文件大小並轉換為可讀格式"""
    try:
        size_bytes = os.path.getsize(file_path)
        # 轉換為可讀格式
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.2f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"
    except Exception as e:
        print(f"獲取文件大小時發生錯誤: {str(e)}")
        return "未知大小"

# 格式化時長
def format_duration(seconds):
    """將秒數轉換為時分秒格式"""
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)

    if hours > 0:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    else:
        return f"{minutes}:{seconds:02d}"

# 確定檔案類別
def get_file_category(file_path):
    """根據檔案路徑確定檔案類別"""
    for category, folder_path in category_folders.items():
        if file_path.startswith(folder_path):
            return category
    return "其他"  # 不在任何定義類別中的檔案

# 獲取MP3文件的元數據
def get_mp3_metadata(file_path):
    """獲取MP3文件的元數據"""
    try:
        # 使用mutagen讀取MP3標籤
        audio = MP3(file_path)

        # 嘗試獲取ID3標籤
        tags = None
        try:
            tags = ID3(file_path)
        except:
            pass

        # 準備元數據
        metadata = {
            'filename': os.path.basename(file_path),
            'filepath': file_path,
            'category': get_file_category(file_path),  # 新增類別欄位
            'filesize': get_file_size(file_path),
            'duration': format_duration(int(audio.info.length)),
            'title': '',
            'artist': '',
            'album': '',
            'year': ''
        }

        # 如果有ID3標籤，則獲取相關信息
        if tags:
            metadata['title'] = str(tags.get('TIT2', '')) if 'TIT2' in tags else ''
            metadata['artist'] = str(tags.get('TPE1', '')) if 'TPE1' in tags else ''
            metadata['album'] = str(tags.get('TALB', '')) if 'TALB' in tags else ''
            metadata['year'] = str(tags.get('TDRC', '')) if 'TDRC' in tags else ''

        return metadata
    except Exception as e:
        print(f"讀取檔案 {os.path.basename(file_path)} 時發生錯誤: {str(e)}")
        return {
            'filename': os.path.basename(file_path),
            'filepath': file_path,
            'category': get_file_category(file_path),
            'filesize': get_file_size(file_path),
            'duration': '未知',
            'title': '',
            'artist': '',
            'album': '',
            'year': ''
        }

def display_category_statistics(all_metadata):
    """顯示各類別檔案統計"""
    print("\n" + "="*80)
    print("📁 各類別檔案統計")
    print("="*80)
    
    # 統計各類別檔案數量
    category_counts = {}
    for metadata in all_metadata:
        category = metadata['category']
        category_counts[category] = category_counts.get(category, 0) + 1
    
    # 顯示統計結果
    total_files = len(all_metadata)
    for category in ["中文歌", "日文歌", "英文歌", "純音樂", "其他"]:
        count = category_counts.get(category, 0)
        percentage = (count / total_files * 100) if total_files > 0 else 0
        print(f"🎵 {category}: {count:4d} 個檔案 ({percentage:5.1f}%)")
    
    print("-" * 40)
    print(f"📊 總計: {total_files:4d} 個檔案 (100.0%)")
    
    return category_counts

def display_comparison_results(scanned_files, existing_files, new_files, category_stats=None):
    """顯示對比分析結果"""
    print("\n" + "="*80)
    print("📊 檔案對比分析結果")
    print("="*80)

    # 基本統計
    print(f"📁 本次掃描檔案數量: {len(scanned_files)}")
    print(f"📋 試算表已有記錄數量: {len(existing_files)}")
    print(f"🆕 本次新增檔案數量: {len(new_files)}")
    print(f"🔄 重複檔案數量: {len(scanned_files) - len(new_files)}")

    # 找出只在掃描中存在的檔案（新檔案）
    scanned_set = set(scanned_files)
    existing_set = set(existing_files)

    only_in_scan = scanned_set - existing_set  # 只在本次掃描中有的
    only_in_sheet = existing_set - scanned_set  # 只在試算表中有的（可能已刪除）
    common_files = scanned_set & existing_set   # 共同存在的

    print("\n" + "-"*60)
    print("🆕 新發現的檔案（需要新增到試算表）:")
    print("-"*60)
    if only_in_scan:
        for i, filename in enumerate(sorted(only_in_scan), 1):
            print(f"{i:3d}. {filename}")
    else:
        print("   沒有新檔案")

    print(f"\n總計: {len(only_in_scan)} 個新檔案")

    print("\n" + "-"*60)
    print("⚠️  可能已刪除的檔案（在試算表中但掃描時未找到）:")
    print("-"*60)
    if only_in_sheet:
        for i, filename in enumerate(sorted(only_in_sheet), 1):
            print(f"{i:3d}. {filename}")
        print(f"\n⚠️  注意: 這些檔案可能已從 {music_dir} 中刪除")
        print("   建議檢查是否需要從試算表中移除這些記錄")
    else:
        print("   沒有遺失的檔案")

    print(f"\n總計: {len(only_in_sheet)} 個可能已刪除的檔案")

    print("\n" + "-"*60)
    print("✅ 已存在的檔案（無需重複新增）:")
    print("-"*60)
    print(f"   共 {len(common_files)} 個檔案已在試算表中")

    # 如果重複檔案不多，可以列出來
    if len(common_files) <= 10:
        for i, filename in enumerate(sorted(common_files), 1):
            print(f"{i:3d}. {filename}")
    elif len(common_files) > 10:
        print("   (檔案數量較多，僅顯示前10個)")
        for i, filename in enumerate(sorted(list(common_files)[:10]), 1):
            print(f"{i:3d}. {filename}")
        print(f"   ... 還有 {len(common_files) - 10} 個檔案")

    print("\n" + "="*80)
    print("📈 統計摘要:")
    print("="*80)
    print(f"   🔍 掃描檔案總數: {len(scanned_files)}")
    print(f"   📋 試算表記錄數: {len(existing_files)}")
    print(f"   🆕 新增檔案數: {len(only_in_scan)}")
    print(f"   ⚠️  可能遺失檔案數: {len(only_in_sheet)}")
    print(f"   ✅ 已存在檔案數: {len(common_files)}")
    print("="*80)

def main():
    print("=== 音樂文件掃描與 Google 試算表匯入工具 ===")
    print(f"音樂目錄: {music_dir}")
    print(f"Google 試算表名稱: {spreadsheet_name}")
    print()

    # 檢查各類別資料夾是否存在
    print("檢查類別資料夾:")
    for category, folder_path in category_folders.items():
        if os.path.exists(folder_path):
            print(f"✅ {category}: {folder_path}")
        else:
            print(f"❌ {category}: {folder_path} (資料夾不存在)")

    # 掃描所有MP3文件
    print("\n正在掃描MP3文件...")
    mp3_files = glob.glob(os.path.join(music_dir, "**", "*.mp3"), recursive=True)
    print(f"找到 {len(mp3_files)} 個MP3文件")

    if len(mp3_files) == 0:
        print("沒有找到MP3文件，程式結束")
        return

    # 收集所有文件的元數據
    print("正在收集文件元數據...")
    all_metadata = []
    scanned_filenames = []  # 儲存本次掃描的檔案名稱

    for i, file_path in enumerate(mp3_files, 1):
        print(f"處理 ({i}/{len(mp3_files)}): {os.path.basename(file_path)}")
        metadata = get_mp3_metadata(file_path)
        all_metadata.append(metadata)
        scanned_filenames.append(metadata['filename'])

    # 顯示各類別統計
    category_stats = display_category_statistics(all_metadata)

    # 創建 DataFrame
    df = pd.DataFrame(all_metadata)

    # 連接到 Google Sheets
    print("\n正在連接 Google Sheets...")

    try:
        # 使用 google-auth 的方式進行認證
        auth.authenticate_user()
        creds, _ = default()
        gc = gspread.authorize(creds)

        # 嘗試打開現有試算表，如果不存在則創建
        try:
            # 檢查試算表是否存在
            try:
                spreadsheet = gc.open(spreadsheet_name)
                print(f"已連接到現有的試算表: {spreadsheet_name}")

                # 選擇要操作的工作表
                worksheet_name = input("請輸入要操作的工作表名稱 (默認為'確認已下載'): ") or "確認已下載"

                try:
                    worksheet = spreadsheet.worksheet(worksheet_name)
                    print(f"已選擇工作表: {worksheet_name}")
                except:
                    # 如果工作表不存在，創建一個新的
                    worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=20)
                    print(f"已創建新的工作表: {worksheet_name}")

                    # 設置標題行（類別欄位放最後）
                    worksheet.update('A1:J1', [["序號", "日期時間", "檔案名稱", "YouTube網址",
                                               "歌曲標題", "藝術家", "專輯", "時長", "文件大小", "類別"]])
                    print("已設置標題行")

            except:
                # 創建新的試算表
                spreadsheet = gc.create(spreadsheet_name)
                print(f"已創建新的試算表: {spreadsheet_name}")

                # 使用預設的第一個工作表
                worksheet = spreadsheet.sheet1
                worksheet.update_title("Sheet1")

                # 設置標題行（類別欄位放最後）
                worksheet.update('A1:J1', [["序號", "日期時間", "檔案名稱", "YouTube網址",
                                           "歌曲標題", "藝術家", "專輯", "時長", "文件大小", "類別"]])
                print("已設置標題行")

            # 檢查工作表中已有的資料
            existing_data = worksheet.get_all_values()

            # 確認是否有標題行
            if not existing_data:
                # 如果試算表是空的，添加標題行（類別欄位放最後）
                worksheet.update('A1:J1', [["序號", "日期時間", "檔案名稱", "YouTube網址",
                                           "歌曲標題", "藝術家", "專輯", "時長", "文件大小", "類別"]])
                existing_data = [["序號", "日期時間", "檔案名稱", "YouTube網址",
                                 "歌曲標題", "藝術家", "專輯", "時長", "文件大小", "類別"]]

            # 獲取現有檔案名稱列表
            existing_filenames = [row[2] for row in existing_data[1:] if len(row) > 2]  # 跳過標題行

            # 準備新資料
            new_rows = []
            new_filenames = []  # 儲存本次實際新增的檔案名稱
            current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            next_row_num = len(existing_data)

            for i, metadata in enumerate(all_metadata):
                # 檢查檔案是否已存在
                if metadata['filename'] in existing_filenames:
                    continue

                # 序號從最後一行+1開始（類別欄位放最後）
                row_data = [
                    str(next_row_num + len(new_rows)),  # 序號
                    current_datetime,  # 日期時間
                    metadata['filename'],  # 檔案名稱
                    "",  # YouTube網址 (留空，讓使用者自行填寫)
                    metadata['title'],  # 歌曲標題
                    metadata['artist'],  # 藝術家
                    metadata['album'],  # 專輯
                    metadata['duration'],  # 時長
                    metadata['filesize'],  # 文件大小
                    metadata['category']  # 類別（放最後）
                ]
                new_rows.append(row_data)
                new_filenames.append(metadata['filename'])

            # 如果有新資料，添加到試算表
            if new_rows:
                # 添加到工作表
                start_row = len(existing_data) + 1  # 從現有資料之後開始
                end_row = start_row + len(new_rows) - 1
                range_str = f'A{start_row}:J{end_row}'  # 更新為J欄（類別欄位放最後）

                worksheet.update(range_str, new_rows)
                print(f"已成功添加 {len(new_rows)} 筆新資料")
            else:
                print("沒有發現新的MP3檔案需要添加")

            # 顯示詳細的對比分析結果
            display_comparison_results(scanned_filenames, existing_filenames, new_filenames, category_stats)

            # 提供試算表的連結
            print(f"\n🔗 試算表連結: {spreadsheet.url}")

        except Exception as e:
            print(f"處理Google試算表時發生錯誤: {str(e)}")

    except Exception as e:
        print(f"Google認證失敗: {str(e)}")
        print("建議確認您已授權足夠的權限給Colab訪問您的Google帳戶")

    print("\n✅ 程序執行完成!")

if __name__ == "__main__":
    main()
