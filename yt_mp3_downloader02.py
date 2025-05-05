import os
import re
import json
import subprocess
import platform
import glob
import sys
import time
import ctypes
import shutil
import tempfile
import zipfile
from urllib.request import urlretrieve

def is_admin():
    """檢查程式是否以管理員權限運行"""
    try:
        if platform.system() == 'Windows':
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        else:
            return os.geteuid() == 0  # Unix-like
    except:
        return False

def check_ffmpeg():
    """檢查系統中是否存在FFmpeg"""
    # 首先檢查PATH中是否有FFmpeg
    ffmpeg_cmd = "ffmpeg.exe" if platform.system() == "Windows" else "ffmpeg"
    
    if shutil.which(ffmpeg_cmd):
        return shutil.which(ffmpeg_cmd)
    
    # 檢查常見的安裝位置
    if platform.system() == "Windows":
        # 檢查用戶下載的FFmpeg目錄
        ffmpeg_paths = [
            os.path.join(os.path.expanduser("~"), "Downloads", "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join(os.path.expanduser("~"), "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join("C:\\", "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join("D:\\", "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "ffmpeg", "bin", "ffmpeg.exe")
        ]
        
        for path in ffmpeg_paths:
            if os.path.exists(path):
                return path
    
    return None

def download_ffmpeg():
    """下載並配置FFmpeg"""
    temp_dir = tempfile.gettempdir()
    ffmpeg_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ffmpeg")
    os.makedirs(ffmpeg_dir, exist_ok=True)
    
    # 根據操作系統選擇下載連結
    if platform.system() == "Windows":
        # Windows 64位 FFmpeg下載連結
        ffmpeg_url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
        zip_path = os.path.join(temp_dir, "ffmpeg.zip")
        
        print(f"正在下載FFmpeg... (這可能需要幾分鐘)")
        try:
            urlretrieve(ffmpeg_url, zip_path)
            
            print("正在解壓FFmpeg...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 找到解壓後的ffmpeg.exe
            extracted_dirs = [d for d in os.listdir(temp_dir) if d.startswith("ffmpeg-master")]
            if not extracted_dirs:
                print("解壓後找不到FFmpeg目錄")
                return None
                
            extracted_dir = os.path.join(temp_dir, extracted_dirs[0])
            bin_dir = os.path.join(extracted_dir, "bin")
            
            # 創建ffmpeg目錄
            bin_target = os.path.join(ffmpeg_dir, "bin")
            os.makedirs(bin_target, exist_ok=True)
            
            # 複製ffmpeg.exe, ffprobe.exe和ffplay.exe
            for file in ["ffmpeg.exe", "ffprobe.exe", "ffplay.exe"]:
                src = os.path.join(bin_dir, file)
                dst = os.path.join(bin_target, file)
                if os.path.exists(src):
                    shutil.copy2(src, dst)
                    print(f"已複製 {file} 到 {dst}")
            
            # 清理臨時文件
            os.remove(zip_path)
            
            # 將ffmpeg/bin添加到PATH環境變量
            ffmpeg_bin_path = os.path.abspath(bin_target)
            os.environ["PATH"] = ffmpeg_bin_path + os.pathsep + os.environ["PATH"]
            
            # 測試ffmpeg是否可用
            ffmpeg_exe = os.path.join(ffmpeg_bin_path, "ffmpeg.exe")
            if os.path.exists(ffmpeg_exe):
                print(f"FFmpeg已安裝至: {ffmpeg_exe}")
                return ffmpeg_exe
            else:
                print("安裝FFmpeg失敗，找不到ffmpeg.exe")
                return None
        except Exception as e:
            print(f"下載FFmpeg時發生錯誤: {str(e)}")
            return None
    else:
        print("非Windows系統，請使用包管理器安裝FFmpeg")
        return None

def install_required_packages():
    """安裝必要的套件"""
    print("正在檢查並安裝必要的套件...")
    try:
        # 嘗試導入yt-dlp，如果不存在則安裝
        try:
            import yt_dlp
            print("✅ yt-dlp 已安裝")
        except ImportError:
            print("正在安裝 yt-dlp...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "yt-dlp"])
            print("✅ yt-dlp 安裝完成")
        
        # 檢查 FFmpeg 是否存在
        print("正在檢查 FFmpeg...")
        ffmpeg_path = check_ffmpeg()
        if ffmpeg_path:
            print(f"✅ 已找到 FFmpeg: {ffmpeg_path}")
        else:
            print("⚠️ 未找到 FFmpeg，正在自動下載...")
            ffmpeg_path = download_ffmpeg()
            if ffmpeg_path:
                print(f"✅ FFmpeg 已安裝至: {ffmpeg_path}")
            else:
                print("❌ 無法自動安裝 FFmpeg，請手動安裝")
                print("請訪問 https://ffmpeg.org/download.html 下載並安裝")
                sys.exit(1)
        
        # 測試 FFmpeg 是否正常工作
        print("正在測試 FFmpeg 是否正常工作...")
        ffmpeg_dir = os.path.dirname(ffmpeg_path)
        test_cmd = [ffmpeg_path, "-version"]
        try:
            result = subprocess.run(test_cmd, capture_output=True, text=True)
            if result.returncode == 0:
                print("✅ FFmpeg 工作正常")
            else:
                print(f"❌ FFmpeg 測試失敗: {result.stderr}")
                return None
        except Exception as e:
            print(f"❌ FFmpeg 測試時發生錯誤: {str(e)}")
            return None
        
        print("所有必要套件已安裝完成")
        return ffmpeg_path
    except Exception as e:
        print(f"安裝套件時發生錯誤: {str(e)}")
        print("提示: 請確保您已安裝Python和pip，並且有安裝套件的權限")
        sys.exit(1)

def select_output_directory():
    """選擇輸出目錄"""
    default_dir = "D:\\music"
    
    print(f"\n預設儲存目錄: {default_dir}")
    custom_dir = input("請輸入自訂儲存路徑，或直接按Enter使用預設路徑: ").strip()
    
    output_dir = custom_dir if custom_dir else default_dir
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"檔案將儲存至: {output_dir}")
    return output_dir

def sanitize_filename(filename):
    """清理檔案名稱，移除不合法字元"""
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

def format_view_count(view_count):
    """將數字轉換為易讀的格式"""
    if view_count >= 1000000000:
        return f"{view_count/1000000000:.1f}B"
    elif view_count >= 1000000:
        return f"{view_count/1000000:.1f}M"
    elif view_count >= 1000:
        return f"{view_count/1000:.1f}K"
    else:
        return str(view_count)

def format_duration(seconds):
    """將秒數轉換為時分秒格式"""
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)

    if hours > 0:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    else:
        return f"{minutes}:{seconds:02d}"

def search_song(song_name, extra_params=""):
    """使用 yt-dlp 搜尋歌曲並返回點閱率最高的影片URL"""
    try:
        print(f"正在搜尋: {song_name}")

        # 使用更精確的搜尋關鍵詞來找到完整歌曲
        keywords = f"{song_name} full song"
        if "完整" not in song_name.lower() and "full" not in song_name.lower():
            keywords = f"{song_name} 完整版 full song"

        # 搜尋更多結果以增加找到完整版的機會
        search_query = f"ytsearch10:{keywords}"
        command = f'yt-dlp {extra_params} --dump-json "{search_query}"'

        # 執行命令並捕獲輸出
        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"搜尋失敗: {result.stderr}")
            return None

        # 分析搜尋結果
        videos = []
        for line in result.stdout.strip().split('\n'):
            if line:
                try:
                    video_info = json.loads(line)
                    duration = video_info.get('duration', 0)  # 獲取影片時長（秒）

                    # 忽略過短的影片（小於60秒的通常是預告或片段）
                    if duration < 60:
                        continue

                    title = video_info.get('title', '未知標題')

                    # 檢查影片標題，排除預告、開場動畫等非完整版本
                    skip_keywords = ['trailer', 'teaser', 'preview', 'short', 'snippet', 'clip',
                                    '預告', '片段', '開場', 'OP', 'アニメ', 'PV', 'CM']

                    # 排除標題中包含明顯非完整版關鍵詞的影片
                    if any(keyword.lower() in title.lower() for keyword in skip_keywords):
                        if not any(fullword in title.lower() for fullword in ['full song', 'full version', '完整版']):
                            continue

                    videos.append({
                        'title': title,
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
            print("找不到相關影片")
            return None

        # 依照觀看次數排序
        videos.sort(key=lambda x: x['view_count'], reverse=True)

        # 顯示搜尋結果
        print("\n找到以下影片:")
        for i, video in enumerate(videos[:5], 1):  # 只顯示前5個結果
            print(f"{i}. {video['title']} - {video['channel']} ({video['view_count_text']} 觀看次數, {video['duration_text']})")

        best_video = videos[0]
        print(f"\n已選擇點閱率最高的影片: {best_video['title']} ({best_video['view_count_text']} 觀看次數, {best_video['duration_text']})")

        return best_video['url']

    except Exception as e:
        print(f"搜尋時發生錯誤: {str(e)}")
        return None

def select_from_search_results(song_name, extra_params=""):
    """讓使用者從搜尋結果中選擇影片"""
    try:
        print(f"正在搜尋: {song_name}")

        # 搜尋更多結果
        search_query = f"ytsearch10:{song_name}"
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

        # 依照觀看次數排序，但顯示時長
        videos.sort(key=lambda x: x['view_count'], reverse=True)

        # 顯示搜尋結果
        print("\n找到以下影片:")
        for i, video in enumerate(videos, 1):
            print(f"{i}. {video['title']} - {video['channel']} ({video['view_count_text']} 觀看次數, {video['duration_text']})")

        # 讓使用者選擇
        while True:
            try:
                choice = input("\n請選擇要下載的影片編號 (或輸入 0 取消): ")
                if choice == '0':
                    return None

                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(videos):
                    selected = videos[choice_idx]
                    print(f"\n已選擇: {selected['title']} ({selected['duration_text']})")
                    return selected['url']
                else:
                    print("無效的選擇，請輸入正確的編號")
            except ValueError:
                print("請輸入數字")

    except Exception as e:
        print(f"搜尋時發生錯誤: {str(e)}")
        return None

def download_as_mp3(youtube_url, output_dir, ffmpeg_path, custom_filename=None):
    """從 YouTube URL 下載 MP3"""
    try:
        print(f"正在處理: {youtube_url}")
        
        # 使用 yt-dlp 直接下載為 MP3
        output_template = os.path.join(output_dir, "%(title)s.%(ext)s")
        if custom_filename:
            # 如果提供了自定義檔名，則使用它
            sanitized_name = sanitize_filename(custom_filename)
            output_template = os.path.join(output_dir, f"{sanitized_name}.%(ext)s")
        
        # 準備命令列參數
        command = [
            "yt-dlp",
            "--no-playlist",
            "-x",
            "--audio-format", "mp3",
            "--audio-quality", "0"
        ]
        
        # 如果找到ffmpeg路徑，則明確指定
        if ffmpeg_path:
            ffmpeg_dir = os.path.dirname(ffmpeg_path)
            command.extend(["--ffmpeg-location", ffmpeg_dir])
            print(f"使用FFmpeg路徑: {ffmpeg_dir}")
        
        # 添加其他參數
        command.extend([
            "--embed-thumbnail",
            "--add-metadata",
            "-o", output_template,
            youtube_url
        ])
        
        print("執行下載命令...")
        print(f"命令: {' '.join(command)}")
        
        # 執行命令
        process = subprocess.run(command, check=True, capture_output=True, text=True)
        
        # 輸出命令結果
        print(process.stdout)
        if process.stderr:
            print(f"警告: {process.stderr}")
        
        print(f"下載完成!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"執行yt-dlp時發生錯誤 (代碼: {e.returncode})")
        if e.stdout:
            print(f"標準輸出: {e.stdout}")
        if e.stderr:
            print(f"錯誤輸出: {e.stderr}")
            
            # 檢查是否為429錯誤
            if "429" in e.stderr:
                print("\n遇到429錯誤 (Too Many Requests)，這意味著YouTube認為我們的下載行為是自動程序。")
                print("請嘗試以下解決方案:")
                print("1. 等待一段時間後再試")
                print("2. 使用VPN或代理改變IP地址")
        return False
    except Exception as e:
        print(f"下載時發生錯誤: {str(e)}")
        return False

def download_song_by_name(output_dir, ffmpeg_path):
    """透過歌曲名稱下載"""
    while True:
        song_name = input("\n請輸入歌曲名稱 (輸入 '0' 退出): ")
        if song_name.lower() == '0':
            break

        print("1. 自動選擇最佳結果 (優先選擇完整版)")
        print("2. 手動從搜尋結果中選擇")
        selection_mode = input("請選擇模式 (1/2): ")

        url = None
        if selection_mode == "2":
            # 手動選擇
            url = select_from_search_results(song_name)
        else:
            # 自動選擇
            url = search_song(song_name)

        if url:
            # 詢問是否使用自定義檔名
            use_custom = input("是否使用自定義檔名? (y/n): ").lower() == 'y'

            custom_name = None
            if use_custom:
                custom_name = input("請輸入檔名 (不需要副檔名): ")

            # 下載歌曲
            download_as_mp3(url, output_dir, ffmpeg_path, custom_name)
        else:
            print("無法找到合適的歌曲，請嘗試更具體的歌名或歌手名稱")

    print("返回主選單")

def batch_download_songs(output_dir, ffmpeg_path):
    """批次下載多首歌曲"""
    song_names = []
    print("\n請輸入多個歌曲名稱 (每行一個，輸入空行結束):")

    while True:
        song = input()
        if not song:
            break
        song_names.append(song)

    if not song_names:
        print("未輸入任何歌曲名稱，返回主選單")
        return

    # 選擇是否手動選擇每首歌
    manual_selection = input("是否手動從搜尋結果中選擇每首歌? (y/n, 預設 n): ").lower() == 'y'

    # 限流選項，避免429錯誤
    apply_rate_limit = input("是否啟用下載間隔以避免429錯誤? (y/n, 預設 y): ").lower() != 'n'
    
    print(f"\n開始搜尋和下載 {len(song_names)} 首歌曲...")
    success_count = 0

    for i, song in enumerate(song_names, 1):
        print(f"\n處理第 {i}/{len(song_names)} 首歌曲:")

        url = None
        if manual_selection:
            url = select_from_search_results(song)
        else:
            url = search_song(song)

        if url:
            # 使用搜尋關鍵字作為檔名
            custom_name = song
            if download_as_mp3(url, output_dir, ffmpeg_path, custom_name):
                success_count += 1
                
            # 如果啟用了間隔，則等待幾秒鐘
            if apply_rate_limit and i < len(song_names):
                wait_time = 5
                print(f"等待 {wait_time} 秒以避免頻率限制...")
                time.sleep(wait_time)
        else:
            print(f"無法找到歌曲: {song}")

    print(f"\n下載完成! 成功: {success_count}/{len(song_names)}")

def download_by_url(output_dir, ffmpeg_path):
    """通過URL下載"""
    while True:
        youtube_url = input("\n請輸入 YouTube 影片網址 (輸入 '0' 退出): ")
        if youtube_url.lower() == '0':
            break

        if "youtube.com" in youtube_url or "youtu.be" in youtube_url:
            # 詢問是否使用自定義檔名
            use_custom = input("是否使用自定義檔名? (y/n): ").lower() == 'y'

            custom_name = None
            if use_custom:
                custom_name = input("請輸入檔名 (不需要副檔名): ")

            download_as_mp3(youtube_url, output_dir, ffmpeg_path, custom_name)
        else:
            print("請輸入有效的 YouTube 網址!")

    print("返回主選單")

def batch_download_urls(output_dir, ffmpeg_path):
    """批次下載多個URL"""
    urls = []
    print("\n請輸入多個 YouTube 網址 (每行一個，輸入空行結束):")

    while True:
        url = input()
        if not url:
            break
        urls.append(url)

    if not urls:
        print("未輸入任何網址，返回主選單")
        return

    # 限流選項，避免429錯誤
    apply_rate_limit = input("是否啟用下載間隔以避免429錯誤? (y/n, 預設 y): ").lower() != 'n'
    
    print(f"\n開始下載 {len(urls)} 個影片...")
    success_count = 0

    for i, url in enumerate(urls, 1):
        print(f"\n處理第 {i}/{len(urls)} 個影片:")
        if download_as_mp3(url, output_dir, ffmpeg_path):
            success_count += 1
            
        # 如果啟用了間隔，則等待幾秒鐘
        if apply_rate_limit and i < len(urls):
            wait_time = 5
            print(f"等待 {wait_time} 秒以避免頻率限制...")
            time.sleep(wait_time)

    print(f"\n下載完成! 成功: {success_count}/{len(urls)}")

def download_from_file(output_dir, ffmpeg_path):
    """從文本文件批量下載"""
    file_path = input("\n請輸入包含YouTube網址的文本文件路徑 (每行一個網址): ")
    
    if not os.path.exists(file_path):
        print(f"找不到檔案: {file_path}")
        return
    
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            urls = [line.strip() for line in file if line.strip()]
        
        if not urls:
            print("檔案中未包含任何網址")
            return
        
        print(f"從文件中讀取了 {len(urls)} 個網址")
        
        # 限流選項，避免429錯誤
        apply_rate_limit = input("是否啟用下載間隔以避免429錯誤? (y/n, 預設 y): ").lower() != 'n'
        
        print(f"開始下載...")
        success_count = 0
        
        for i, url in enumerate(urls, 1):
            print(f"\n處理第 {i}/{len(urls)} 個影片:")
            if download_as_mp3(url, output_dir, ffmpeg_path):
                success_count += 1
                
            # 如果啟用了間隔，則等待幾秒鐘
            if apply_rate_limit and i < len(urls):
                wait_time = 5
                print(f"等待 {wait_time} 秒以避免頻率限制...")
                time.sleep(wait_time)
        
        print(f"\n下載完成! 成功: {success_count}/{len(urls)}")
    except Exception as e:
        print(f"讀取文件時發生錯誤: {str(e)}")

def show_menu(output_dir, ffmpeg_path):
    """顯示主選單"""
    while True:
        print("\n===== YouTube 影片轉 MP3 下載器 =====")
        print(f"檔案將儲存至: {output_dir}")
        print("1. 單一影片下載 (輸入 URL)")
        print("2. 批次下載多個影片 (輸入多個 URL)")
        print("3. 通過歌曲名稱下載")
        print("4. 批次下載多首歌曲 (輸入名稱)")
        print("5. 從文本文件批量下載 URL")
        print("6. 變更儲存目錄")
        print("0. 退出程式")
        
        choice = input("\n請選擇模式: ")
        
        if choice == "1":
            download_by_url(output_dir, ffmpeg_path)
        elif choice == "2":
            batch_download_urls(output_dir, ffmpeg_path)
        elif choice == "3":
            download_song_by_name(output_dir, ffmpeg_path)
        elif choice == "4":
            batch_download_songs(output_dir, ffmpeg_path)
        elif choice == "5":
            download_from_file(output_dir, ffmpeg_path)
        elif choice == "6":
            output_dir = select_output_directory()
        elif choice == "0":
            print("感謝使用，再見！")
            break
        else:
            print("無效的選擇，請重新輸入")

if __name__ == "__main__":
    print("===== YouTube 影片轉 MP3 下載器 =====")
    print("解決 FFmpeg 找不到的問題版本")
    
    # 檢查並安裝必要的套件
    ffmpeg_path = install_required_packages()
    
    if not ffmpeg_path:
        print("無法找到或安裝 FFmpeg，程式無法繼續。")
        sys.exit(1)
    
    # 選擇輸出目錄
    output_dir = select_output_directory()
    
    # 顯示主選單
    show_menu(output_dir, ffmpeg_path)
