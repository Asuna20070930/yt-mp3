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
from mutagen.mp3 import MP3
from mutagen.id3 import ID3

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

def search_song(song_name, prefer_full=True, extra_params=""):
    """使用 yt-dlp 搜尋歌曲並返回點閱率最高的影片URL，排除過長和過短的影片

    參數:
    song_name: 搜尋的歌曲名稱
    prefer_full: 是否優先選擇較長的影片版本（很可能是完整版）
    extra_params: 額外的yt-dlp參數字符串
    """
    try:
        print(f"正在搜尋: {song_name}")

        # 使用更精確的搜尋關鍵詞來找到完整歌曲
        keywords = f"{song_name} full song"
        if "完整" not in song_name.lower() and "full" not in song_name.lower():
            keywords = f"{song_name} 完整版 full song"

        # 搜尋更多結果以增加找到合適版本的機會
        search_query = f"ytsearch15:{keywords}"  # 增加到15個結果
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

                    # 過濾條件：排除少於1分鐘和超過20分鐘的影片
                    if duration < 60 or duration > 1200:  # 60秒=1分鐘, 1200秒=20分鐘
                        continue

                    title = video_info.get('title', '未知標題')

                    # 檢查影片標題，排除預告、開場動畫等非完整版本
                    skip_keywords = ['trailer', 'teaser', 'preview', 'short', 'snippet', 'clip',
                                    '預告', '片段', '開場', 'アニメ', 'PV', 'CM']

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
            print("找不到符合時長要求的相關影片，嘗試使用更寬鬆的搜尋條件...")
            # 如果找不到結果，使用原始的搜尋方式
            return search_song_fallback(song_name, extra_params)

        # 首先按觀看次數排序
        videos.sort(key=lambda x: x['view_count'], reverse=True)

        # 顯示搜尋結果
        print("\n找到以下符合時長要求的影片 (1分鐘 ~ 20分鐘):")
        for i, video in enumerate(videos[:5], 1):  # 只顯示前5個結果
            print(f"{i}. {video['title']} - {video['channel']} ({video['view_count_text']} 觀看次數, {video['duration_text']})")

        best_video = videos[0]
        print(f"\n已選擇點閱率最高的影片: {best_video['title']} ({best_video['view_count_text']} 觀看次數, {best_video['duration_text']})")

        return best_video['url']

    except Exception as e:
        print(f"搜尋時發生錯誤: {str(e)}")
        return None

def search_song_fallback(song_name, extra_params=""):
    """備用搜尋方法，使用較寬鬆的條件，但仍然過濾過長和過短的影片"""
    try:
        print(f"使用備用搜尋方法尋找: {song_name}")

        # 使用基本搜尋
        search_query = f"ytsearch15:{song_name}"
        command = f'yt-dlp {extra_params} --dump-json "{search_query}"'

        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"備用搜尋失敗: {result.stderr}")
            return None

        videos = []
        for line in result.stdout.strip().split('\n'):
            if line:
                try:
                    video_info = json.loads(line)
                    duration = video_info.get('duration', 0)

                    # 過濾條件：排除少於1分鐘和超過20分鐘的影片
                    if duration < 60 or duration > 1200:  # 60秒=1分鐘, 1200秒=20分鐘
                        continue

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
            print("找不到符合時長要求的相關影片!")
            return None

        # 依照觀看次數排序
        videos.sort(key=lambda x: x['view_count'], reverse=True)

        print("\n備用搜尋找到以下符合時長要求的影片 (1分鐘 ~ 20分鐘):")
        for i, video in enumerate(videos[:5], 1):
            print(f"{i}. {video['title']} - {video['channel']} ({video['view_count_text']} 觀看次數, {video['duration_text']})")

        best_video = videos[0]
        print(f"\n已選擇點閱率最高的影片: {best_video['title']} ({best_video['view_count_text']} 觀看次數, {best_video['duration_text']})")

        return best_video['url']

    except Exception as e:
        print(f"備用搜尋時發生錯誤: {str(e)}")
        return None

def get_mp3_metadata(file_path):
    """獲取MP3文件的元數據"""
    try:
        # 使用mutagen讀取MP3標籤
        audio = MP3(file_path, ID3=ID3)
        if audio.tags is None:
            return None

        metadata = {
            'title': str(audio.tags.get('TIT2', '未知標題')),
            'artist': str(audio.tags.get('TPE1', '未知藝術家')),
            'album': str(audio.tags.get('TALB', '未知專輯')),
            'year': str(audio.tags.get('TDRC', '未知年份')),
            'genre': str(audio.tags.get('TCON', '未知類型')),
            'duration': format_duration(int(audio.info.length))
        }
        return metadata
    except Exception as e:
        print(f"讀取ID3標籤時發生錯誤: {str(e)}")
        return None

def download_as_mp3(youtube_url, extra_params=""):
    """從 YouTube URL 下載 MP3，並在下載後顯示檔案資訊"""
    try:
        print(f"正在處理: {youtube_url}")

        # 使用 yt-dlp 直接下載為 MP3
        output_template = f"{output_dir}/%(title)s.%(ext)s"

        # 使用更輕量級的選項來下載MP3，保留元數據但不下載縮圖
        command = f'yt-dlp {extra_params} --no-playlist -x --audio-format mp3 --audio-quality 0 --add-metadata --no-embed-thumbnail --no-write-thumbnail -o "{output_template}" "{youtube_url}"'

        print("執行下載命令...")
        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"下載失敗: {result.stderr}")

            # 如果是429錯誤，提供更具體的提示
            if "429" in result.stderr:
                print("\n遇到429錯誤 (Too Many Requests)，這意味著YouTube認為我們的下載行為是自動程序。")
                print("請嘗試以下解決方案:")
                print("1. 等待一段時間後再試")
                print("2. 使用cookies選項重新運行此程序")
                print("3. 使用VPN或代理改變IP地址")
                return False
            return False

        # 獲取最新下載的檔案
        files = glob.glob(f"{output_dir}/*.mp3")
        if files:
            latest_file = max(files, key=os.path.getctime)
            filename = os.path.basename(latest_file)
            
            # 從檔案本身獲取ID3標籤資訊
            print("\n下載完成!")
            print(f"檔案名稱: {filename}")
            
            # 從檔案本身讀取標籤資訊
            metadata = get_mp3_metadata(latest_file)
            if metadata:
                print(f"標題: {metadata['title']}")
                print(f"演出者: {metadata['artist']}")
                if metadata['album'] != '未知專輯':
                    print(f"專輯: {metadata['album']}")
            
            # 直接詢問新檔名，如果輸入為空則保持原檔名
            new_name = input("請輸入新檔名（直接按Enter保持原檔名，無需.mp3副檔名）: ")
            if new_name:
                new_filepath = f"{output_dir}/{sanitize_filename(new_name)}.mp3"
                try:
                    os.rename(latest_file, new_filepath)
                    print(f"已重新命名為: {os.path.basename(new_filepath)}")
                except Exception as e:
                    print(f"重新命名檔案時發生錯誤: {str(e)}")
            
            return True
        else:
            print("找不到下載的檔案。")
            return False
            
    except Exception as e:
        print(f"下載時發生錯誤: {str(e)}")
        return False

def select_from_search_results(song_name, extra_params=""):
    """讓用戶從搜尋結果中選擇一個影片"""
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
                    
                    # 過濾條件：排除少於30秒和超過30分鐘的影片
                    if duration < 30 or duration > 1800:
                        continue
                        
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
            
        # 顯示搜尋結果供用戶選擇
        print("\n請從以下結果中選擇一個影片:")
        for i, video in enumerate(videos[:10], 1):  # 只顯示前10個結果
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

# 新增: 下載歌曲函數
def download_song_by_name(extra_params=""):
    while True:
        song_name = input("請輸入歌曲名稱 (輸入 '0' 退出): ")
        if song_name.lower() == '0':
            break
            
        print("1. 自動選擇最佳結果 (優先選擇完整版)")
        print("2. 手動從搜尋結果中選擇")
        selection_mode = input("請選擇模式 (1/2): ")

        url = None
        if selection_mode == "2":
            # 手動選擇
            url = select_from_search_results(song_name, extra_params)
        else:
            # 自動選擇
            url = search_song(song_name, prefer_full=True, extra_params=extra_params)

        if url:
            # 直接詢問檔名，若為空白則不修改
            custom_name = input("請輸入自定義檔名 (直接按Enter則使用原檔名): ")
            
            # 若輸入為空白，則傳入None表示使用原檔名
            if custom_name.strip() == "":
                custom_name = None
                
            # 下載歌曲
            download_as_mp3(url, custom_name, extra_params)
        else:
            print("無法找到合適的歌曲，請嘗試更具體的歌名或歌手名稱")
            
    print("程式已結束")

# 批次下載多首歌曲 (改進版)
def batch_download_songs(extra_params=""):
    song_names = []
    print("請輸入多個歌曲名稱 (每行一個，輸入空行結束):")

    while True:
        song = input()
        if not song:
            break
        song_names.append(song)

    # 選擇是否優先下載完整版
    prefer_full = input("是否優先選擇完整版歌曲? (y/n, 預設 y): ").lower() != 'n'

    # 選擇是否手動選擇每首歌
    manual_selection = input("是否手動從搜尋結果中選擇每首歌? (y/n, 預設 n): ").lower() == 'y'

    # 新增：限流選項，避免429錯誤
    apply_rate_limit = input("是否啟用下載間隔以避免429錯誤? (y/n, 預設 y): ").lower() != 'n'
    rate_limit_args = "--limit-rate 500K --sleep-interval 10" if apply_rate_limit else ""
    if rate_limit_args and extra_params:
        extra_params = f"{extra_params} {rate_limit_args}"
    elif rate_limit_args:
        extra_params = rate_limit_args

    print(f"\n開始搜尋和下載 {len(song_names)} 首歌曲...")
    success_count = 0

    for i, song in enumerate(song_names, 1):
        print(f"\n處理第 {i}/{len(song_names)} 首歌曲:")

        url = None
        if manual_selection:
            url = select_from_search_results(song, extra_params)
        else:
            url = search_song(song, prefer_full=prefer_full, extra_params=extra_params)

        if url:
            if download_as_mp3(url, extra_params):
                success_count += 1
        else:
            print(f"無法找到歌曲: {song}")

    print(f"\n下載完成! 成功: {success_count}/{len(song_names)}")

# 原有的URL下載功能
def download_by_url(extra_params=""):
    while True:
        youtube_url = input("請輸入 YouTube 影片網址 (輸入 '0' 退出): ")
        if youtube_url.lower() == '0':
            break

        if "youtube.com" in youtube_url or "youtu.be" in youtube_url:
            # 下載歌曲，不再提前詢問檔名
            download_as_mp3(youtube_url, extra_params)
        else:
            print("請輸入有效的 YouTube 網址!")

    print("程式已結束")

# 批次下載多個URL
def batch_download_urls(extra_params=""):
    urls = []
    print("請輸入多個 YouTube 網址 (每行一個，輸入空行結束):")

    while True:
        url = input()
        if not url:
            break
        urls.append(url)

    # 新增：限流選項，避免429錯誤
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
        print("1. 單一影片下載")
        print("2. 批次下載多個影片")
        print("3. 通過歌曲名稱下載")
        print("4. 批次下載多首歌曲 ")
        print("5. 從文本文件批量下載 ")
        print("6. 變更儲存目錄")
        print("0. 退出程式")
        
        choice = input("\n請選擇模式: ")
        
        # 整合 ffmpeg_path 到 extra_params
        extra_params = f"--ffmpeg-location \"{ffmpeg_path}\""
        
        if choice == "1":
            download_by_url(extra_params)
        elif choice == "2":
            batch_download_urls(extra_params)
        elif choice == "3":
            download_song_by_name(extra_params)
        elif choice == "4":
            batch_download_songs(extra_params)
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
