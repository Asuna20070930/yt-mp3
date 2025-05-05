import os
import re
import subprocess
import sys
import platform
import shutil
import zipfile
import tempfile
from urllib.request import urlretrieve
import time

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
        except ImportError:
            print("正在安裝 yt-dlp...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "yt-dlp"])
        
        # 檢查 FFmpeg 是否存在
        print("正在檢查 FFmpeg...")
        ffmpeg_path = check_ffmpeg()
        if ffmpeg_path:
            print(f"已找到 FFmpeg: {ffmpeg_path}")
        else:
            print("未找到 FFmpeg，正在自動下載...")
            ffmpeg_path = download_ffmpeg()
            if ffmpeg_path:
                print(f"FFmpeg 已安裝至: {ffmpeg_path}")
            else:
                print("無法自動安裝 FFmpeg，請手動安裝")
                print("請訪問 https://ffmpeg.org/download.html 下載並安裝")
                sys.exit(1)
        
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

def download_as_mp3(youtube_url, output_dir, ffmpeg_path):
    """下載YouTube影片並轉換為MP3格式"""
    try:
        print(f"正在處理: {youtube_url}")
        
        # 使用 yt-dlp 直接下載為 MP3
        command = [
            "yt-dlp",
            "--no-playlist",
            "-x",
            "--audio-format", "mp3",
            "--audio-quality", "0"
        ]
        
        # 指定FFmpeg位置
        if ffmpeg_path:
            ffmpeg_dir = os.path.dirname(ffmpeg_path)
            command.extend(["--ffmpeg-location", ffmpeg_dir])
            print(f"使用FFmpeg路徑: {ffmpeg_dir}")
        
        # 添加輸出路徑和URL
        output_template = os.path.join(output_dir, "%(title)s.%(ext)s").replace("\\", "/")
        command.extend([
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
        print(f"執行yt-dlp時發生錯誤: {str(e)}")
        if e.stdout:
            print(f"標準輸出: {e.stdout}")
        if e.stderr:
            print(f"錯誤輸出: {e.stderr}")
        return False
    except Exception as e:
        print(f"下載時發生錯誤: {str(e)}")
        return False

def main(output_dir, ffmpeg_path):
    """單一下載模式"""
    while True:
        youtube_url = input("\n請輸入YouTube影片網址 (輸入 '0' 退出): ")
        if youtube_url == '0':
            break
        
        if "youtube.com" in youtube_url or "youtu.be" in youtube_url:
            download_as_mp3(youtube_url, output_dir, ffmpeg_path)
        else:
            print("請輸入有效的YouTube網址!")
    
    print("程式已結束")

def batch_download(output_dir, ffmpeg_path):
    """批次下載多個影片"""
    urls = []
    print("\n請輸入多個YouTube網址 (每行一個，輸入空行結束):")
    
    while True:
        url = input()
        if not url:
            break
        urls.append(url)
    
    if not urls:
        print("未輸入任何網址，返回主選單")
        return
    
    print(f"開始下載 {len(urls)} 個影片...")
    success_count = 0
    
    for i, url in enumerate(urls, 1):
        print(f"\n處理第 {i}/{len(urls)} 個影片:")
        if download_as_mp3(url, output_dir, ffmpeg_path):
            success_count += 1
        # 短暫暫停，避免過快請求
        time.sleep(1)
    
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
        print(f"開始下載...")
        success_count = 0
        
        for i, url in enumerate(urls, 1):
            print(f"\n處理第 {i}/{len(urls)} 個影片:")
            if download_as_mp3(url, output_dir, ffmpeg_path):
                success_count += 1
            # 短暫暫停，避免過快請求
            time.sleep(1)
        
        print(f"\n下載完成! 成功: {success_count}/{len(urls)}")
    except Exception as e:
        print(f"讀取文件時發生錯誤: {str(e)}")

def show_menu(output_dir, ffmpeg_path):
    """顯示主選單"""
    while True:
        print("\n===== YouTube影片轉MP3下載器 =====")
        print(f"檔案將儲存至: {output_dir}")
        print("1. 單一影片下載")
        print("2. 批次下載多個影片")
        print("3. 從文本文件批量下載")
        print("4. 變更儲存目錄")
        print("0. 退出程式")
        
        choice = input("\n請選擇模式: ")
        
        if choice == "1":
            main(output_dir, ffmpeg_path)
        elif choice == "2":
            batch_download(output_dir, ffmpeg_path)
        elif choice == "3":
            download_from_file(output_dir, ffmpeg_path)
        elif choice == "4":
            output_dir = select_output_directory()
        elif choice == "0":
            print("感謝使用，再見！")
            break
        else:
            print("無效的選擇，請重新輸入")

if __name__ == "__main__":
    print("===== YouTube影片轉MP3下載器 =====")
    print("作者: 移植自Colab版本")
    
    # 檢查並安裝必要的套件
    ffmpeg_path = install_required_packages()
    
    # 選擇輸出目錄
    output_dir = select_output_directory()
    
    # 顯示主選單
    show_menu(output_dir, ffmpeg_path)