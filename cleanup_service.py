import os
import time
import shutil

# 获取当前脚本所在目录
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
# 目标目录定义
UPLOAD_DIR = os.path.join(CURRENT_DIR, "docx服务", "uploads")
OUTPUT_DIR = os.path.join(CURRENT_DIR, "docx服务", "outputs")

CLEAN_INTERVAL = 12 * 60 * 60  # 12小时（秒）

def cleanup_directories():
    """清理上传和输出目录"""
    for directory in [UPLOAD_DIR, OUTPUT_DIR]:
        if not os.path.exists(directory):
            continue
            
        print(f"正在清理目录: {directory}")
        try:
            # 清理目录下的所有文件和子目录
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f'清理 {file_path} 失败: {e}')
        except Exception as e:
            print(f"访问 {directory} 失败: {e}")

if __name__ == "__main__":
    print(f"定时清理脚本启动，清理间隔: 12小时")
    while True:
        cleanup_directories()
        print(f"下次清理时间: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time() + CLEAN_INTERVAL))}")
        time.sleep(CLEAN_INTERVAL)
