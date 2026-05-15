import os
import shutil

save_dir = "工资条拆分文件"

if os.path.exists(save_dir):
    for filename in os.listdir(save_dir):
        file_path = os.path.join(save_dir, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"删除: {file_path}")
        except Exception as e:
            print(f"删除失败 {file_path}: {e}")

print("清理完成")