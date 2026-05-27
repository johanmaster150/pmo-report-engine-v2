import os
import shutil
import stat

def remove_readonly(func, path, _):
    os.chmod(path, stat.S_IWRITE)
    func(path)

try:
    shutil.rmtree('.git', onerror=remove_readonly)
    print("Deleted .git successfully")
except Exception as e:
    print(f"Error deleting .git: {e}")
