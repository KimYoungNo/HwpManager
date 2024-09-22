import os
from enum import Enum

class WindowsPath(str, Enum):
    user = os.path.expanduser('~')
    document = os.path.join(user, "Documents")
    
def ReplaceFileExt(filepath, ext):
    return f"{filepath[:filepath.rfind('.')]}.{ext}"
    
def ExtractFileString(filepath):
    return filepath.rsplit('\\')[-1][:filepath.rfind('.')]