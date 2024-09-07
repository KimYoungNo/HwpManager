import os

if os.name != 'nt':
    raise OSError(f"Invalid Operating System: {os.name}")
    
from .HwpManager import HwpManager
