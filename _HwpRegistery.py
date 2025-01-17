import os
import subprocess
from enum import Enum

class HwpSecurityReg(str, Enum):
    name = "HwpSecurityModule"
    base_path = os.path.join(os.path.dirname(__file__), "Windows_HwpSecurityModule_Register")
    register_path = os.path.join(base_path, "Register.bat")
    unregister_path = os.path.join(base_path, "Unregister.bat")

for idx, pth in enumerate(HwpSecurityReg):
    if idx == 0:
        continue
        
    if not os.path.exists(pth.value):
        raise OSError(f"Unknown Path: {pth.value}")


def _RunBatch(batch_path, **kwargs):
    return subprocess.run([batch_path], **kwargs)


class HwpSecurityModule:
    def __init__(self):
        self.Register()
        
    def __del__(self):
        self.Unregister()

    @staticmethod
    def Register():
        return _RunBatch(HwpSecurityReg.register_path.value, timeout=5)

    @staticmethod
    def Unregister():
        return _RunBatch(HwpSecurityReg.unregister_path.value, timeout=5)