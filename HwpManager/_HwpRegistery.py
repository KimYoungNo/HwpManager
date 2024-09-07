import os
import subprocess
from enum import Enum

class HwpSecurityReg(str, Enum):
    base_path = os.path.join(os.path.dirname(__file__), "Windows_HwpSecurityModule_Register")
    register_path = os.path.join(base_path, "Register.bat")
    unregister_path = os.path.join(base_path, "Unregister.bat")

for pth in HwpSecurityReg:
    if not os.path.exists(pth.value):
        raise OSError(f"Unknown Path: {pth.value}")


def _RunBatch(batch_path, **kwargs):
    return subprocess.run([batch_path], **kwargs)
    
def HwpSecurityModuleRegister():
    return _RunBatch(HwpSecurityReg.register_path.value, timeout=5)
    
def HwpSecurityModuleUnregister():
    return _RunBatch(HwpSecurityReg.unregister_path.value, timeout=5)


class HwpSecurityModule:
    def __str__(self):
        return "HwpSecurityModule.dll"
        
    def __enter__(self):
        HwpSecurityModuleRegister()
        
    def __exit__(self, exc_val, exc_type, exc_trace):
        HwpSecurityModuleUnregister()
