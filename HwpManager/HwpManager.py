import win32com.client as win32
from ._HwpRegistery import HwpSecurityModule

def _new_hwp():
    hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
    hwp.RegisterModule("FilePathCheckDLL", str(HwpSecurityModule))
    return hwp

class HwpManager:
    def __init__(self, run_new=True, visible=True):
        HwpSecurityModule.Register()

    def __del__(self):
        HwpSecurityModule.Unregister()
