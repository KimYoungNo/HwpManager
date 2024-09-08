import win32com.client as win32
from ._HwpRegistery import HwpSecurityModule

def _new_hwp():
    hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
    hwp.RegisterModule("FilePathCheckDLL", str(HwpSecurityModule))
    return hwp

class HwpManager:
    _n_instances = 0
    
    def __new__(cls, *args, **kwargs):
        cls._n_insts += 1
        
        if cls._n_instances == 1:
            HwpSecurityModule.Register()
        return object.__new__(cls)
        
    def __init__(self, run_new=True, visible=True):
        pass

    @classmethod
    def __del__(cls):
        cls._n_insts -= 1

        if cls._n_instances == 0:
            HwpSecurityModule.Unregister()
