import time
import pythoncom as pycom
import win32con as con32
import win32gui as gui32
import win32com.client as win32
from collections import deque
from dataclasses import dataclass
from ._HwpRegistery import HwpSecurityModule

def _enumerate_hwps():
    context = pycom.CreateBindCtx(0)
    running_coms = context.GetRunningObjectTable()

    hwp_monikers = tuple(moniker for moniker in running_coms.EnumRunning()
        if "!HwpObject"==moniker.GetDisplayName(context, moniker).split('.', 1)[0])
    
    return tuple(com32.Dispatch(running_coms.GetObject(moniker).QueryInterface(IID_IDispatch))
        for moniker in hwp_monikers)

def _grab_hwp():
    hwnd = gui32.GetForegroundWindow()
    hwp = None

    if hwnd != 0:
        window_name = gui32.GetWindowText(hwnd)
        
        for hwp in _enumerate_hwps():
            try:
                filepath, filename = hwp.XHwpDocuments.Active_XHwpDocument.FullName.rsplit('\\')
            
                if f"{filename} [{filepath}\\] - 한글" == window_name:
                    hwp = _register_hwp(hwp)
            except:
                continue
    return hwp
    
def _new_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    return _register_hwp(hwp)

def _register_hwp(hwp):
    hwp.RegisterModule("FilePathCheckDLL", str(HwpSecurityModule))
    return hwp

class _HwpWrapper:
    def __init__(self, hwp):
        self._hwp = hwp

    def __getattr__(self, name):
        return getattr(self._hwp, name)

    def __del__(self):
        self._hwp.Quit()

    def __str__(self):
        return self._hwp.XHwpDocuments.Active_XHwpDocument.FullName

    def __bool__(self):
        try:
            self._hwp.CheckXObject(True)
        except:
            return False
        else:
            return True
    

class HwpManager:
    _hwps = deque()
        
    def __init__(self, hwp_id=None):
        self._hwp_id = hwp_id

    def __del__(self):
        self.Release()

    @classmethod
    def __len__(cls):
        return len(cls._hwps)

    @classmethod
    def __getitem__(cls, index):
        return cls._hwps[index]

    @classmethod
    def _RenewSecurityModule(cls):
        num = len(cls._hwps)
        
        if num < 1:
            HwpSecurityModule.Unregister()
        elif num == 1:
            HwpSecurityModule.Register()

    def _AfterAppend(self):
        self._hwp_id = -1
        self._RenewSecurityModule()

    @classmethod
    def _DequeInvalidHwp(cls):
        invalids = tuple(hwp for hwp in cls._hwps if not hwp)

        for invalid_hwp in invalids:
            cls._hwps.remove(invalid_hwp)
        
    def New(self):
        self.__class__._hwps.append(_HwpWrapper(_new_hwp()))
        self._AfterAppend()

    def Grab(self):
        self.__class__._hwps.append(_HwpWrapper(_grab_hwp()))
        self._AfterAppend()

    def Select(self, nth):
        hwps = self.__class__._hwps
        hwp_id = nth % len(hwps)
        
        if not hwps[hwp_id].occupied:
            hwps[hwp_id].occupied = True
            hwps[self._hwp_id].occupied = False
            self._hwp_id = hwp_id
        else:
            ValueError("Pre-ocuupied Instance Selected")

    def Release(self):
        self._hwp_id = None
        self._RenewSecurityModule()
    
    @classmethod
    def KillAll(cls):
        for hwp in cls._hwps:
            try:
                del hwp
            except:
                pass
        cls._hwps.clear()
    
