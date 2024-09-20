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

def _grab_hwnd_hwp():
    hwnd = gui32.GetForegroundWindow()

    if hwnd != 0:
        window_name = gui32.GetWindowText(hwnd)
        
        for hwp in _enumerate_hwps():
            try:
                filepath, filename = hwp.XHwpDocuments.Active_XHwpDocument.FullName.rsplit('\\')
            
                if f"{filename} [{filepath}\\] - 한글" == window_name:
                    return hwnd, _register_hwp(hwp)
            except:
                continue
    return None, None

def _grab_hwp():
    return _grab_hwnd_hwp()[1]
    
def _new_hwp():
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    return _register_hwp(hwp)

def _register_hwp(hwp):
    hwp.RegisterModule("FilePathCheckDLL", str(HwpSecurityModule))
    return hwp

def _is_alive_hwp(hwp):
    try:
        hwp.XHwpDocuments.Active_XHwpDocument.FullName
    except:
        return False
    else:
        return True

@dataclass(init=True)
class InstanceOccupied:
    instance: object
    occupied: bool
    
    def __bool__(self):
        return self.occupied

class HwpManager:
    _hwps = deque()
        
    def __init__(self, hwp_id=None):
        self._hwp_id = hwp_id

    def __getattr__(self, name):
        if self._hwp_id is not None:
            return getattr(self.__class__._hwps[self._hwp_id].instance, name)
        else:
            raise ValueError()

    def __del__(self):
        self.Release()

    @classmethod
    def __len__(cls):
        return len(cls._hwps)

    @classmethod
    def __getitem__(cls, index):
        return cls._hwps[index].instance

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
        invalids = tuple(hwp for hwp in cls._hwps if not _is_alive_hwp(hwp))

        for invalid_hwp in invalids:
            cls._hwps.remove(invalid_hwp)
        
    def New(self):
        self.__class__._hwps.append(InstanceOccupied(_new_hwp(), True))
        self._AfterAppend()

    def Grab(self):
        self.__class__._hwps.append(InstanceOccupied(_grab_hwp(), True))
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
        if self._hwp_id is not None:
            self.__class__._hwps[self._hwp_id].occupied = False
            self._hwp_id = None
        self._RenewSecurityModule()
    
    @classmethod
    def QuitAll(cls):
        for hwp in cls._hwps:
            try:
                hwp.instance.Quit()
            except:
                pass
        cls._hwps.clear()
    
