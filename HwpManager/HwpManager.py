import time
import pythoncom as pycom
import win32gui as gui32
import win32com.client as com32
from collections import deque
from dataclasses import dataclass
from ._HwpRegistery import FilePathCheckDLL

def _enumerate_hwps():
    context = pycom.CreateBindCtx(0)
    running_coms = context.GetRunningObjectTable()

    hwp_monikers = tuple(moniker for moniker in running_coms.EnumRunning()
        if "!HwpObject"==moniker.GetDisplayName(context, moniker).split('.', 1)[0])
    
    return tuple(com32.Dispatch(running_coms.GetObject(moniker).QueryInterface(pycom.IID_IDispatch))
        for moniker in hwp_monikers)

def _grab_hwp():
    hwnd = gui32.GetForegroundWindow()

    if hwnd != 0:
        window_name = gui32.GetWindowText(hwnd)
        
        for hwp in _enumerate_hwps():
            try:
                filepath, filename = hwp.XHwpDocuments.Active_XHwpDocument.FullName.rsplit('\\', 1)
                
                if f"{filename} [{filepath}\\] - 한글" == window_name:
                    return hwp
            except:
                continue
    return None
    
def _new_hwp():
    hwp = com32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    return hwp

def _register_hwp(hwp):
    hwp.RegisterModule("FilePathCheckDLL", str(FilePathCheckDLL))
    return hwp

class _HwpWrapper:
    def __init__(self, hwp):
        self._hwp = _register_hwp(hwp)
        
    def __bool__(self):
        try:
            self._hwp.XHwpDocuments.Active_XHwpDocument.DocumentID
        except:
            return False
        else:
            return True

    def __getattr__(self, name):
        return getattr(self._hwp, name)

    def __del__(self):
        try:
            self._hwp.Quit()
        except:
            pass

    def __str__(self):
        if self:
            return self._hwp.XHwpDocuments.Active_XHwpDocument.FullName
        else:
            return ''
            
    def Open(self, filepath):
        self._hwp.Open(filepath,
            arg="versionwarning:False;forceopen:True;suspendpassword:True")
            
    def Visible(self, option=True):
        if self:
            self._hwp.XHwpWindows.Item(0).Visible = option
            
    def Release(self):
        self.Visible(True)
        self._hwp = None
            
class _HwpQueue(deque):
    def __del__(self):
        for hwp in self:
            del hwp
    
    def append(self, hwp):
        super().append(_HwpWrapper(hwp))

class HwpManager:
    _inst = None
    _hwps = _HwpQueue()
    
    def __new__(cls, *_):
        if cls._inst is None:
            cls._inst = object.__new__(cls)
        return cls._inst
        
    def __init__(self, hwp_id=-1):
        self._hwp_id = hwp_id

    @classmethod
    def __len__(cls):
        return len(cls._hwps)
        
    @classmethod
    def __iter__(cls):
        return iter(cls._hwps)

    @classmethod
    def __getitem__(cls, index):
        return cls._hwps[index]
        
    def __bool__(self):
        return bool(self.__class__._hwps[self._hwp_id])
        
    def __getattr__(self, name):
        return getattr(self.__class__._hwps[self._hwp_id], name)
        
    def New(self):
        self.__class__._hwps.append(_new_hwp())
        self._hwp_id = -1

    def Grab(self):
        hwps = self.__class__._hwps
        hwp = None
        
        while True:
            hwp = _grab_hwp()
            
            if hwp is not None:
                break
            
        if hwp in hwps:
            self._hwp_id = hwps.index(hwp)
        else:
            hwps.append(hwp)
            self._hwp_id = -1

    def Select(self, index):
        self._hwp_id = index % len(self)

    def Release(self):
        hwps = self.__class__._hwps
        
        hwps[self._hwp_id].Release()
        del hwps[self._hwp_id]
        self._hwp_id = -1
        
    def Refresh(self):
        hwps = self.__class__._hwps
        invalids = tuple(hwp for hwp in hwps if not hwp)

        for invalid_hwp in invalids:
            hwps.remove(invalid_hwp)
        
        if invalids:
            self._hwp_id = -1
    
    def KillAll(self):
        hwps = self.__class__._hwps
        
        for hwp in hwps:
            del hwp
                
        hwps.clear()
        self._hwp_id = -1
    
    @property
    def CurrentID(self):
        return self._hwp_id
