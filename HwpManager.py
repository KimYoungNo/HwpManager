import pythoncom as pycom
import win32gui as gui32
import win32com.client as com32
from .HwpWrapper import HwpWrapper
from ._HwpRegistery import HwpSecurityModule

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
                filename = hwp.XHwpDocuments.Active_XHwpDocument.FullName.rsplit('\\', 1)[-1]

                if filename in window_name and window_name[-4:] == "- 한글":
                    return HwpWrapper(hwp)
            except:
                continue
    return None
    
def _new_hwp():
    return HwpWrapper(com32.gencache.EnsureDispatch("HWPFrame.HwpObject"))

class _HwpQueue(list):
    def __del__(self):
        self._deque_all()
        
    def _deque_all(self):
        while len(self):
            hwp = self.pop(0)
            del hwp

class HwpManager:
    _inst = None
    _sec = None
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
        return bool(self.hwp)
        
    def __getattr__(self, name):
        return getattr(self.hwp, name)
        
    @classmethod
    @property
    def hwps(cls):
        return cls._hwps
    
    @property
    def hwp(self):
        return self.hwps[self._hwp_id]
    
    @classmethod
    def MainThread(cls):
        cls._sec = HwpSecurityModule()
        
    def New(self, filepath=''):
        hwp = _new_hwp()
        
        if filepath:
            hwp.Open(filepath)
            
        self.hwps.append(hwp)
        self._hwp_id = -1

    def Grab(self):
        hwp = None
        
        while hwp is None:
            hwp = _grab_hwp()
            
        if hwp in self.hwps:
            self._hwp_id = hwps.index(hwp)
        else:
            self.hwps.append(hwp)
            self._hwp_id = -1

    def Select(self, index):
        self._hwp_id = index % len(self)

    def Release(self):
        self.hwp._Release()
        self.hwps.remove(self.hwp)
        self._hwp_id = -1
        
    def Refresh(self):
        invalids = tuple(hwp for hwp in self.hwps if not hwp)

        for invalid_hwp in invalids:
            self.hwps.remove(invalid_hwp)
        
        if invalids:
            self._hwp_id = -1
    
    def KillAll(self):
        self.hwps._deque_all()
        self._hwp_id = -1
    
    @property
    def CurrentID(self):
        return self._hwp_id