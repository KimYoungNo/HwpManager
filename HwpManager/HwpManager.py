import pythoncom as pycom
import win32con as con32
import win32gui as gui32
import win32com.client as win32
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

class HwpManager:
    _n_instances = 0
    
    def __new__(cls, *args, **kwargs):
        cls._n_instances += 1
        
        if cls._n_instances == 1:
            HwpSecurityModule.Register()
        return object.__new__(cls)

    @classmethod
    def __del__(cls):
        cls._n_instances -= 1

        if cls._n_instances == 0:
            HwpSecurityModule.Unregister()

    @classmethod
    def __len__(cls):
        return cls._n_instances
        
    def __init__(self):
        self._hwp = None

    def __getattr__(self, name):
        if self._hwp is not None:
            return getattr(self._hwp, name)
        else:
            raise AttributeError()

    def New(self):
        while self._hwp is None:
            self._hwp = _new_hwp()

    def Grab(self):
        while self._hwp is None:
            self._hwp = _grab_hwp()

    def Release(self):
        if self._hwp is not None:
            try:
                self._hwp.Quit()
            except:
                pass
            self._hwp = None

    
