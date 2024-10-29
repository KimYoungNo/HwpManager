from . import HwpUtils
from ._HwpRegistery import HwpSecurityReg

class _InheritHwp:
    _hwp = None
        
    @classmethod
    def __init__(cls, hwp):
        cls._hwp = hwp

class _HRun(_InheritHwp):
    @classmethod
    def __getattr__(cls, name):
        return lambda: cls._hwp.HAction.Run(name)

class _HParameterSet(_InheritHwp):
    _in_action = False
        
    def __init__(self, hwp, hparam, action_obj):
        super().__init__(hwp)
        self._hparameter = getattr(self.__class__._hwp.HParameterSet, hparam)
        self._haction = action_obj
        
    def __enter__(self):
        cls = self.__class__
            
        if cls._in_action:
            raise RuntimeError(
                f"Do not open the new HParameterSet context until the prior one is closed: {self._haction}")
                    
        cls._hwp.HAction.GetDefault(self._haction, self._hparameter.HSet)
        cls._in_action = True
        return self._hparameter
            
    def __exit__(self, *_):
        cls = self.__class__
        cls._hwp.HAction.Execute(self._haction, self._hparameter.HSet)
        cls._in_action = False
        
        del self
        
class _HUtils:
    def __init__(self, hwp_wrapper):
        self._wrap = hwp_wrapper
        
    def __getattr__(self, name):
        return lambda *args, **kwargs: (
            getattr(HwpUtils, name)(self._wrap, *args, **kwargs))

def _register_hwp(hwp):
    hwp.RegisterModule("FilePathCheckDLL", HwpSecurityReg.name.value)
    return hwp
    

class HwpWrapper:
    def __init__(self, hwp):
        self._hwp = _register_hwp(hwp)
        self._run = _HRun(self._hwp)
        self._util = _HUtils(self)
        
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
            self._hwp.Clear(1)
            self._hwp.Quit()
        except:
            pass

    def __str__(self):
        if self:
            return self._hwp.XHwpDocuments.Active_XHwpDocument.FullName
        else:
            return ''
            
    def _Release(self):
        self.Visible(True)
        self._hwp = None
        
    def Visible(self, option=True):
        if self:
            self._hwp.XHwpWindows.Item(0).Visible = option
            
    def Open(self, filepath):
        self._hwp.Open(filepath,
            arg="versionwarning:False;forceopen:True;suspendpassword:True")
            
    def HParameterSet(self, hparam, action_obj):
        return _HParameterSet(self._hwp, str(hparam), str(action_obj))
    
    @property
    def Run(self):
        return self._run
        
    @property
    def Util(self):
        return self._util