module core.sys.windows.winrt.roapi;

import core.sys.windows.windef,
  core.sys.windows.basetyps,
  core.sys.windows.winrt.hstring,
  core.sys.windows.winrt.inspectable;

enum RO_INIT_TYPE {
  RO_INIT_SINGLETHREADED,
  RO_INIT_MULTITHREADED
}

extern(Windows):

HRESULT RoInitialize(RO_INIT_TYPE initType);
void RoUninitialize();

HRESULT RoActivateInstance(HSTRING activatableClassId, IInspectable* instance);
HRESULT RoGetActivationFactory(HSTRING activatableClassId, IID* iid, void** factory);