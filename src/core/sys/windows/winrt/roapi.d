module core.sys.windows.winrt.roapi;

import core.sys.windows.windef,
  core.sys.windows.basetyps,
  core.sys.windows.winrt.hstring,
  core.sys.windows.winrt.inspectable,
  core.sys.windows.winrt.activation;

enum RO_INIT_TYPE {
  RO_INIT_SINGLETHREADED = 0,
  RO_INIT_MULTITHREADED = 1
}

struct _RO_REGISTRATION_COOKIE {}
alias RO_REGISTRATION_COOKIE = _RO_REGISTRATION_COOKIE*;

interface IApartmentShutdown {}

alias APARTMENT_SHUTDOWN_REGISTRATION_COOKIE = HANDLE;

extern(Windows):

alias PFNGETACTIVATIONFACTORY = HRESULT function(HSTRING, IActivationFactory*);

HRESULT RoInitialize(RO_INIT_TYPE initType);
void RoUninitialize();

HRESULT RoActivateInstance(HSTRING activatableClassId, IInspectable* instance);
HRESULT RoRegisterActivationFactories(HSTRING* activatableClassIds, PFNGETACTIVATIONFACTORY* activationFactoryCallbacks, uint count, RO_REGISTRATION_COOKIE* cookie);
void RoRevokeActivationFactories(RO_REGISTRATION_COOKIE cookie);
HRESULT RoGetActivationFactory(HSTRING activatableClassId, IID* iid, void** factory);
HRESULT RoRegisterForApartmentShutdown(IApartmentShutdown callbackObject, uint* apartmentIdentifier, APARTMENT_SHUTDOWN_REGISTRATION_COOKIE* regCookie);
HRESULT RoUnregisterForApartmentShutdown(APARTMENT_SHUTDOWN_REGISTRATION_COOKIE regCookie);
HRESULT RoGetApartmentIdentifier(ulong* apartmentIdentifier);