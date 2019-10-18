module core.sys.windows.winrt.inspectable;

import core.sys.windows.windef,
  core.sys.windows.basetyps,
  core.sys.windows.unknwn,
  core.sys.windows.winrt.hstring;

enum TrustLevel {
  BaseTrust,
  PartialTrust,
  FullTrust
}

immutable IID IID_IInspectable = { 0xAF86E2E0, 0xB12D, 0x4c6a, [0x9C, 0x5A, 0xD7, 0xAA, 0x65, 0x10, 0x1E, 0x90] };

interface IInspectable : IUnknown {
  HRESULT GetIids(uint* iidCount, IID** iids);
  HRESULT GetRuntimeClassName(HSTRING* className);
  HRESULT GetTrustLevel(TrustLevel* trustLevel);
}