module core.sys.windows.winrt.activation;

import core.sys.windows.windef,
  core.sys.windows.basetyps,
  core.sys.windows.winrt.inspectable;

immutable IID IID_IActivationFactory = { 0x00000035, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46] };

interface IActivationFactory : IInspectable {
  HRESULT ActivateInstance(IInspectable* instance);
}