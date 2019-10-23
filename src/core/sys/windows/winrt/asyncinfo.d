module core.sys.windows.winrt.asyncinfo;

import core.sys.windows.windef,
  core.sys.windows.basetyps,
  core.sys.windows.winrt.inspectable;

enum AsyncStatus {
  Started,
  Completed,
  Canceled,
  Error
}

immutable GUID IID_IAsyncInfo = { 0x00000036, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46] };

interface IAsyncInfo : IInspectable {
  HRESULT get_Id(uint* id);
  HRESULT get_Status(AsyncStatus* status);
  HRESULT get_ErrorCode(HRESULT* errorCode);
  HRESULT Cancel();
  HRESULT Close();
}