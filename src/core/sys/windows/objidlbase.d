module core.sys.windows.objidlbase;

import core.sys.windows.w32api,
  core.sys.windows.windef,
  core.sys.windows.basetyps,
  core.sys.windows.unknwn;

enum {
  APTTYPE_CURRENT = -1,
  APTTYPE_STA = 0,
  APTTYPE_MTA = 1,
  APTTYPE_NA = 2,
  APTTYPE_MAINSTA = 3
}
alias APTTYPE = int;

enum {
  APTTYPEQUALIFIER_NONE,
  APTTYPEQUALIFIER_IMPLICIT_MTA,
  APTTYPEQUALIFIER_NA_ON_MTA,
  APTTYPEQUALIFIER_NA_ON_STA,
  APTTYPEQUALIFIER_NA_ON_IMPLICIT_MTA,
  APTTYPEQUALIFIER_NA_ON_MAINSTA,
  APTTYPEQUALIFIER_APPLICATION_STA,
  APTTYPEQUALIFIER_RESERVED_1
}
alias APTTYPEQUALIFIER = int;

static if (_WIN32_WINNT >= 0x601)
extern(Windows) HRESULT CoGetApartmentType(APTTYPE* pAptType, APTTYPEQUALIFIER* pAptQualifier);

immutable GUID IID_IAgileObject = { 0x94ea2b94, 0xe9cc, 0x49e0, [0xc0, 0xff, 0xee, 0x64, 0xca, 0x8f, 0x5b, 0x90] };

interface IAgileObject : IUnknown {
}