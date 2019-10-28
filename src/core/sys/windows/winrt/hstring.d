module core.sys.windows.winrt.hstring;

import core.sys.windows.basetsd;

struct HSTRING__ {
  int unused;
}
alias HSTRING = HSTRING__*;

struct HSTRING_HEADER {
  union {
    void* Reserved1;
    version(Win64) char[24] Reserved2;
    else char[20] Reserved2;
  }
}

alias HSTRING_BUFFER = HANDLE;