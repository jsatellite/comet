module core.sys.windows.winrt.winstring;

import core.sys.windows.windef,
  core.sys.windows.winrt.hstring;

extern(Windows):

HRESULT WindowsCreateString(const(wchar)* sourceString, uint length, HSTRING* string);
HRESULT WindowsCreateStringReference(const(wchar)* sourceString, uint length, HSTRING_HEADER* hstringHeader, HSTRING* string);
HRESULT WindowsDeleteString(HSTRING string);
HRESULT WindowsDuplicateString(HSTRING string, HSTRING* newString);
uint WindowsGetStringLen(HSTRING string);
wchar* WindowsGetStringRawBuffer(HSTRING string, uint* length);
BOOL WindowsIsStringEmpty(HSTRING string);
HRESULT WindowsStringHasEmbeddedNull(HSTRING string, BOOL* hasEmbedNull);
HRESULT WindowsCompareStringOrdinal(HSTRING string1, HSTRING string2, int* result);
HRESULT WindowsSubstring(HSTRING string, uint startIndex, HSTRING* newString);
HRESULT WindowsSubstringWithSpecifiedLength(HSTRING string, uint startIndex, uint length, HSTRING* newString);
HRESULT WindowsConcatString(HSTRING string1, HSTRING string2, HSTRING* newString);
HRESULT WindowsReplaceString(HSTRING string, HSTRING stringReplaced, HSTRING stringReplaceWith, HSTRING* newString);
HRESULT WindowsTrimStringStart(HSTRING string, HSTRING trimString, HSTRING* newString);
HRESULT WindowsTrimStringEnd(HSTRING string, HSTRING trimString, HSTRING* newString);
HRESULT WindowsPreallocateStringBuffer(uint length, wchar** charBuffer, HSTRING_BUFFER* bufferHandle);
HRESULT WindowsPromoteStringBuffer(HSTRING_BUFFER bufferHandle, HSTRING* string);
HRESULT WindowsDeleteStringBuffer(HSTRING_BUFFER bufferHandle);