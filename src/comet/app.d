module comet.app;

import std.stdio : writeln;
import core.sys.windows.oaidl,
  core.sys.windows.objidl,
  core.sys.windows.ocidl,
  core.sys.windows.oleauto,
  core.sys.windows.objbase,
  core.sys.windows.wtypes,
  core.sys.windows.basetyps,
  core.sys.windows.windef,
  core.stdc.wchar_,
  core.sys.windows.unknwn,
  core.sys.windows.winrt.inspectable,
  core.sys.windows.winrt.hstring,
  core.sys.windows.winrt.winstring,
  core.sys.windows.winuser,
  core.sys.windows.commctrl,
  std.string,
  std.algorithm,
  std.variant,
  std.datetime,
  std.range,
  std.traits,
  comet.core,
  comet.attributes,
  comet.utils;

@guid("9E365E57-48B2-4160-956F-C7385120BBFC")
interface IUriRuntimeClass : IInspectable {
  HRESULT get_AbsoluteUri(HSTRING* value);
  HRESULT get_DisplayUri(HSTRING* value);
  HRESULT get_Domain(HSTRING* value);
  HRESULT get_Extension(HSTRING* value);
  HRESULT get_Fragment(HSTRING* value);
  HRESULT get_Host(HSTRING* value);
  HRESULT get_Password(HSTRING* value);
  HRESULT get_Path(HSTRING* value);
  HRESULT get_Query(HSTRING* value);
  HRESULT get_QueryParsed(void* ppWwwFormUrlDecoder);
  HRESULT get_RawUri(HSTRING* value);
  HRESULT get_SchemeName(HSTRING* value);
  HRESULT get_UserName(HSTRING* value);
  HRESULT get_Port(int* value);
  HRESULT get_Suspicious(bool* value);
  HRESULT Equals(IUriRuntimeClass pUri, bool* value);
  HRESULT CombineUri(HSTRING relativeUri, IUriRuntimeClass* instance);
}

@guid("44A9796F-723E-4FDF-A218-033E75B0C084")
interface IUriRuntimeClassFactory : IInspectable {
  HRESULT CreateUri(HSTRING uri, IUriRuntimeClass* instance);
  HRESULT CreateWithRelativeUri(HSTRING baseUri, HSTRING relativeUri, IUriRuntimeClass* instance);
}

mixin(RuntimeClassName!(IUriRuntimeClassFactory, "Windows.Foundation.Uri"));
mixin(RuntimeClassName!(IUriRuntimeClass, "Windows.Foundation.Uri"));

void main() {
  with (apartmentScope()) {
    /+auto u = activateWith!IUriRuntimeClassFactory("http://www.google.com");
    writeln(u.absoluteUri);+/

    UnknownPtr!IDispatch doc = make("MSXML2.DOMDocument");

    //doc.async = false;
    //doc.onreadystatechange += () => writeln("readyState change ", cast(int)doc.readyState);
    doc.loadXML(`<?xml version="1.0"?>
                 <books>
                   <book title="Weaveworld" author="Clive Barker">Nothing ever begins...</book>
                 </books>`);
    /+foreach (ref item; doc.selectNodes(`//book[@author="Clive Barker"]`)) {
      writeln(cast(string)item.nodeTypedValue);
    }+/

    /*UnknownPtr!IDispatch http = make("MSXML2.XMLHTTP");
    http.onReadyStateChange = () => writeln("readyState: ", cast(int)http.readyState);
    http.open("get", "http://httpbin.org/get");*/
  }
}