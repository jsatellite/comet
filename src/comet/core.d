///
module comet.core;

import core.sys.windows.w32api,
  core.sys.windows.core,
  core.sys.windows.basetyps,
  core.sys.windows.windef,
  core.sys.windows.wtypes,
  core.sys.windows.unknwn,
  core.sys.windows.objbase,
  core.sys.windows.objidlbase,
  core.sys.windows.oaidl,
  core.sys.windows.oleauto,
  core.sys.windows.ocidl,
  core.sys.windows.uuid,
  std.exception,
  std.array,
  std.string,
  std.utf,
  std.algorithm,
  std.range,
  std.traits,
  std.meta,
  std.variant,
  std.datetime,
  comet.attributes,
  comet.utils,
  comet.internal;
import core.stdc.wchar_ : wcslen;
import std.conv : to;
import std.typecons : tuple;
static if (_WIN32_WINNT >= 0x602) import core.sys.windows.winrt.roapi,
  core.sys.windows.winrt.hstring,
  core.sys.windows.winrt.winstring;
debug import std.stdio : writeln;

static if (_WIN32_WINNT >= 0x602) pragma(lib, "windowsapp");

/**
 The exception that is thrown when a COM operation fails.
 */
class HResultException : Exception {

  /**
   The error code.
   */
  HRESULT errorCode;

  /**
   Initializes a new instance with the specified error code.
   Params:
     errorCode = The error code associated with the exception.
   */
  this(HRESULT errorCode = E_FAIL, string file = __FILE__, size_t line = __LINE__) {
    string getMessage() {
      wchar* p;
      uint cch = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER |
                                FORMAT_MESSAGE_IGNORE_INSERTS, null, cast(uint)errorCode,
                                MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL), cast(wchar*)&p, 0, null);
      scope(exit) LocalFree(p);

      if (cch != 0) return p[0 .. cch].toUTF8().strip();
      else return "Unknown error.";
    }
    super("%s (0x%08X)".format(getMessage(), cast(uint)(this.errorCode = errorCode)), file, line);
  }

  /// ditto
  this(string message, HRESULT errorCode, string file = __FILE__, size_t line = __LINE__) {
    super("%s (0x%08X)".format(message, cast(uint)(this.errorCode = errorCode)), file, line);
  }

}

/// Indicates whether the COM operation succeeded.
pragma(inline, true) bool isSuccess(HRESULT status) { return status >= S_OK; }
/// ditto
pragma(inline, true) bool isSuccessOK(HRESULT status) { return status == S_OK; }
/// ditto
pragma(inline, true) bool isFailure(HRESULT status) { return status < S_OK; }

/**
 Checks the result of a COM operation, throwing an HResultException if it fails.
 Params:
   status = The _status code.
   file   = The source file of the caller.
   line   = The line number in the source file of the caller.
 Returns:
   The _status code.
 */
HRESULT checkHR(HRESULT status, string file = __FILE__, size_t line = __LINE__) {
  enforce(status.isSuccess, new HResultException(status, file, line));
  return status;
}

enum HRESULT E_BOUNDS = 0x8000000B;

/**
 Catches exceptions and translates them into HRESULT values.
 */
HRESULT catchHR(T)(T delegate() block) {
  import core.exception : OutOfMemoryError, RangeError;

  try {
         static if (is(T == HRESULT)) return block();
    else static if (is(T == bool))    return block() ? S_OK : S_FALSE;
    else static if (is(T== void))   { block(); return S_OK; }
    else static assert(false);
  }
  catch (HResultException e) { return e.errorCode; }
  catch (OutOfMemoryError)   { return E_OUTOFMEMORY; }
  catch (RangeError)         { return E_BOUNDS; }
  catch (Exception)          { return E_FAIL; }
}

/**
 The apartment model for the current thread.
 Remarks:
   Multithreaded apartments are intended for use by non-GUI threads.
 */
enum Apartment {
  singleThreaded, /// A single-threaded apartment.
  multiThreaded,  /// A multithreaded apartment.
  unknown,        /// The apartment model is not known.
}

/**
 Initializes COM and sets the _apartment model for the current thread.
 Params:
   apartment = The _apartment model.
 Remarks:
   This is a helper function that calls [RoInitialize](https://docs.microsoft.com/en-us/windows/win32/api/roapi/nf-roapi-roinitialize) 
   where Windows Runtime is available or [CoInitializeEx](https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-coinitializeex) otherwise.
 */
void initApartment(Apartment apartment = Apartment.singleThreaded) {
  static if (is(typeof(RoInitialize)))
    with (RO_INIT_TYPE) RoInitialize((apartment == Apartment.singleThreaded)
                                     ? RO_INIT_SINGLETHREADED : RO_INIT_MULTITHREADED).checkHR();
  else
    with (COINIT) CoInitializeEx(null, (apartment == Apartment.singleThreaded)
                                 ? COINIT_APARTMENTTHREADED : COINIT_MULTITHREADED).checkHR();
}

/**
 Closes COM on the current thread.
 */
pragma(inline, true) void uninitApartment() {
  static if (is(typeof(RoUninitialize)))
    RoUninitialize();
  else
    CoUninitialize();
}

/**
 Starts a scope wherein COM is initialized with the specified _apartment model.
 At the end of the scope, COM is closed on the current thread.
 Examples:
 ---
 new Thread({
   with (apartmentScope(Apartment.multiThreaded)) {
     // Perform tasks
   }
 }).start();
 ---
 */
auto apartmentScope(Apartment apartment = Apartment.multiThreaded) {

  static struct ApartmentScope {
    this(Apartment apartment) { initApartment(apartment); }
    ~this() { uninitApartment(); }
  }

  return ApartmentScope(apartment);
}

/**
 Returns a value indicating the current apartment model.
 */
Apartment apartmentState() @property {
  APTTYPE type;
  APTTYPEQUALIFIER qualifier;
  CoGetApartmentType(&type, &qualifier);

  switch (type) {
  case APTTYPE_STA, APTTYPE_MAINSTA:
    return Apartment.singleThreaded;
  case APTTYPE_MTA:
    return Apartment.multiThreaded;
  case APTTYPE_NA:
    switch (qualifier) {
    case APTTYPEQUALIFIER_NA_ON_STA, APTTYPEQUALIFIER_NA_ON_MAINSTA:
      return Apartment.singleThreaded;
    case APTTYPEQUALIFIER_NA_ON_MTA, APTTYPEQUALIFIER_IMPLICIT_MTA:
      return Apartment.multiThreaded;
    default:
      break;
    }
    break;
  default:
    break;
  }
  return Apartment.unknown;
}

/**
 Retrieves the GUID associated with the specified type or reference.
 Remarks:
   If null is supplied as the type, guidOf will return a GUID whose value is all zeros.
 Examples:
 ---
 immutable iid = guidOf!IXMLDOMDocument;
 ---
 */
template guidOf(alias T) 
  if (is(T : IUnknown) || is(typeof(T) == typeof(null))) {
  import core.sys.windows.uuid;

  enum typeName = __traits(identifier, T);

  static if (!is(typeof(T == null)) && __traits(compiles, .moduleName!T)) {
    enum moduleName = .moduleName!T;
    enum code = mixin(q{
      static import $moduleName;
           static if (is(typeof($moduleName.IID_$typeName)))   enum guidOf = $moduleName.IID_$typeName;
      else static if (is(typeof($moduleName.CLSID_$typeName))) enum guidOf = $moduleName.CLSID_$typeName;
      else static if (is(typeof($moduleName.DIID_$typeName)))  enum guidOf = $moduleName.DIID_$typeName;
      else static assert(false);
    }.inject);
  }

       static if (is(typeof(T == null)))                  enum guidOf = GUID.init;
  else static if (hasUDA!(T, GUID))                       enum guidOf = getUDAs!(T, GUID)[0];
  else static if (is(typeof(mixin("IID_" ~ typeName))))   mixin("enum guidOf = IID_" ~ typeName ~ ";");
  else static if (is(typeof(mixin("CLSID_" ~ typeName)))) mixin("enum guidOf = CLSID_" ~ typeName ~ ";");
  else static if (is(typeof(mixin("DIID_" ~ typeName))))  mixin("enum guidOf = DIID_" ~ typeName ~ ";");
  else static if (__traits(compiles, { mixin(code); }))   mixin(code);
  else static                                             assert(false, "no GUID has been associated with `" ~ typeName ~ "`");
}

/**
 Creates a new, globally unique identifier.
 Returns:
   A new GUID.
 */
GUID makeGUID() {
  GUID g;
  checkHR(CoCreateGuid(&g));
  return g;
}

/**
 Returns the requested interface if it is supported.
 Returns:
   A reference to the requested interface if supported. Otherwise, null.
 Remarks:
   This is a helper function that calls [QueryInterface](https://docs.microsoft.com/en-us/windows/desktop/api/unknwn/nf-unknwn-iunknown-queryinterface(refiid_void)).
 */
T tryAs(T)(IUnknown source, IID targetId)
  if (is(T : IUnknown)) {
  T target;
  source.QueryInterface(&targetId, cast(void**)&target);
  return target;
}

/// ditto
T tryAs(T)(IUnknown source)
  if (is(T : IUnknown)) {
  return tryAs!T(source, guidOf!T);
}

enum ExecutionContext : uint {
  inProcessServer = CLSCTX.CLSCTX_INPROC_SERVER,
  inProcessHandler = CLSCTX.CLSCTX_INPROC_HANDLER,
  localServer = CLSCTX.CLSCTX_LOCAL_SERVER,
  remoteServer = CLSCTX.CLSCTX_REMOTE_SERVER,
  inProcess = CLSCTX_INPROC,
  server = CLSCTX_SERVER,
  all = CLSCTX_ALL
}

alias defaultExecutionContext = ExecutionContext.all;

pragma(inline, true) private bool tryMake(T)(CLSID classId, IID interfaceId, ExecutionContext context, out T result) {
  return CoCreateInstance(&classId, null, context, &interfaceId, cast(void**)&result).isSuccessOK;
}

/**
 Creates an instance of the COM class associated with the specified identifier.
 Params:
   classId     = The identifier of the class.
   interfaceId = The identifier of the interface.
   context     = The _context in which the instance will run.
 Remarks:
   This is a helper function that calls [CoCreateInstance](https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance).
 */
auto make(T)(CLSID classId, IID interfaceId, ExecutionContext context = defaultExecutionContext)
if (is(T : IUnknown)) {
  T result;
  if ((context & ExecutionContext.inProcessServer) &&
      tryMake(classId, interfaceId, context, result))
    return result;
  if ((context & ExecutionContext.inProcessHandler) &&
      tryMake(classId, interfaceId, context, result))
    return result;
  if ((context & ExecutionContext.localServer) &&
      tryMake(classId, interfaceId, context, result))
    return result;
  if ((context & ExecutionContext.remoteServer) &&
      tryMake(classId, interfaceId, context, result))
    return result;
  if (tryMake(classId, interfaceId, ExecutionContext.all, result))
    return result;
  return null;
}

/// ditto
auto make(T = IUnknown)(CLSID classId, ExecutionContext context = defaultExecutionContext) {
  return make!T(classId, guidOf!T, context);
}

/// ditto
auto make(T = IUnknown)(string classId, ExecutionContext context = defaultExecutionContext) {
  GUID clsid;
  if (!CLSIDFromProgID(classId.toUTFz!(wchar*), &clsid).isSuccessOK) {
    try clsid = guid(classId);
    catch return null;
  }
  return make!T(clsid, guidOf!T, context);
}

private template hasMember(T, string name) {
  static if (__traits(hasMember, T, name)) {
    enum comet_core_hasMember_result = true;
  } else static foreach (m; __traits(allMembers, T)) {
    static if (!is(typeof(comet_core_hasMember_result))) {
      static if (hasUDA!(__traits(getMember, T, m), overload) &&
                 getUDAs!(__traits(getMember, T, m), overload)[0].method == name) {
        enum comet_core_hasMember_result = true;
      }
    }
  }

  static if (!is(typeof(comet_core_hasMember_result))) enum hasMember = false;
  else                                                 enum hasMember = true;
}

pragma(inline) private void checkNotNull(T)(T ptr, string argument) {
  enforce(ptr !is null, "`" ~ argument ~ "` was null");
}

/**
 Represents a pointer to the interface specified by the template parameter.
 */
struct UnknownPtr(T)
  if (is(T : IUnknown)) {

  private T ptr_;

  private uint internalAddRef() {
    if (ptr_ !is null)
      return ptr_.AddRef();
    return 0;
  }

  private uint internalRelease() {
    if (ptr_ !is null) {
      scope(exit) ptr_ = null;
      return ptr_.Release();
    }
    return 0;
  }

  /**
   Initializes a new instance with the specified data.
   */
  this(U)(U other) if (is(U : T)) { ptr_ = other; }

  /// ditto
  this(U)(auto ref U other)
    if (!is(U : T) &&
        !is(U == typeof(this)) &&
        !is(U : VARIANT) &&
        !is(U == typeof(null))) {
    ptr_ = tryAs!T(other);
  }

  /// ditto
  this(U)(auto ref U other)
    if (is(U : VARIANT)) {
    with (VARENUM) if (other.vt == VT_DISPATCH) ptr_ = tryAs!T(other.pdispVal);
    else           if (other.vt == VT_UNKNOWN)  ptr_ = tryAs!T(other.punkVal);
  }

  /// ditto
  this(U)(U)
    if (is(U == typeof(null))) {}

  /// ditto
  this(ref typeof(this) other) {
    ptr_ = other.ptr_;
    internalAddRef();
  }

  ~this() {
    internalRelease();
  }

  /**
   Returns the underlying raw pointer.
   */
  ref T ptr() @property { return ptr_; }
  alias ptr this;

  // Can't overload address operator (&) so repurpose dereference operator instead
  T* opUnary(string op : "*")() {
    internalRelease();
    return &ptr_;
  }

  ref opAssign(U)(auto ref U other)
    if (!is(U == T) &&
        !is(U == typeof(this)) &&
        !is(U : VARIANT) &&
        !is(U == typeof(null))) {
    internalRelease();
    ptr_ = tryAs!T(other);
    return this;
  }

  ref opAssign(U)(auto ref U other)
    if (is(U : VARIANT)) {
    internalRelease();
    with (VARIANT) if (other.vt == VT_DISPATCH) ptr_ = tryAs!T(other.pdispVal);
    else           if (other.vt == VT_UNKNOWN)  ptr_ = tryAs!T(other.punkVal);
    return this;
  }

  ref opAssign(U)(U)
    if (is(U == typeof(null))) {
    internalRelease();
    return this;
  }

  U opCast(U)()
    if (is(U == T)) { return ptr_; }

  U opCast(U)()
    if (!is(U == T) &&
        !is(U == Pointer!A, A) &&
        !is(U : VARIANT) &&
        !is(U == bool)) { return tryAs!U(ptr_); }

  U opCast(U)()
    if (is(U == Pointer!A, A)) { return Pointer!(TemplateArgsOf!U)(ptr_); }

  bool opCast(U)()
    if (is(U == bool)) { return !!ptr_; }

  bool opEquals(U)(U)
    if (is(U == typeof(null))) { return !ptr_; }

  /**
   Forwards calls to the underlying pointer.
   */
  template opDispatch(string name) {
    enum isCamelCase = name.isCamelCase;
    enum realName = name.toPascalCase();

    enum getPrefix = "get_", setPrefix = "put_", setRefPrefix = "putref_";
    enum addPrefix = "add_", removePrefix = "remove_";

    // For property accessors
    enum getName = getPrefix ~ name, setName = setPrefix ~ name, setRefName = setRefPrefix ~ name;
    enum realGetName = getPrefix ~ realName, realSetName = setPrefix ~ realName, realSetRefName = setRefPrefix ~ realName;

    // For event operations
    enum addName = addPrefix ~ name, removeName = removePrefix ~ name;
    enum realAddName = addPrefix ~ realName, realRemoveName = removePrefix ~ realName;

    auto opDispatchInner(Args...)(auto ref Args args) {
             static if (!__traits(hasMember, ptr_, name) && isCamelCase && hasMember!(T, realName)) {
        return opDispatch!realName(args);
      } else static if (!__traits(hasMember, ptr_, name) && isCamelCase && hasMember!(T, realGetName) && args.length == 0) {
        return opDispatch!realGetName();
      } else static if (!__traits(hasMember, ptr_, name) && isCamelCase && hasMember!(T, realSetName) && args.length == 1) {
        return opDispatch!realSetName(args);
      } else static if (!__traits(hasMember, ptr_, name) && isCamelCase && hasMember!(T, realSetRefName) && args.length == 1) {
        return opDispatch!realSetRefName(args);
      } else static if (!__traits(hasMember, ptr_, name) && isCamelCase &&
                        (__traits(hasMember, ptr_, realAddName) || __traits(hasMember, ptr_, realRemoveName)) && args.length == 0) {
        return opDispatchEventImpl!realName(ptr_);
      } else static if (__traits(hasMember, ptr_, getName) && args.length == 0) {
        return opDispatch!getName();
      } else static if (__traits(hasMember, ptr_, setName) && args.length == 1) {
        return opDispatch!setName(args);
      } else static if (__traits(hasMember, ptr_, setRefName) && args.length == 1) {
        return opDispatch!setRefName(args);
      } else static if ((__traits(hasMember, ptr_, addName) || __traits(hasMember, ptr_, removeName)) && args.length == 0) {
        return opDispatchEventImpl!name(ptr_);
      } else static if (__traits(compiles, opDispatchImpl!name(ptr_, args))) {
        return opDispatchImpl!name(ptr_, args);
      } else static if (__traits(compiles, opDispatchTransformImpl!name(ptr_, args))) {
        return opDispatchTransformImpl!name(ptr_, args);
      } else static if (is(T : IDispatch)) {
        return opDispatchInvokeImpl(ptr_, name, args);
      } else {
        static assert(false, opDispatchError!(name, T, Args));
      }
    }
    alias opDispatch = opDispatchInner;
  }

  static if (isEnumerator!T) {
    private alias U = PointerTarget!(Parameters!(T.Next)[1]);

         static if (is(U == VARIANT)) alias V = _Variant;
    else static if (is(U == wchar*))  alias V = string;
    else                              alias V = U;

    int opApply(scope int delegate(ref V) it)         { return opApplyEnumerator(ptr_, it, null); }
    int opApply(scope int delegate(size_t, ref V) it) { return opApplyEnumerator(ptr_, null, it); }
  } else static if (hasEnumerator!T) {
    private enum enumerator = enumeratorName!T;

    private alias E = PointerTarget!(Parameters!(mixin("T." ~ enumerator))[0]);
    static assert(isEnumerator!E);

    private alias U = PointerTarget!(Parameters!(E.Next)[1]);

    static if (is(U == VARIANT)) alias V = _Variant;
    else                         alias V = U;

    int opApply(scope int delegate(ref V) it)         { return opApplyImpl(it, null); }
    int opApply(scope int delegate(size_t, ref V) it) { return opApplyImpl(null, it); }

    private int opApplyImpl(scope int delegate(ref V) it1, scope int delegate(size_t, ref V) it2) {
      E e;
      mixin("auto hr = ptr_." ~ enumerator ~ "(cast(E*)&e);");
      if (e !is null && hr.isSuccess) {
        try return opApplyEnumerator(e, it1, it2);
        finally e.Release();
      }
      return 0;
    }
  } else static if (is(T : IDispatch)) {
    int opApply(scope int delegate(ref _Variant) it)         { return opApplyInvokeImpl(ptr_, it, null); }
    int opApply(scope int delegate(size_t, ref _Variant) it) { return opApplyInvokeImpl(ptr_, null, it); }
  }

  static if (hasGetIndexer!T) {
    private enum getIndexer = getIndexerName!T;
    auto opIndex(TIndex)(TIndex index) { return opDispatch!getIndexer(index); }
  } else static if (is(T : IDispatch)) {
    auto ref opIndex(TIndices...)(auto ref TIndices) { return opIndexInvokeImpl(ptr_, indices); }
  }

  static if (hasSetIndexer!T) {
    private enum setIndexer = setIndexerName!T;
    void opIndexAssign(TValue, TIndex)(TValue value, TIndex index) { opDispatch!setIndexer(index, value); }
  } else static if (is(T : IDispatch)) {
    void opIndexAssign(T, TIndices...)(auto ref T value, auto ref TIndices indices) {
      opIndexAssignInvokeImpl(ptr_, value, indices);
    }
  }

}

/+private class OpApplyRange(T) {

  private int delegate(int delegate(ref T)) opApply_;

  this(int delegate(int delegate(ref T)) opApply) {
    opApply_ = opApply;
    popFront();
  }

  bool empty;
  T front;
  void popFront() { empty = (opApply_((ref value) {
    front = value;
    return 1;
  })) == 0; }

}+/

auto asUnknownPtr(T)(T source)
  if (is(T : IUnknown)) { return UnknownPtr!T(source); }

auto asUnknownPtr(T, U)(auto ref U source)
  if (is(T : IUnknown)) { return UnknownPtr!T(source); }

private auto opDispatchImpl(string name, T, Args...)(T ptr, ref Args args)
  if (is(T : IUnknown)) {
  checkNotNull(ptr, nameOf!ptr);

  static if (__traits(hasMember, ptr, name)) {
    return mixin("ptr." ~ name ~ "(args)");
  } else {
    static foreach (m; __traits(allMembers, T)) {
      static if (hasUDA!(__traits(getMember, ptr, m), overload) &&
                 getUDAs!(__traits(getMember, ptr, m), overload)[0].method == name) {
        static if (__traits(compiles, mixin("ptr." ~ m ~ "(args)"))) {
          enum opDispatchImplOverloadFound = true;
          return mixin("ptr." ~ m ~ "(args)");
        }
      }
    }
    static if (!is(typeof(opDispatchImplOverloadFound)))
      static assert(false, opDispatchError!(name, T, Args));
  }
}

pragma(inline, true) bool validateBSTR(wchar* s, out size_t length) {
  // SysStringLen only works for BSTR, while wcslen works for any BSTR/LPWSTR/wchar* etc.
  // So if wcslen and SysStringLen return the same value, it's a BSTR.
  return (length = (s != null) ? wcslen(s) : 0) == SysStringLen(s);
}

private void generateOpDispatchTransform(alias member, T, Args...)(T ptr, ref Args args) {
  checkNotNull(ptr, nameOf!ptr);

  alias Return = ReturnType!member,
        Params = Parameters!member;

  enum hasReturnParam = Params.length == Args.length + 1 && isPointer!(Params[$ - 1]);
  static if (hasReturnParam) alias ReturnParam = PointerTarget!(Params[$ - 1]);

  enum memberName = __traits(identifier, member);

  static string generateMarshaling() {
    string code;

    static foreach (i, arg; args) {{
      alias Source = Unqual!(typeof(arg)), Target = Params[i];

      static if (// string types
                 // Both BSTR and LPWSTR are aliased to wchar* so we've no way of differentiating, but this is COM
                 // so we prefer BSTR.
                 (isString!Source && (is(Target == BSTR) || is(Target == LPCWSTR) ||
                                      (is(HSTRING) && is(Target == HSTRING)))) ||
                 (isPointer!Source && isString!(PointerTarget!Source) && (is(Target == BSTR*) ||
                                                                          is(Target == LPCWSTR*) ||
                                                                          (is(HSTRING) && is(Target == HSTRING*)))) ||
                 
                 // bool types
                 (is(Source == bool) && (is(Target == BOOL) || is(Target == VARIANT_BOOL))) ||
                 (is(Source == bool*) && (is(Target == BOOL*) || is(Target == VARIANT_BOOL*))) ||
                 
                 // where a VARIANT is expected allow any type to be passed
                 (!is(Source : VARIANT) && is(Target == VARIANT)) ||
                 (isPointer!Source && !is(PointerTarget!Source : VARIANT) && is(Target == VARIANT*)) ||
                 
                 // arrays
                 (isDynamicArray!Source && !isString!Source && (is(Target == SAFEARRAY*) ||
                                                                is(PointerTarget!Target == ElementEncodingType!Source))) ||
                 (isPointer!Source && isDynamicArray!(PointerTarget!Source) && is(Target == SAFEARRAY**)) ||
                 
                 // u/long -> U/LARGE_INTEGER
                 (is(Source == long) && is(Target == LARGE_INTEGER)) ||
                 (is(Source == ulong) && is(Target == ULARGE_INTEGER)) ||
                 (is(Source == long*) && is(Target == LARGE_INTEGER*)) ||
                 (is(Source == ulong*) && is(Target == ULARGE_INTEGER*)) ||

                 // delegate/function -> IDispatch callback
                 ((isDelegate!Source || isFunctionPointer!Source) && is(Target == IDispatch)) ||
                 
                 // IIDs by value
                 (is(Source == GUID) && is(Target == IID*)) ) {
             code ~= mixin(opDispatchMarshalParam!Target.inject);
      } else static if (isInstanceOf!(TaskBuffer, Source) && isPointer!Target) {
             code ~= mixin(opDispatchMarshalParam!Source.inject);
      } else code ~= mixin(opDispatchMarshalParam!().inject);
    }}

    static if (hasReturnParam) code ~= "ReturnParam returnParam;
      ";
    static if (!is(Return == void)) code ~= "auto returnValue = ";

    code ~= "ptr." ~ memberName ~ "(";
    auto params = iota(Args.length).map!(i => mixin("param$i".inject));
    code ~= choose(!hasReturnParam, params, params.chain("&returnParam".only))
      .join(", ");
    code ~= ");
      ";

    static if (is(Return == HRESULT)) code ~= "scope(exit) checkHR(returnValue);
      ";

    static if (hasReturnParam) {
      // VARIANT_BOOL is aliased to `short` and BOOL to `int`, meaning we can't disambiguate, so leave bool conversion to user
      // Workaround is for user to pass an address of a bool as last parameter
      static if (is(ReturnParam == BSTR) ||
                 (is(HSTRING) && is(ReturnParam == HSTRING)) ||
                 is(ReturnParam == LARGE_INTEGER) || is(ReturnParam == ULARGE_INTEGER) ||
                 is(ReturnParam == VARIANT) ||
                 is(ReturnParam : IUnknown)) {
             code ~= opDispatchMarshalReturn!ReturnParam;
      } else code ~= "return returnParam;";
    } else static if (!is(Return == void)) {
      code ~= "return returnValue;";
    }

    return code;
  }

  enum code = generateMarshaling;
  //pragma(msg, code);
  mixin(code);
}

private void opDispatchTransformImpl(string name, T, Args...)(T ptr, ref Args args)
  if (is(T : IUnknown)) {
  static if (__traits(hasMember, ptr, name)) {
    alias overloads = __traits(getOverloads, T, name);

    static foreach (member; overloads) {
      static if (__traits(compiles, generateOpDispatchTransform!member(ptr, args))) {
        static if (!is(typeof(opDispatchTransformImplOverloadFound))) {
          enum opDispatchTransformImplOverloadFound = true;

          generateOpDispatchTransform!member(ptr, args);
        }
      }
    }
  }

  static if (!is(typeof(opDispatchTransformImplOverloadFound))) static assert(false);
}

private auto opDispatchInvokeImpl(Args...)(IDispatch ptr, string name, ref Args args) {
  checkNotNull(ptr, nameOf!ptr);

  static TypeDescription[GUID] cache;

  // Bypass IConnectionPoint... event malarky for delegates assigned directly
  static if (!(args.length == 1 && (isDelegate!args || isFunctionPointer!args))) {
    if (auto type = scanForEvents(ptr, cache)) {
      if (!type.events.empty) {
        if (auto event = type.events.get(name.toLower(), null))
          return _Variant(ptr, event.dispId, event.sourceIid);
      }
    }
  }

  int dispId = DISPID_UNKNOWN;
  ushort flags;

  if (auto type = scanForMethods(ptr, cache)) {
    if (!type.setters.empty) {
      if (auto setter = type.setters.get(name.toLower(), null)) {
        // Getters and setters have the same name, disambiguate by comparing the parameter count
        if (args.length == setter.paramCount) {
          dispId = setter.dispId;
          flags = cast(ushort)setter.invokeKind;
        }
      }
    }

    if (dispId == DISPID_UNKNOWN && !type.methods.empty) {
      if (auto method = type.methods.get(name.toLower(), null)) {
        dispId = method.dispId;
        flags = cast(ushort)method.invokeKind;
      }
    }
  }

  if (dispId == DISPID_UNKNOWN) {
    // Fallback if the type library lookup yielded nothing
    auto szName = name.toUTFz!(wchar*);
    ptr.GetIDsOfNames(&IID_NULL, &szName, 1, LOCALE_USER_DEFAULT, &dispId);
    flags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
  }

  return ptr.invokeById(dispId, flags, args);
}

auto invokeById(Args...)(IDispatch ptr, int dispId, ushort flags, ref Args args) {
  checkNotNull(ptr, nameOf!ptr);
  enforce(dispId != DISPID_UNKNOWN, new HResultException(DISP_E_UNKNOWNNAME));

  VARIANT[args.length] arguments;
  static if (args.length > 0) {
    foreach (i, arg; args)
      variantPut(arguments[$ - i - 1], arg);
  }

  DISPPARAMS params = { cArgs: cast(uint)args.length, rgvarg: arguments.ptr };
  VARIANT result;
  EXCEPINFO ex;
  HRESULT hr;

  if ((flags & DISPATCH_PROPERTYPUT) == 0) {
    if ((hr = ptr.Invoke(dispId, &IID_NULL, LOCALE_USER_DEFAULT, flags, &params, &result, &ex, null)) == DISP_E_MEMBERNOTFOUND)
      // If the method or getter weren't found try invoking the setter
      flags = DISPATCH_PROPERTYPUT;
  }

  if (flags & DISPATCH_PROPERTYPUT) {
    immutable dispIdNamed = DISPID_PROPERTYPUT;
    params.cNamedArgs = 1;
    params.rgdispidNamedArgs = cast(int*)&dispIdNamed;

    hr = ptr.Invoke(dispId, &IID_NULL, LOCALE_USER_DEFAULT, flags, &params, &result, &ex, null);
  }

  scope(exit) variantClear(result);

  static if (args.length > 0) {
    foreach (i; 0 .. args.length)
      variantClear(params.rgvarg[i]);
  }

  string message;
  if (hr == DISP_E_EXCEPTION && (ex.scode != 0 || ex.wCode != 0)) {
    message = ex.bstrDescription[0 .. SysStringLen(ex.bstrDescription)].toUTF8();
    SysFreeString(ex.bstrDescription);
    hr = (ex.scode != 0) ? ex.scode : ex.wCode;
  }

  switch (hr) {
  case S_OK, S_FALSE:
    return _Variant(result);
  case E_ABORT:
    return _Variant.init;
  default:
    throw !message.empty ? new HResultException(message, hr) : new HResultException(hr);
  }
}

private int opApplyEnumerator(E, T)(E enumerator, int delegate(ref T) it1, int delegate(size_t, ref T) it2) {
  checkNotNull(enumerator, nameOf!enumerator);

  import core.stdc.wchar_ : wcslen;

  PointerTarget!(Parameters!(E.Next)[1]) item;
  uint count;
  size_t index;

  while (enumerator.Next(1, &item, &count).isSuccessOK && count > 0) {
         static if (is(T == _Variant))                        T p = item;
    else static if (isString!T && is(typeof(item) == wchar*)) T p = item[0 .. wcslen(item)].toUTF8();
    else                                                      alias p = item;

    if (it1 !is null) {
      if (auto r = it1(p)) return r;
    } else if (it2 !is null) {
      if (auto r = it2(index++, p)) return r;
    }

    static if (is(typeof(item) == wchar*)) CoTaskMemFree(item);
  }

  return 0;
}

private int opApplyInvokeImpl(T)(IDispatch ptr, int delegate(ref T) it1, int delegate(size_t, ref T) it2) {
  checkNotNull(ptr, nameOf!ptr);

  DISPPARAMS params;
  VARIANT result;
  immutable iid = guidOf!null;

  if (ptr.Invoke(DISPID_NEWENUM, &iid, LOCALE_USER_DEFAULT, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                 &params, &result, null, null).isSuccessOK) {
    scope(exit) variantClear(result);

    if (auto e = result.pdispVal.tryAs!IEnumVARIANT()) {
      try return opApplyEnumerator(e, it1, it2);
      finally e.Release();
    }
  }

  return 0;
}

private auto opIndexInvokeImpl(TIndices...)(IDispatch ptr, ref TIndices indices) {
  checkNotNull(ptr, nameOf!ptr);

  VARIANT[indices.length] args;
  foreach (i, index; indices)
    variantPut(args[$ - i - 1], index);

  DISPPARAMS params = { cArgs: cast(uint)indices.length, rgvarg: args.ptr };
  VARIANT result;
  immutable iid = guidOf!null;

  checkHR(ptr.Invoke(DISPID_VALUE, &iid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &result, null, null));
  scope(exit) variantClear(result);

  foreach (i; 0 .. indices.length)
    variantClear(args[i]);

  return _Variant(result);
}

private void opIndexAssignInvokeImpl(T, TIndices...)(IDispatch ptr, ref T value, ref TIndices indices) {
  checkNotNull(ptr, nameOf!ptr);

  if (auto item = opIndexInvokeImpl(ptr, indices)) {
    immutable dispIdNamed = DISPID_PROPERTYPUT;

    auto arg = variantInit(value);
    scope(exit) variantClear(arg);

    DISPPARAMS params = {
      cArgs: 1, 
      rgvarg: &arg,
      cNamedArgs: 1, 
      rgdispidNamedArgs: cast(int*)&dispIdNamed
    };
    immutable iid = guidOf!null;
    checkHR(item.pdispVal.Invoke(DISPID_VALUE, &iid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, null, null, null));
  }
}

private auto opDispatchEventImpl(string name, T)(T ptr) {
  checkNotNull(ptr, nameOf!ptr);
  static if (__traits(hasMember, ptr, "add_" ~ name) && 
             __traits(hasMember, ptr, "remove_" ~ name)) {
    alias TEventHandler = Parameters!(__traits(getMember, ptr, "add_" ~ name))[0];
    alias TToken = PointerTarget!(Parameters!(__traits(getMember, ptr, "add_" ~ name))[1]);
    return EventOpHelper!(TEventHandler, TToken)(&__traits(getMember, ptr, "add_" ~ name),
                                                 &__traits(getMember, ptr, "remove_" ~ name));
  } else {
    return opDispatchImpl!name(ptr);
  }
}

enum isString(T)                 = is(T == string) || is(T == wstring) || is(T == dstring);
private enum isNonCharArray(T)   = isArray!T && !isString!T;
private enum isNonCharPointer(T) = isPointer!T && !is(Unqual!T == wchar*);

template VariantType(T)
  if (!isNonCharArray!T &&
      !isNonCharPointer!T &&
      !is(T == typeof(null)) &&
      !is(T == enum) &&
      !is(T == Variant)) {
  enum VariantType = match!(Unqual!T,
    long, () => VARENUM.VT_I8,
    int, () => VARENUM.VT_I4,
    ubyte, () => VARENUM.VT_UI1,
    short, () => VARENUM.VT_I2,
    float, () => VARENUM.VT_R4,
    double, () => VARENUM.VT_R8,
    bool, () => VARENUM.VT_BOOL,
    DateTime, () => VARENUM.VT_DATE,
    isString!T || is(T == wchar*), () => VARENUM.VT_BSTR,
    is(T : IDispatch), () => VARENUM.VT_DISPATCH,
    is(T : IUnknown), () => VARENUM.VT_UNKNOWN,
    is(T == char) || is(T == byte), () => VARENUM.VT_I1,
    ushort, () => VARENUM.VT_UI2,
    uint, () => VARENUM.VT_UI4,
    ulong, () => VARENUM.VT_UI8,
    is(T : DECIMAL), () => VARENUM.VT_DECIMAL,
    is(T : VARIANT), () => VARENUM.VT_VARIANT,
    () => VARENUM.VT_VOID);
}

unittest {
  assert(VariantType!long == VARENUM.VT_I8);
  assert(VariantType!int == VARENUM.VT_I4);
  assert(VariantType!ubyte == VARENUM.VT_UI1);
  assert(VariantType!short == VARENUM.VT_I2);
  assert(VariantType!float == VARENUM.VT_R4);
  assert(VariantType!double == VARENUM.VT_R8);
  assert(VariantType!bool == VARENUM.VT_BOOL);
  assert(VariantType!DateTime == VARENUM.VT_DATE);
  assert(VariantType!char == VARENUM.VT_I1);
  assert(VariantType!byte == VARENUM.VT_I1);
  assert(VariantType!ushort == VARENUM.VT_UI2);
  assert(VariantType!uint == VARENUM.VT_UI4);
  assert(VariantType!ulong == VARENUM.VT_UI8);
  assert(VariantType!DECIMAL == VARENUM.VT_DECIMAL);
  assert(VariantType!VARIANT == VARENUM.VT_VARIANT);
  assert(VariantType!void == VARENUM.VT_VOID);
}

template VariantType(T)
  if (isNonCharArray!T ||
      isNonCharPointer!T ||
      is(T == typeof(null)) ||
      is(T == enum)) {
       static if (is(T == typeof(null))) enum VariantType = VARENUM.VT_NULL;
  else static if (is(T E == enum))       enum VariantType = VariantType!E;
  else static if (isArray!T)             enum VariantType = VARENUM.VT_ARRAY | VariantType!(ElementEncodingType!T);
  else static if (isPointer!T)           enum VariantType = VARENUM.VT_BYREF | VariantType!(PointerTarget!T);
}

unittest {
  assert(VariantType!(typeof(null)) == VARENUM.VT_NULL);
  enum Suit : ubyte { spades, hearts, clubs, diamonds }
  assert(VariantType!Suit == VARENUM.VT_UI1);
  assert(VariantType!Suit != VARENUM.VT_I4);
  assert(VariantType!(Suit[]) == (VARENUM.VT_ARRAY | VARENUM.VT_UI1));
  assert(VariantType!(Suit[]) != (VARENUM.VT_ARRAY | VARENUM.VT_I4));
  assert(VariantType!(Suit*) == (VARENUM.VT_BYREF | VARENUM.VT_UI1));
  assert(VariantType!(Suit*) != (VARENUM.VT_BYREF | VARENUM.VT_I4));
}

void variantPut(T)(ref VARIANT target, T value)
  if (!isNonCharArray!T &&
      !isDelegate!T &&
      !isFunctionPointer!T &&
      !is(T == Variant)) {
  variantPut(target, value, (target.vt != VARENUM.VT_EMPTY) ? target.vt : VariantType!T);
}

void variantPut(T)(ref VARIANT target, T value, int type)
  if (!isNonCharArray!T &&
      !isDelegate!T &&
      !isFunctionPointer!T &&
      !is(T == Variant)) {
  with (target) {
    if (vt != VARENUM.VT_EMPTY)
      variantClear(target);

         static if (is(T == typeof(null)))       byref = null;
    else static if (is(T : VARIANT))             target = value;
    else static if (is(T == long))               llVal = value;
    else static if (is(T == int))                lVal = value;
    else static if (is(T == ubyte))              bVal = value;
    else static if (is(T == short))              iVal = value;
    else static if (is(T == float))              fltVal = value;
    else static if (is(T == double))             dblVal = value;
    else static if (is(T == bool))               boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    else static if (is(T == DateTime))         { SYSTEMTIME temp = {
                                                   value.year, value.month, value.day, value.dayOfWeek,
                                                   value.hour, value.minute, value.second
                                                 };
                                                 SystemTimeToVariantTime(&temp, &date); }
    else static if (isString!T)                { auto temp = value.toUTF16();
                                                 bstrVal = SysAllocStringLen(temp.ptr, cast(uint)temp.length); }
    else static if (is(T : IDispatch))         { if ((pdispVal = value) !is null) pdispVal.AddRef(); }
    else static if (is(T : IUnknown))          { if ((punkVal = value) !is null) punkVal.AddRef(); }
    else static if (is(T : Object))              byref = cast(void*)value;
    else static if (is(T == char))               cVal = value;
    else static if (is(T == ushort))             uiVal = value;
    else static if (is(T == uint))               ulVal = value;
    else static if (is(T == ulong))              ullVal = value;
    else static if (is(T : DECIMAL))             decVal = value;
    else static if (is(T == ubyte*))             pbVal = value;
    else static if (is(T == short*))             puiVal = value;
    else static if (is(T == int*))               plVal = value;
    else static if (is(T == long*))              pllVal = value;
    else static if (is(T == float*))             pfltVal = value;
    else static if (is(T == double*))            pdecVal = value;
    else static if (is(T == bool*))            { auto temp = *value ? VARIANT_TRUE : VARIANT_FALSE; 
                                                 pboolVal = &temp; }
    else static if (is(typeof(*T) == string))  { auto temp = (*value).toUTF16();
                                                 auto pTemp = SysAllocStringLen(temp.ptr, cast(uint)temp.length);
                                                 pbstrVal = &pTemp; }
    else static if (is(typeof(*T) : IDispatch))  ppdispVal = value;
    else static if (is(typeof(*T) : IUnknown))   ppunkVal = value;
    else static if (is(T == VARIANT*))           pvarVal = value;
    else static if (is(T == DECIMAL*))           pdecVal = value;
    else static if (is(T == char*))              pcVal = value;
    else static if (is(T == ushort*))            puiVal = value;
    else static if (is(T == uint*))              pulVal = value;
    else static if (is(T == ulong*))             pullVal = value;
    else static                                  assert(false, "Unsupported type `" ~ nameOf!T ~ "`");

    if ((vt = VariantType!T) != type) {
      with (VARENUM) if ((type == VT_ERROR && vt == VT_I4) ||
                         (type == VT_BOOL && vt == VT_I2) ||
                         (type == VT_UINT && vt == VT_UI4) ||
                         (type == VT_INT && vt == VT_I4) ||
                         (type == VT_DATE && vt == VT_R8) ||
                         (type == VT_CY && vt == VT_I8))
        vt = cast(ushort)type;
      else
        VariantChangeTypeEx(&target, &target, LOCALE_USER_DEFAULT, 0, cast(ushort)type);
    }
  }
}

void variantPut(T)(ref VARIANT target, T value)
  if (isNonCharArray!T) {
  with (target) {
    if (vt != VARENUM.VT_EMPTY)
      variantClear(target);

    vt = VariantType!T;
    parray = value.safeArray;
  }
}

void variantPut(T)(ref VARIANT target, T value)
  if (isDelegate!T || isFunctionPointer!T) {
  with (target) {
    if (vt != VARENUM.VT_EMPTY)
      variantClear(target);

    vt = VARENUM.VT_DISPATCH;
    pdispVal = new class IDispatch {
      mixin Dispatch;

      @dispId(DISPID_VALUE)
      extern(D) ReturnType!T invoke(Parameters!T args) { return value(args); }
    };
  }
}

VARIANT variantInit(T)(T value)
  if (!isNonCharArray!T &&
      !isDelegate!T &&
      !isFunctionPointer!T &&
      !is(T == Variant)) {
  VARIANT target;
  variantPut(target, value);
  return target;
}

VARIANT variantInit(T)(T value, int type)
  if (!isNonCharArray!T &&
      !isDelegate!T &&
      !isFunctionPointer!T &&
      !is(T == Variant)) {
  VARIANT target;
  variantPut(target, value, type);
  return target;
}

VARIANT variantInit(T)(T value)
  if (isNonCharArray!T ||
      isDelegate!T || 
      isFunctionPointer!T) {
  VARIANT target;
  variantPut(target, value);
  return target;
}

T variantGet(T)(auto ref VARIANT source, lazy T defaultValue = T.init)
  if (!isNonCharArray!T &&
      !is(T == Variant)) {
  with (source) with (VARENUM) {
         static if (is(T == long))       return (vt == VT_I8) ? llVal : defaultValue;
    else static if (is(T == int))        return (vt == VT_I4) ? lVal : defaultValue;
    else static if (is(T == ubyte))      return (vt == VT_UI1) ? bVal : defaultValue;
    else static if (is(T == short))      return (vt == VT_I2) ? iVal : defaultValue;
    else static if (is(T == float))      return (vt == VT_R4) ? fltVal : defaultValue;
    else static if (is(T == double))     return (vt == VT_R8) ? dblVal : defaultValue;
    else static if (is(T == bool))       return (vt == VT_BOOL) ? (boolVal != VARIANT_FALSE) : defaultValue;
    else static if (is(T == DateTime)) { if (vt != VT_DATE) return defaultValue;
                                         SYSTEMTIME temp;
                                         VariantTimeToSystemTime(date, &temp);
                                         return DateTime(temp.wYear, temp.wMonth, temp.wDay, 
                                                           temp.wHour, temp.wMinute, temp.wSecond); }
    else static if (isString!T)        { if (vt != VT_BSTR) return defaultValue;
                                         return bstrVal[0 .. SysStringLen(bstrVal)].to!T; }
    else static if (is(T : IDispatch)) { if (vt != VT_DISPATCH) return defaultValue;
                                         if (!!pdispVal) pdispVal.AddRef();
                                         return pdispVal; }
    else static if (is(T : IUnknown))  { if (vt != VT_UNKNOWN) return defaultValue;
                                         if (!!punkVal) punkVal.AddRef();
                                         return punkVal; }
    else static if (is(T : Object))      return (vt == VT_BYREF) ? cast(T)byref : defaultValue;
    else static if (is(T == char))       return (vt == VT_I1) ? cVal : defaultValue;
    else static if (is(T == ushort))     return (vt == VT_UI2) ? uiVal : defaultValue;
    else static if (is(T == uint))       return (vt == VT_UI4) ? ulVal : defaultValue;
    else static if (is(T == ulong))      return (vt == VT_UI8) ? ullVal : defaultValue;
    else static if (is(T : DECIMAL))     return (vt == VT_DECIMAL) ? cast(T)decVal : defaultValue;
    else static if (is(T == ubyte*))     return ((vt & ~VT_BYREF) == VT_UI1) ? pbVal : defaultValue;
    else static if (is(T == short*))     return ((vt & ~VT_BYREF) == VT_I2) ? piVal : defaultValue;
    else static if (is(T == int*))       return ((vt & ~VT_BYREF) == VT_I4) ? plVal : defaultValue;
    else static if (is(T == long*))      return ((vt & ~VT_BYREF) == VT_I8) ? pllVal : defaultValue;
    else static if (is(T == float*))     return ((vt & ~VT_BYREF) == VT_R4) ? pfltVal : defaultValue;
    else static if (is(T == double*))    return ((vt & ~VT_BYREF) == VT_R8) ? pdblVal : defaultValue;
    else static if (is(T == char*))      return ((vt & ~VT_BYREF) == VT_I1) ? pcVal : defaultValue;
    else static if (is(T == ushort*))    return ((vt & ~VT_BYREF) == VT_UI2) ? puiVal : defaultValue;
    else static if (is(T == uint*))      return ((vt & ~VT_BYREF) == VT_UI4) ? pulVal : defaultValue;
    else static if (is(T == ulong*))     return ((vt & ~VT_BYREF) == VT_UI8) ? pullVal : defaultValue;
    else static if (is(T == DECIMAL*))   return ((vt & ~VT_BYREF) == VT_DECIMAL) ? pdecVal : defaultValue;
    else static if (is(T == VARIANT*))   return ((vt & ~VT_BYREF) == VT_VARIANT) ? pvarVal : defaultValue;
    else                                 return defaultValue;
  }
}

T variantGet(T)(auto ref VARIANT source, lazy T defaultValue = T.init)
  if (is(T == Variant)) {
  with (source) with (VARENUM) {
    switch (vt) {
    case VT_NULL:      return Variant(null);
    case VT_I8:        return Variant(llVal);
    case VT_I4:        return Variant(iVal);
    case VT_UI1:       return Variant(bVal);
    case VT_I2:        return Variant(iVal);
    case VT_R4:        return Variant(fltVal);
    case VT_R8:        return Variant(dblVal);
    case VT_BOOL:      return Variant(boolVal != VARIANT_FALSE);
    case VT_DATE:    { SYSTEMTIME temp;
                       VariantTimeToSystemTime(date, &temp);
                       return Variant(DateTime(temp.wYear, temp.wMonth, temp.wDay, 
                         temp.wHour, temp.wMinute, temp.wSecond)); }
    case VT_BSTR:      return Variant(bstrVal[0 .. SysStringLen(bstrVal)].toUTF16());
    case VT_BYREF:     return Variant(cast(Object)byref);
    case VT_I1:        return Variant(cVal);
    case VT_UI2:       return Variant(uiVal);
    case VT_UI4:       return Variant(ulVal);
    case VT_UI8:       return Variant(ullVal);
    case VT_INT:       return Variant(intVal);
    case VT_UINT:      return Variant(uintVal);
    case VT_DECIMAL:   return Variant(decVal);
    default:
      switch (vt & ~VT_BYREF) {
      case VT_UI1:     return Variant(pbVal);
      case VT_I2:      return Variant(piVal);
      case VT_I4:      return Variant(plVal);
      case VT_I8:      return Variant(pllVal);
      case VT_R4:      return Variant(pfltVal);
      case VT_R8:      return Variant(pdblVal);
      case VT_UI2:     return Variant(puiVal);
      case VT_UI4:     return Variant(pulVal);
      case VT_UI8:     return Variant(pullVal);
      case VT_DECIMAL: return Variant(pdecVal);
      default:         break;
      }
    }
  }
  return defaultValue;
}

T variantGet(T)(auto ref VARIANT source, lazy T defaultValue = T.init)
  if (isNonCharArray!T) {
  with (source) {
    return (vt == VARENUM.VT_EMPTY || vt == VARENUM.VT_NULL || parray is null)
      ? defaultValue
      : parray.dynamicArray!(ElementEncodingType!T);
  }
}

T variantAs(T)(auto ref VARIANT source, lazy T defaultValue = T.init)
  if (!is(T == Variant)) {
  if (variantIs!T(source))
    return variantGet!T(source, defaultValue);

  VARIANT target;
  if (!VariantChangeTypeEx(&target, &source, LOCALE_USER_DEFAULT, VARIANT_ALPHABOOL, VariantType!T).isSuccessOK)
    return defaultValue;

  scope(exit) variantClear(target);
  return cast(T)variantGet!(OriginalType!T)(target, defaultValue);
}

T variantAs(T)(auto ref VARIANT source, lazy T defaultValue = T.init)
  if (is(T == Variant)) { return variantGet(source, defaultValue); }

pragma(inline, true)
void variantCopy(ref VARIANT source, ref VARIANT target) { checkHR(VariantCopy(&target, &source)); }

pragma(inline, true)
void variantClear(ref VARIANT source) { checkHR(VariantClear(&source)); }

bool variantIs(T)(auto ref VARIANT source) {
       static if (is(T == int))  return source.vt == VARENUM.VT_I4 || source.vt == VARENUM.VT_INT;
  else static if (is(T == uint)) return source.vt == VARENUM.VT_UI4 || source.vt == VARENUM.VT_UINT;
  else                           return source.vt == VariantType!T;
}

unittest {
  auto x = variantInit(12345);
  assert(variantIs!int(x));
  
  x = variantInit(12345L);
  assert(variantIs!long(x));
  
  x = variantInit(123.45);
  assert(variantIs!double(x));

  x = variantInit(true);
  assert(variantIs!bool(x));
  
  x = variantInit("Hello");
  assert(variantIs!string(x));
  variantClear(x);
}

bool variantIs(alias T)(auto ref VARIANT source) {
  return match!(T, 
    null, () => source.vt == VARENUM.VT_NULL,
    () => false);
}

unittest {
  auto x = variantInit(null);
  assert(variantIs!null(x));
  auto y = variantInit(12345);
  assert(!variantIs!null(y));
}

private enum VariantBinaryOpName(string op) = match!(op,
  op == "+" || op == "+=", () => "Add",
  op == "-" || op == "-=", () => "Sub",
  op == "*" || op == "*=", () => "Mul",
  op == "/" || op == "/=", () => "Div",
  op == "%" || op == "%=", () => "Mod",
  op == "^" || op == "^=", () => "Xor",
  op == "&" || op == "&=", () => "And",
  op == "|" || op == "|=", () => "Or",
  op == "^^" || op == "^^=", () => "Pow",
  op == "~" || op == "~=", () => "Cat",
  op == "||", () => "Imp");

enum isVariantBinaryOp(string op) = is(typeof(VariantBinaryOpName!op));

unittest {
  assert(isVariantBinaryOp!"+" && isVariantBinaryOp!"+=");
  assert(isVariantBinaryOp!"-" && isVariantBinaryOp!"-=");
  assert(isVariantBinaryOp!"*" && isVariantBinaryOp!"*=");
  assert(isVariantBinaryOp!"/" && isVariantBinaryOp!"/=");
  assert(isVariantBinaryOp!"%" && isVariantBinaryOp!"%=");
  assert(isVariantBinaryOp!"^" && isVariantBinaryOp!"^=");
  assert(isVariantBinaryOp!"&" && isVariantBinaryOp!"&=");
  assert(isVariantBinaryOp!"|" && isVariantBinaryOp!"|=");
  assert(isVariantBinaryOp!"^^" && isVariantBinaryOp!"^^=");
  assert(isVariantBinaryOp!"~" && isVariantBinaryOp!"~=");
  assert(isVariantBinaryOp!"||");
  assert(!isVariantBinaryOp!"$");
}

private enum VariantUnaryOpName(string op) = match!(op,
  "!", () => "Not",
  "-", () => "Neg",
  "+", () => "Abs");

enum isVariantUnaryOp(string op) = is(typeof(VariantUnaryOpName!op));

unittest {
  assert(isVariantUnaryOp!"!");
  assert(isVariantUnaryOp!"-");
  assert(isVariantUnaryOp!"+");
  assert(!isVariantUnaryOp!"$");
}

VARIANT variantBinaryOp(string op)(auto ref VARIANT left, auto ref VARIANT right)
  if (isVariantBinaryOp!op) {
  static if (op.endsWith("=")) {
    checkHR(mixin("Var" ~ VariantBinaryOpName!op ~ "(&left, &right, &left)"));
    return left;
  } else {
    VARIANT result;
    checkHR(mixin("Var" ~ VariantBinaryOpName!op ~ "(&left, &right, &result)"));
    return result;
  }
}

unittest {
  VARIANT a = { vt: VARENUM.VT_I4, lVal: 20 };
  VARIANT b = { vt: VARENUM.VT_I4, lVal: 2 };
  variantBinaryOp!"+="(a, b);
  assert(a.lVal == 22);
  variantBinaryOp!"*="(a, b);
  assert(a.lVal == 44);
  variantBinaryOp!"/="(a, b);
  assert(a.dblVal == 22);
  variantBinaryOp!"-="(a, b);
  assert(a.dblVal == 20);
  variantBinaryOp!"%="(a, b);
  assert(a.lVal == 0);
}

VARIANT variantUnaryOp(string op)(auto ref VARIANT source)
  if (isVariantUnaryOp!op) {
  VARIANT result;
  checkHR(mixin("Var" ~ VariantUnaryOpName!op ~ "(&source, &result)"));
  return result;
}

unittest {
  VARIANT a = { vt: VARENUM.VT_I4, lVal: 20 };
  a = variantUnaryOp!"-"(a);
  assert(a.lVal == -20);
  a = variantUnaryOp!"+"(a);
  assert(a.lVal == 20);
}

bool variantCmpOp(string op)(auto ref VARIANT left, auto ref VARIANT right) {
  auto result = VarCmp(&left, &right, LOCALE_USER_DEFAULT, 0);
  return match!(op,
    "==", () => result == VARCMP_EQ,
    "!=", () => result != VARCMP_EQ,
    "<", () => result == VARCMP_LT,
    "<=", () => result <= VARCMP_EQ,
    ">", () => result == VARCMP_GT,
    ">=", () => result >= VARCMP_EQ);
}

template VectorType(T) {
       static if (is(ElementEncodingType!T == string)) alias VectorType = BSTR;
  else static if (is(ElementEncodingType!T == bool))   alias VectorType = VARIANT_BOOL;
  else static if (is(ElementEncodingType!T E == enum)) alias VectorType = VectorType!(E[]);
  else                                                 alias VectorType = ElementEncodingType!T;
}

void safeArray(T)(T[] source, out SAFEARRAY* destination) {
  alias E = Unqual!(VectorType!(T[]));
  destination = SafeArrayCreateVector(VariantType!E, 0, cast(uint)source.length);
  E* data;
  checkHR(SafeArrayAccessData(destination, cast(void**)&data));
  foreach (index; 0 .. source.length) {
         static if (is(E == BSTR))       { auto temp = (cast(OriginalType!T)source[index]).toUTF16();
                                           data[index] = SysAllocStringLen(temp.ptr, cast(uint)temp.length); }
    else static if (is(E == VARIANT_BOOL)) data[index] = source[index] ? VARIANT_TRUE : VARIANT_FALSE;
    else                                   data[index] = cast(E)source[index];
  }
  checkHR(SafeArrayUnaccessData(destination));
}

SAFEARRAY* safeArray(T)(T[] source) {
  SAFEARRAY* result;
  source.safeArray(result);
  return result;
}

void dynamicArray(T)(SAFEARRAY* source, out T[] destination) {
  alias E = Unqual!(VectorType!(T[]));
  int ubound, lbound;
  checkHR(SafeArrayGetUBound(source, 1, &ubound));
  checkHR(SafeArrayGetLBound(source, 1, &lbound));
  auto count = ubound - lbound + 1;
  if (count > 0) {
    destination.length = count;
    E* data;
    checkHR(SafeArrayAccessData(source, cast(void**)&data));
    foreach (index; lbound .. ubound + 1) {
           static if (is(E == BSTR))         destination[index] = data[index][0 .. SysStringLen(data[index])].to!T;
      else static if (is(E == VARIANT_BOOL)) destination[index] = data[index] != VARIANT_FALSE;
      else                                   destination[index] = cast(T)data[index];
    }
    checkHR(SafeArrayUnaccessData(source));
  }
}

T[] dynamicArray(T)(SAFEARRAY* source) {
  T[] result;
  source.dynamicArray(result);
  return result;
}

void safeArrayDestroy(SAFEARRAY* source) { SafeArrayDestroy(source); }

private auto eventListener(T)(T handler, int dispId)
  if (isDelegate!T || isFunctionPointer!T) {

  static class EventListener : IDispatch {
    mixin Dispatch;

    T handler_;
    int dispId_;

    this(T handler, int dispId) {
      handler_ = handler;
      dispId_ = dispId;
    }

    override HRESULT Invoke(int dispIdMember, REFIID, uint, ushort, DISPPARAMS* pDispParams,
                            VARIANT* pVarResult, EXCEPINFO* pExcepInfo, uint* puArgErr) {
      if (dispIdMember == dispId_)
        return invokeHandler(dispIdMember, pDispParams, pVarResult, pExcepInfo, puArgErr, handler_);
      return S_OK;
    }

  }
  return new EventListener(handler, dispId).asUnknownPtr();
}

struct _Variant {

  private VARIANT value_;

  // Event handler stuff
  private int eventDispId_;
  private GUID eventSourceIid_;
  private __gshared uint[size_t][size_t] eventCookies_;

  this(T)(auto ref T value)
    if (!isNonCharArray!T && !is(T == Variant)) {
    variantPut(value_, value, VariantType!T);
  }

  this(T)(auto ref T value, int type)
    if (!isNonCharArray!T &&
        !isDelegate!T &&
        !isFunctionPointer!T &&
        !is(T == Variant)) {
    variantPut(value_, value, type);
  }

  this(T)(T value)
    if (isNonCharArray!T || is(T == Variant)) {
    variantPut(value_, value);
  }

  this(ref VARIANT value) { variantCopy(value, value_); }
  this(ref typeof(this) value) { variantCopy(value.value_, value_); }

  private this(IDispatch ptr, int dispId, GUID sourceIid) {
    this(ptr);
    eventDispId_ = dispId;
    eventSourceIid_ = sourceIid;
  }

  ~this() { variantClear(value_); }

  ref VARIANT wrappedValue() @property return { return value_; }
  alias wrappedValue this;

  ref opAssign(T)(auto ref T value)
    if (!is(T == typeof(null)) && !is(T == typeof(this))) {
    variantPut(value_, value);
    return this;
  }

  ref opAssign(T)(T)
    if (is(T == typeof(null))) {
    variantPut(value_, null);
    return this;
  }

  void opOpAssign(string op, T)(T handler)
    if (isDelegate!T || isFunctionPointer!T) {
    static assert(op == "+" || op == "~" || op == "-",
                  "Events can only appear on the left hand side of +=, ~= or -=");

    import std.functional : toDelegate;

    static if (is(typeof(handle.ptr)))
      auto dg = !!handler.ptr ? handler : handle.funcptr.toDelegate();
    else
      auto dg = handler.toDelegate();

    auto key = dg.hashOf();
    auto target = dg.ptr;
    assert(target !is null);

    static if (op == "+" || op == "~") {
      auto container = value_.pdispVal.tryAs!IConnectionPointContainer();
      IConnectionPoint connectionPoint;
      try checkHR(container.FindConnectionPoint(&eventSourceIid_, &connectionPoint));
      finally container.Release();

      uint cookie;
      try checkHR(connectionPoint.Advise(dg.eventListener(eventDispId_), &cookie));
      finally connectionPoint.Release();

      synchronized eventCookies_[cast(size_t)target][key] = cookie;
    } else static if (op == "-") {
      synchronized if (auto cookies = cast(size_t)target in eventCookies_) {
        if (auto cookie = key in *cookies) {
          if (auto container = value_.pdispVal.tryAs!IConnectionPointContainer()) {
            IConnectionPoint connectionPoint;
            try checkHR(container.FindConnectionPoint(&eventSourceIid_, &connectionPoint));
            finally container.Release();

            (*cookies).remove(key);

            try checkHR(connectionPoint.Unadvise(*cookie));
            finally connectionPoint.Release();
          }
        }
      }
    }
  }

  int opCmp(ref typeof(this) other) {
    return VarCmp(&value_, &other.value_, LOCALE_USER_DEFAULT, 0) - VARCMP_EQ;
  }

  bool opEquals(ref typeof(this) other) { return opCmp(other) == 0; }

  bool opEquals(T)(T)
    if (is(T == typeof(null))) {
    with (value_) with (VARENUM) {
      return vt == VT_NULL || vt == VT_EMPTY || 
        (vt == VT_DISPATCH && pdispVal is null) ||
        (vt == VT_UNKNOWN && punkVal is null);
    }
  }

  ref opCast(T)() if (is(T == VARIANT)) { return value_; }
  T opCast(T)() { return variantAs!T(value_); }
  
  version(none) string toString() {
    string str = variantAs!string(value_);

    if (str.empty && value_.vt == VARENUM.VT_DISPATCH) {
      ITypeInfo typeInfo;
      if (value_.pdispVal.GetTypeInfo(0, LOCALE_USER_DEFAULT, &typeInfo).isSuccess 
          && typeInfo !is null) {
        scope(exit) typeInfo.Release();
        BSTR name;
        if (typeInfo.GetDocumentation(MEMBERID_NIL, &name, null, null, null).isSuccess) {
          scope(exit) SysFreeString(name);
          str = name[0 .. SysStringLen(name)].toUTF8();
        }
      }
    }

    return str;
  }

  template opDispatch(string name) {
    template opDispatch(T...) {
      static if (T.length != 0) {
        auto ref opDispatch(Args...)(auto ref Args args) {
          return mixin("value_." ~ name ~ "!T(args)");
        }
      } else static if (__traits(hasMember, value_, name)) {
        auto ref opDispatch() @property { return mixin("value_." ~ name); }
        auto ref opDispatch(Args...)(auto ref Args args) { return mixin("value_." ~ name ~ "(args)"); }
      } else {
        auto ref opDispatch(Args...)(auto ref Args args) {
          enforce(value_.vt == VARENUM.VT_DISPATCH, "no property `" ~ name ~ "` for type `VARIANT`");
          return opDispatchInvokeImpl(value_.pdispVal, name, args);
        }
      }
    }
  }

  int opApply(scope int delegate(ref typeof(this)) it) {
    enforce(value_.vt == VARENUM.VT_DISPATCH, "invalid `foreach` aggregate");
    return opApplyInvokeImpl(value_.pdispVal, it, null);
  }

  int opApply(scope int delegate(size_t, ref typeof(this)) it) {
    enforce(value_.vt == VARENUM.VT_DISPATCH, "invalid `foreach` aggregate");
    return opApplyInvokeImpl(value_.pdispVal, null, it);
  }

  auto ref opIndex(TIndices...)(auto ref TIndices indices) {
    enforce(value_.vt == VARENUM.VT_DISPATCH, "cannot use `[]` operator on type `VARIANT`");
    return opIndexInvokeImpl(value_.pdispVal, indices);
  }

  void opIndexAssign(T, TIndices)(auto ref T value, auto ref TIndices indices) {
    enforce(value_.vt == VARENUM.VT_DISPATCH, "cannot use `[]` operator on type `VARIANT`");
    opIndexAssignInvokeImpl(value_.pdispVal, value, indices);
  }

}

// Enables COM objects to be managed by the GC
extern(C) Object _d_newclass(const(ClassInfo) typeInfo) {
  import core.memory : GC;

  auto attr = GC.BlkAttr.NONE;
  if (typeInfo.m_flags & TypeInfo_Class.ClassFlags.hasDtor &&
      !(typeInfo.m_flags & TypeInfo_Class.ClassFlags.isCPPclass))
    attr |= GC.BlkAttr.FINALIZE;
  if (typeInfo.m_flags & TypeInfo_Class.ClassFlags.noPointers)
    attr |= GC.BlkAttr.NO_SCAN;

  void* p = GC.malloc(typeInfo.initializer.length, attr, typeInfo);
  p[0 .. typeInfo.initializer.length] = typeInfo.initializer[];
  return cast(Object)p;
}

extern (C) void rt_finalize(void* p, bool det = true);

/**
 A template that provides overrides for methods required by the IUnknown interface.
 Examples:
 ---
 class MyUnknown : IUnknown {
   mixin Unknown;
 }
 ---
 */
mixin template Unknown() {

  import core.sys.windows.windef,
    core.sys.windows.basetyps;

  private enum referencesInit_ = 1;
  private shared int references_ = referencesInit_;

  override extern(Windows) HRESULT QueryInterface(REFIID riid, void** ppv) {
    import core.sys.windows.unknwn, std.traits : InterfacesTuple;

    if (ppv is null) return E_POINTER;

    foreach (T; InterfacesTuple!(typeof(this))) {
      static if (is(T : IUnknown) && is(typeof(guidOf!T))) {
        if (*riid == guidOf!T) {
          *ppv = cast(void*)cast(T)this;
          (cast(IUnknown)this).AddRef();
          return S_OK;
        }
      }
    }

    return E_NOINTERFACE;
  }

  override extern(Windows) uint AddRef() {
    import core.atomic : atomicOp;
    import core.memory : GC;
    
    if (references_.atomicOp!"+="(1) == referencesInit_ + 1)
      GC.addRoot(cast(void*)this);
    return references_;
  }

  override extern(Windows) uint Release() {
    import core.atomic : atomicOp;
    //import core.stdc.stdlib : free;
    import core.memory : GC;

    immutable remaining = references_.atomicOp!"-="(1);
    if (remaining == 0) {
      //static if (is(typeof(this.__dtor))) this.__dtor();
      // lifetime.d's _d_newclass uses `malloc` for COM classes so `free` here
      //free(cast(void*)this);
      GC.removeRoot(cast(void*)this);
      rt_finalize(cast(void*)this);
    }
    return remaining;
  }

}

template Mirror(T)
  if (is(T : IUnknown)) {

  alias interfaces          = InterfacesTuple!T;

  alias allMembers          = __traits(allMembers, T);
  alias overloads(string m) = __traits(getOverloads, T, m);
  alias getMember(string m) = __traits(getMember, T, m);

  alias specialMembers = AliasSeq!("opUnary", "opBinary", "opBinaryRight", "opOpAssign", "opAssign", 
                                   "opApply", "opIndex", "opIndexAssign", "opIndexUnary", "opIndexOpAssign",
                                   "opSlice", "opSliceOpAssign", "opDollar", "opCmp", "opEquals", 
                                   "opIn", "opIn_r", "opDispatch", "opCall", "__ctor", "__dtor");

  enum isMemberValid(string m) = is(typeof(getMember!m));
  enum hasMember(string m)     = __traits(hasMember, T, m);
  enum isFunction(string m)    = .isFunction!(getMember!m);
  enum isStatic(string m)      = __traits(isStaticFunction, getMember!m);
  enum isAbstract(string m)    = __traits(isAbstractFunction, getMember!m);
  enum isSpecialMember(string m) = m.among(specialMembers);
  enum isObjectMember(string m)  = __traits(hasMember, Object, m);
  enum isIDispatchMember(string m) = __traits(hasMember, IDispatch, m);
  enum hasLinkage(string m, l...)  = __traits(getLinkage, getMember!m).among(l);

  enum methods = Filter!(templateAnd!(isMemberValid, isFunction, templateNot!isStatic, templateNot!isSpecialMember,
                                      templateNot!isObjectMember, templateNot!isIDispatchMember),
                         allMembers);

}

auto invokeHandler(T)(int dispIdMember, DISPPARAMS* pDispParams, VARIANT* pVarResult,
                      EXCEPINFO* pExcepInfo, uint* puArgErr, T handler)
  if (isDelegate!T || isFunctionPointer!T) {
  try {
    Parameters!T args;

    if (pDispParams) {
      if (pDispParams.cArgs != args.length) return DISP_E_BADPARAMCOUNT;
      if (pDispParams.cNamedArgs != 0) return DISP_E_NONAMEDARGS;

      foreach (i, arg; args) {
        try {
          args[i] = variantAs(pDispParams.rgvarg[pDispParams.cArgs - i - 1],
                              typeof(arg).init);
        } catch {
          if (puArgErr)
            *puArgErr = cast(uint)(pDispParams.cArgs - i - 1);
          return DISP_E_TYPEMISMATCH;
        }
      }
    }

    static if (is(ReturnType!handler == void)) {
      handler(args);
    } else {
      auto result = handler(args);
      if (pVarResult)
        variantPut(*pVarResult, result);
    }

    alias stc = ParameterStorageClassTuple!handler;
    
    static if (stc.length >= 1) {
      assert(stc.length == args.length);
    
      if (pDispParams) {
        foreach (i, arg; args) {
          static if (stc[i] == ParameterStorageClass.out_ ||
                     stc[i] == ParameterStorageClass.ref_)
            variantPut(pDispParams.rgvarg[pDisParams.cArgs - i - 1], &arg);
        }
      }
    }

    return S_OK;
  } catch (Exception ex) {
    return DISP_E_EXCEPTION;
  }
}

/**
 A template that provides overrides of methods required by the IDispatch interface.
 */
mixin template Dispatch() {

  import core.sys.windows.windef,
    core.sys.windows.basetyps;

  mixin Unknown;

  private alias mirror = Mirror!(typeof(this));

  override extern(Windows) HRESULT GetTypeInfoCount(uint*) { return E_NOTIMPL; }
  override extern(Windows) HRESULT GetTypeInfo(uint, uint, ITypeInfo*) { return E_NOTIMPL; }

  override extern(Windows) HRESULT GetIDsOfNames(REFIID riid, wchar** rgszNames, uint cNames, uint, int* rgDispId) {
    import std.traits, std.utf, core.stdc.wchar_ : wcslen;

    if (riid !is null && *riid != guidOf!null) return DISP_E_UNKNOWNINTERFACE;
    if (cNames == 0) return DISP_E_UNKNOWNNAME;

    bool found;
    foreach (i; 0 .. cNames) {
      if (auto szName = rgszNames[i]) {
        auto name = szName[0 .. wcslen(szName)].toUTF8();
        rgDispId[i] = DISPID_UNKNOWN;
        found = false;

        int dispIdBase = DISPID_VALUE;
        foreach (m; mirror.methods) {
          if (m.icmp(name) == 0) {
            alias member = mirror.getMember!m;
            found = true;
            static if (hasUDA!(member, dispId))
              rgDispId[i] = getUDAs!(member, dispId)[0].value;
            else
              rgDispId[i] = i + dispIdBase;
          }
          dispIdBase++;
        }
      }
    }

    return found ? S_OK : DISP_E_UNKNOWNNAME;
  }

  override extern(Windows) HRESULT Invoke(int dispIdMember, REFIID, uint, ushort, DISPPARAMS* pDispParams,
                                          VARIANT* pVarResult, EXCEPINFO* pExcepInfo, uint* puArgErr) {
    import std.traits;

    int dispIdBase = DISPID_VALUE;
    foreach (m; mirror.methods) {
      alias member = mirror.getMember!m;
      int dispIdToTest = dispIdBase;
      static if (hasUDA!(member, dispId))
        dispIdToTest = getUDAs!(member, dispId)[0].value;

      if (dispIdToTest == dispIdMember)
        return invokeHandler(dispIdMember, pDispParams, pVarResult, pExcepInfo, puArgErr, &member);

      dispIdBase++;
    }

    return DISP_E_MEMBERNOTFOUND;
 }

}

class RangeEnumerator(Enumerator) : Enumerator if (isEnumerator!Enumerator) {
  mixin Unknown;

  import core.stdc.wchar_ : wcscpy;

  private alias Element = EnumeratorElementType!Enumerator;
  static if (is(Element == WCHAR*)) private alias Storage = string;
  else                              private alias Storage = Element;

  private InputRange!Storage range_;

  this(R)(R range) { range_ = inputRangeObject(range); }
  ~this() { range_ = null; }

  override HRESULT Next(uint celt, Element* rgelt, uint* pceltFetched) {
    if (pceltFetched !is null) 
      *pceltFetched = 0;
    if (rgelt is null) 
      return E_INVALIDARG;
    if (range_.empty)
      return S_FALSE;

    uint n;
    foreach (_; 0 .. celt) {
      if (range_.empty) break;
      static if (is(Element == Storage)) {
        rgelt[n++] = range_.front;
      } else static if (isString!Storage && is(Element == WCHAR*)) {
        auto temp = range_.front.toUTF16();
        rgelt[n] = cast(WCHAR*)CoTaskMemAlloc((temp.length + 1) * WCHAR.sizeof);
        wcscpy(rgelt[n++], (temp ~ '\0').ptr);
      }
      range_.popFront();
    }

    if (pceltFetched !is null)
      *pceltFetched = n;
    return n == celt ? S_OK : S_FALSE;
  }

  override HRESULT Skip(uint celt) { return range_.popFrontN(celt) == celt ? S_OK : S_FALSE; }
  override HRESULT Reset() { return S_FALSE; }
  override HRESULT Clone(Enumerator* ppenum) { return E_NOTIMPL; }

}

enum Contains(T, U...) = staticIndexOf!(T, U) != -1;

string toCamelCase(string input) {
  import std.uni : toLower;
  return cast(char)input.front.toLower() ~ input.drop(1);
}

template signatureOf(alias F) if (isFunction!F) {
  immutable signatureOf = typeof(&F).stringof.replace("function", __traits(identifier, F));
}

template ParameterIdentifiers(alias F) {
  static if (is(FunctionTypeOf!F T == __parameters)) {
    template Identifier(size_t i) {
      static if (is(typeof(__traits(identifier, T[i .. i + 1]))) &&
                 T[i].stringof != T[i .. i + 1].stringof[1 .. $ - 1])
        enum Identifier = __traits(identifier, T[i .. i + 1]);
      else
        // Relies on DMD giving unnamed parameters a name like so: _param_<index>
        enum Identifier = mixin("_param_$i".inject);
    }
  }

  template Impl(size_t i = 0) {
    static if (i == T.length) alias Impl = AliasSeq!();
    else                      alias Impl = AliasSeq!(Identifier!i, Impl!(i + 1));
  }

  alias ParameterIdentifiers = Impl!();
}

enum canMarshalReturn(T, F) = is(T == F) ||
                              (is(T == string)   && (is(F == BSTR) || (is(HSTRING) && is(F == HSTRING)))) ||
                              (is(T == bool)     && (is(F == BOOL) || is(F == VARIANT_BOOL))) ||
                              (is(T == long)     && is(F == LARGE_INTEGER)) ||
                              (is(T == ulong)    && is(F == ULARGE_INTEGER)) ||
                              (isDynamicArray!T  && is(F == SAFEARRAY*)) ||
                              (isInputRange!T    && isEnumerator!F 
                                                 && (is(ElementEncodingType!T == EnumeratorElementType!F) ||
                                                     isString!(ElementEncodingType!T) && is(EnumeratorElementType!F == WCHAR*))) ||
                              (!is(T == VARIANT) && is(F == VARIANT));

enum defaultMarshalingType(T) = match!(T,
  BOOL, () => MarshalingType.bool_,
  VARIANT_BOOL, () => MarshalingType.variantBool,
  BSTR, () => MarshalingType.bstr,
  () => cast(MarshalingType)0);

enum marshalParam() = q{
  alias _$memberName_out_$i = $(paramNames[i]);
};

enum marshalParam(T : BSTR) = q{
  auto _$memberName_out_$i = $(paramNames[i])[0 .. SysStringLen($(paramNames[i]))].to!$(Source.stringof);
};

enum marshalParam(T : LPCWSTR) = q{
  auto _$memberName_out_$i = $(paramNames[i])[0 .. wcslen($(paramNames[i]))].to!$(Source.stringof);
};

static if (is(HSTRING))
enum marshalParam(T : HSTRING) = q{
  uint _$memberName_length_$i;
  auto _$memberName_ptr_$i = WindowsGetStringRawBuffer($(paramNames[i]), &_$memberName_length_$i);
  auto _$memberName_out_$i = !!_$memberName_ptr_$i && !!_$memberName_length_$i
    ? _$memberName_ptr_$i[0 .. _$memberName_length_$i].to!$(Source.stringof)
    : null;
};

enum marshalParam(T : BOOL) = q{
  auto _$memberName_out_$i = $(paramNames[i]) != FALSE;
};

enum marshalParam(T : VARIANT_BOOL) = q{
  auto _$memberName_out_$i = $(paramNames[i]) != VARIANT_FALSE;
};

enum marshalParam(T : VARIANT) = q{
  auto _$memberName_out_$i = variantGet($(paramNames[i]), $(Source.stringof));
};

template marshalParam(T)
  if (is(T == LARGE_INTEGER) || is(T == ULARGE_INTEGER)) {
  enum marshalParam = q{
    auto _$memberName_out_$i = $(paramNames[i]).QuadPart;
  };
}

// We have to alias the parameter name in case it was auto-generated (ie, unnamed) - it would shadow our 
// lambda's arguments.
enum marshalParam(T : IDispatch) = q{
  alias _$memberName_param_$i = $(paramNames[i]);
  auto _$memberName_out_$i = (Parameters!($(Source.stringof)) _$memberName_args_$i) => cast(ReturnType!($(Source.stringof)))
    _$memberName_param_$i.invokeById(DISPID_VALUE, cast(ushort)DISPATCH_METHOD, _$memberName_args_$i);
};

enum marshalParam(T : SAFEARRAY*, U : U[]) = q{
  auto _$memberName_out_$i = $(paramNames[i]).dynamicArray!$(ElementEncodingType!Source.stringof);
};

enum marshalParam(T : WCHAR**, ParameterStorageClass _, MarshalingType __) = q{
  auto _$memberName_out_$i = (*$(paramNames[i]) !is null)
    ? (*$(paramNames[i]))[0 .. SysStringLen(*$(paramNames[i]))].to!$(Source.stringof)
    : null;
  scope(exit) {
    auto _$memberName_temp_$i = _$memberName_out_$i.to!wstring;
    *$(paramNames[i]) = SysAllocStringLen(_$memberName_temp_$i.ptr, cast(uint)_$memberName_temp_$i.length);
  }
};

enum marshalParam(T : WCHAR**, ParameterStorageClass _, MarshalingType __ : MarshalingType.lpwstr) = q{
  auto _$memberName_out_$i = (*$(paramNames[i]) !is null)
    ? (*$(paramNames[i]))[0 .. wcslen(*$(paramNames[i]))].to!$(Source.stringof)
    : null;
  scope(exit) {
    auto _$memberName_temp_$i = _$memberName_out_$i.to!wstring;
    *$(paramNames[i]) = cast(wchar*)(_$memberName_temp_$i ~ '\0').ptr;
  }
};

static if (is(HSTRING))
enum marshalParam(T : HSTRING*, ParameterStorageClass _, MarshalingType __) = q{
  uint _$memberName_length_$i;
  auto _$memberName_ptr_$i = WindowsGetStringRawBuffer(*$(paramNames[i]), &_$memberName_length_$i);
  auto _$memberName_out_$i = !!_$memberName_ptr_$i && !!_$memberName_length_$i
    ? _$memberName_ptr_$i[0 .. _$memberName_length_$i].to!$(Source.stringof)
    : null;
  scope(exit) {
    auto _$memberName_temp_$i = _$memberName_out_$i.to!wstring;
    if (*$(paramNames[i]) !is null) WindowsDeleteString(*$(paramNames[i]));
    WindowsCreateString((_$memberName_temp_$i ~ '\0').ptr, cast(uint)_$memberName_temp_$i.length, $(paramNames[i]));
  }
};

enum marshalParam(T : BOOL*, ParameterStorageClass _, MarshalingType __) = q{
  bool _$memberName_out_$i = (*$(paramNames[i]) != FALSE);
  scope(exit) *$(paramNames[i]) = _$memberName_out_$i ? TRUE : FALSE;
};

enum marshalParam(T : VARIANT_BOOL*, ParameterStorageClass _, MarshalingType __) = q{
  bool _$memberName_out_$i = (*$(paramNames[i]) != VARIANT_FALSE);
  scope(exit) *$(paramNames[i]) = _$memberName_out_$i ? VARIANT_TRUE : VARIANT_FALSE;
};

enum marshalParam(T : VARIANT*, ParameterStorageClass _, MarshalingType __) = q{
  auto _$memberName_out_$i = variantGet(*$(paramNames[i]), $(Source.stringof).init);
  scope(exit) variantPut(*$(paramNames[i]), _$memberName_out_$i);
};

template marshalParam(T, ParameterStorageClass _, MarshalingType __)
  if (is(T == LARGE_INTEGER*) || is(T == ULARGE_INTEGER*)) {
  enum marshalParam = q{
    $(Source.stringof) _$memberName_out_$i = (*$(paramNames[i])).QuadPart;
    scope(exit) (*$(paramNames[i])).QuadPart = _$memberName_out_$i;
  };
}

enum marshalParam(T : SAFEARRAY**, ParameterStorageClass _) = q{
  auto _$memberName_out_$i = (*$(paramNames[i])).dynamicArray!$(ElementEncodingType!Source.stringof);
  scope(exit) {
    SafeArrayDestroy(*$(paramNames[i]));
    *$(paramNames[i]) = _$memberNames_out_$i.safeArray;
  }
};

template marshalParam(T : void*, ParameterStorageClass _)
  if (!(is(T == LARGE_INTEGER*) || is(T == ULARGE_INTEGER*))) {
  enum marshalParam = q{
    auto _$memberName_out_$i = *$(paramNames[i]);
    scope(exit) *$(paramNames[i]) = _$memberName_out_$i;
  };
}

enum marshalReturn() = q{
  *$(paramNames[$ - 1]) = _$memberName_result;
};

enum marshalReturn(T : WCHAR*, MarshalingType _) = q{
  auto _$memberName_result_temp = _$memberName_result.to!wstring;
  *$(paramNames[$ - 1]) = SysAllocStringLen(_$memberName_result_temp.ptr, cast(uint)_$memberName_result_temp.length);
};

enum marhsalReturn(T : WCHAR*, MarshalingType _ : MarshalingType.lpstr) = q{
  *$(paramNames[$ - 1]) = _$memberName_result.toUTFz!(char*);
};

enum marshalReturn(T : WCHAR*, MarshalingType _ : MarshalingType.lpwstr) = q{
  *$(paramNames[$ - 1]) = _$memberName_result.toUTFz!(wchar*);
};

static if (is(HSTRING))
enum marshalReturn(T : HSTRING, MarshalingType _) = q{
  auto _$memberName_result_temp = _$memberName_result.to!wstring;
  checkHR(WindowsCreateString((_$memberName_result_temp ~ '\0').ptr, cast(uint)_$memberName_result_temp.length,
                               $(paramNames[$ - 1])));
};

template marshalReturn(T, MarshalingType _) if (isEnumerator!T) {
  enum marshalReturn = q{
    *$(paramNames[$ - 1]) = new RangeEnumerator!$(TReturnParam.stringof)(_$memberName_result);
  };
}

enum marshalReturn(T : BOOL, MarshalingType _ : MarshalingType.bool_) = q{
  *$(paramNames[$ - 1]) = _$memberName_result ? TRUE : FALSE;
};

enum marshalReturn(T : VARIANT_BOOL, MarshalingType _ : MarshalingType.variantBool) = q{
  *$(paramNames[$ - 1]) = _$memberName_result ? VARIANT_TRUE : VARIANT_FALSE;
};

template marshalReturn(T, MarshalingType _)
  if (is(T == LARGE_INTEGER) || is(T == ULARGE_INTEGER)) {
  enum marshalReturn = q{
    (*$(paramNames[$ - 1])).QuadPart = _$memberName_result;
  };
}

enum marshalReturn(T : VARIANT, MarshalingType _) = q{
  variantPut(*$(paramNames[$ - 1]), _$memberName_result);
};

enum marshalReturn(T : SAFEARRAY*, MarshalingType _) = q{
  *$(paramNames[$ - 1]) = _$memberName_result.safeArray;
};

/**
 A template that generates overrides for any interface in the type's base interface list.
 */
mixin template Marshaling() {

  import core.sys.windows.w32api,
    core.sys.windows.winbase,
    core.sys.windows.windef, 
    core.sys.windows.wtypes, 
    core.sys.windows.basetyps,
    core.sys.windows.oleauto,
    std.traits;
  import core.stdc.wchar_ : wcslen;
  import std.conv : to;
  static if (_WIN32_WINNT >= 0x602) import core.sys.windows.winrt.hstring,
    core.sys.windows.winrt.winstring;

  private static string generateMarshaling() {
    import std.traits, std.meta, std.algorithm, std.range, 
      comet.attributes, comet.internal;
    static if (_WIN32_WINNT >= 0x602) import core.sys.windows.winrt.inspectable;

    static if (is(IInspectable)) alias excludeTypes = AliasSeq!(IUnknown, IDispatch, IInspectable);
    else                         alias excludeTypes = AliasSeq!(IUnknown, IDispatch);
    alias self                = Mirror!(typeof(this));
    enum isMarshaled(alias m) = !hasUDA!(m, notMarshaled);

    string code;
    foreach (T; Filter!(ApplyRight!(templateNot!Contains, excludeTypes), self.interfaces)) {
      alias base = Mirror!T;

      foreach (baseMemberName; Filter!(templateAnd!(base.isAbstract,
                                                    ApplyRight!(base.hasLinkage, "Windows", "System", "C++")),
                                       base.methods)) {
        static if (base.hasMember!baseMemberName && (self.hasMember!baseMemberName ||
                                                     self.hasMember!(baseMemberName.toCamelCase()))) {
          foreach (baseMember; base.overloads!baseMemberName) {
            enum altMemberName = baseMemberName.toCamelCase();
            enum memberName = Select!(self.hasMember!altMemberName, altMemberName, baseMemberName);

            foreach (member; Filter!(isMarshaled, self.overloads!memberName)) {
              alias TReturn = ReturnType!member, TBaseReturn = ReturnType!baseMember;
              alias TParams = Parameters!member, TBaseParams = Parameters!baseMember;

              enum paramStorage = ParameterStorageClassTuple!member,
                   paramNames = ParameterIdentifiers!baseMember,
                   paramCount = paramNames.length;

              static if (TBaseParams.length == TParams.length + 1 &&
                         isPointer!(TBaseParams[$ - 1]) &&
                         canMarshalReturn!(TReturn, PointerTarget!(TBaseParams[$ - 1]))) enum hasReturnParam = true;
              else                                                                       enum hasReturnParam = false;

              static if ((TBaseParams.length == TParams.length || hasReturnParam) &&
                         (!is(TReturn == TBaseReturn) ||
                          !is(TParams == TBaseParams) ||
                          self.hasMember!altMemberName)) enum needsMarshaling = true;
              else                                       enum needsMarshaling = false;

              static if (needsMarshaling) {
                static if (hasReturnParam) alias TReturnParam = PointerTarget!(TBaseParams[$ - 1]);

                code ~= "override " ~ signatureOf!baseMember ~ " {
                  ";

                static if (!hasUDA!(member, keepReturn)) {
                  static if (!is(TBaseReturn == void)) code ~= "return ";
                  code ~= "catchHR({
                    ";
                }

                static foreach (i; 0 .. TParams.length) {{
                  alias Source = Unqual!(TParams[i]), Target = TBaseParams[i];

                  // getUDAs doesn't work on parameters
                  static if (__traits(compiles, __traits(getAttributes, TParams[i .. i + 1]))) {
                    enum attributes = __traits(getAttributes, TParams[i .. i + 1]);
                    static if (!is(typeof(marshalingType)) && 
                               attributes.length != 0 &&
                               is(typeof(attributes[0]) == marshalAs))
                      enum marshalingType = attributes[0].value;
                  }
                  static if (!is(typeof(marshalingType)))
                    enum marshalingType = defaultMarshalingType!Source;

                  static if (is(Source) &&

                             // string types
                             ((isString!Source && (is(Target == BSTR) || is(Target == LPCWSTR) ||
                                                   (is(HSTRING) && is(Target == HSTRING)))) ||

                              // bool types
                              (is(Source == bool) && (is(Target == VARIANT_BOOL) || is(Target == BOOL))) ||

                              (!is(Source : VARIANT) && is(Target == VARIANT)) ||
                              
                              // u/long -> U/LARGE_INTEGER
                              (is(Source == long) && is(Target == LARGE_INTEGER)) ||
                              (is(Source == ulong) && is(Target == ULARGE_INTEGER)) ||
                              
                              // delegates -> IDispatch callbacks (extern(D) only)
                              ((isDelegate!Source || isFunctionPointer!Source) && 
                               is(Target : IDispatch) && self.hasLinkage!(memberName, "D"))) ) {
                       code ~= mixin(marshalParam!Target.inject);
                  } else static if (is(Source) &&

                                    ((isDynamicArray!Source && is(Target == SAFEARRAY*))/+ ||
                                     (paramStorage[i] == ParameterStorageClass.none &&
                                      isDynamicArray!Source && !isString!Source && isPointer!Target)+/) ) {
                       code ~= mixin(marshalParam!(Target, Source).inject);
                  } else static if (is(Source) &&

                                    // ref/out
                                    (((paramStorage[i] & ParameterStorageClass.ref_) ||
                                      (paramStorage[i] & ParameterStorageClass.out_)) && isPointer!Target) ) {
                    static if (// string types
                               (isString!Source && (is(Target == BSTR*) || (is(HSTRING) && is(Target == HSTRING*)))) ||

                               // bool types
                               (is(Source == bool) && (is(Target == BOOL*) || is(Target == VARIANT_BOOL*))) ||
                               
                               // any -> VARIANT
                               (!is(Source : VARIANT) && is(Target == VARIANT*)) ||
                               
                               // u/long -> U/LARGE_INTEGER
                               (is(Source == long) && is(Target == LARGE_INTEGER*)) ||
                               (is(Source == ulong) && is(Target == ULARGE_INTEGER*)))
                       code ~= mixin(marshalParam!(Target, paramStorage[i .. i + 1], marshalingType).inject);
                    else static if (isDynamicArray!Source && is(Target == SAFEARRAY**))
                       code ~= mixin(marshalParam!(Target, Source, paramStorage[i .. i + 1]).inject);
                    else
                       code ~= mixin(marshalParam!(Target, paramStorage[i .. i + 1]).inject);
                  } else static if ((isDelegate!Source || isFunctionPointer!Source) &&
                                    !self.hasLinkage!(memberName, "D")) {
                    static assert(false, mixin(("`$(fullyQualifiedName!(typeof(this))).$(memberName, 
                                                formatParameters!member)` must be `extern (D)`").inject));
                  }
                  else code ~= mixin(marshalParam!().inject);
                }}

                static if (!hasReturnParam) immutable count = paramCount;
                else                        immutable count = paramCount - 1;

                static if (hasReturnParam || !is(TReturn == void)) code ~= "auto _" ~ memberName ~ "_result = ";

                code ~= "this." ~ memberName ~ "(";
                code ~= iota(count).map!(i => mixin("_$memberName_out_$i".inject)).join(", ");
                code ~= ");
                  ";

                static if (hasReturnParam) {
                  static if (hasUDA!(member, marshalReturnAs))
                    enum returnMarshalingType = getUDAs!(member, marshalReturnAs)[0].value;
                  else
                    enum returnMarshalingType = defaultMarshalingType!TReturnParam;

                  static if (canMarshalReturn!(TReturn, TReturnParam))
                    code ~= mixin(marshalReturn!(TReturnParam, returnMarshalingType).inject);
                  else
                    code ~= mixin(marshalReturn!().inject);
                } else static if (!is(TReturn == void)) 
                  code ~= "return _" ~ memberName ~ "_result;
                  ";

                static if (!hasUDA!(member, keepReturn)) code ~= "});
                  ";
                code ~= "}
                  ";
              } else {
                code ~= "override " ~ signatureOf!baseMember ~ " { return S_FALSE; }
                  ";
              }
            }
          }
        }
      }
    }
    return code;
  }

  enum code = generateMarshaling;
  //pragma(msg, code);
  mixin(code);
}

struct EventOpHelper(TEventHandler, TToken) {

  alias TArgs = Parameters!(TEventHandler.Invoke);
  alias THandler = void delegate(TArgs);

  static class Callback : TEventHandler, IAgileObject {
    mixin Unknown;

    THandler handler_;
    this(THandler handler) { handler_ = handler; }
    HRESULT Invoke(TArgs args) { return catchHR(() => handler_(args)); }

  }

  import std.functional : toDelegate;

  extern(Windows) alias TAddMethod    = HRESULT delegate(TEventHandler, TToken*);
  extern(Windows) alias TRemoveMethod = HRESULT delegate(TToken);

  private TAddMethod addMethod_;
  private TRemoveMethod removeMethod_;
  private __gshared TToken[size_t][size_t] eventRegistrationTokens_;

  this(TAddMethod addMethod, TRemoveMethod removeMethod) {
    addMethod_ = addMethod;
    removeMethod_ = removeMethod;
  }

  void opOpAssign(string op)(THandler handler) {
    static assert(op == "+" || op == "~" || op == "-",
                  "Events can only appear on the left size of +=, ~= or -=");

    static if (is(typeof(handler.ptr)))
      auto dg = !!handler.ptr ? handler : handler.funcptr.toDelegate();
    else
      auto dg = handler.toDelegate();

    auto key = dg.hashOf();
    auto target = dg.ptr;
    enforce(target !is null);

    static if (op == "+" || op == "~") {
      TToken token;
      checkHR(addMethod_((cast(TEventHandler)new Callback(handler)).asUnknownPtr(), &token));
      synchronized eventRegistrationTokens_[cast(size_t)target][key] = token;
    } else static if (op == "-") {
      synchronized if (auto tokens = cast(size_t)target in eventRegistrationTokens_) {
        if (auto token = key in *tokens) {
          (*token).remove(key);
          checkHR(removeMethod_(*token));
        }
      }
    }
  }

}

static if (_WIN32_WINNT >= 0x602):

import core.sys.windows.winrt.inspectable,
  core.sys.windows.winrt.activation,
  core.sys.windows.winrt.eventtoken;

/**
 A template that provides overrides for methods required by the IInspectable interface.
 */
mixin template Inspectable() {

  mixin Unknown;

  import core.sys.windows.windef,
    core.sys.windows.wtypes,
    core.sys.windows.basetyps,
    core.sys.windows.winrt.inspectable;

  override extern(Windows) HRESULT GetIids(uint* iidCount, IID** iids) {
    import core.sys.windows.unknwn,
      core.sys.windows.objbase : CoTaskMemAlloc;
    import std.meta : Filter, templateAnd, templateNot, ApplyRight;
    import std.traits : InterfacesTuple;

    enum hasGUID(T) = is(typeof(guidOf!T));

    alias interfaces = Filter!(templateAnd!(hasGUID, 
                                            ApplyRight!(templateNot!Contains, IUnknown, IInspectable)),
                               InterfacesTuple!(typeof(this)));

    *iids = null;
    *iidCount = 0;

    static if (interfaces.length > 0) {
      auto piids = cast(GUID*)CoTaskMemAlloc(GUID.sizeof * interfaces.length);
      foreach (i, T; interfaces)
        piids[i] = guidOf!T;

      *iidCount = cast(uint)interfaces.length;
      *iids = piids;
      return S_OK;
    } else {
      return S_FALSE;
    }
  }

  override extern(Windows) HRESULT GetRuntimeClassName(HSTRING* className) {
    import std.utf : toUTF16;
    import core.sys.windows.winrt.winstring : WindowsCreateString;

    auto temp = typeid(this).name.toUTF16();
    return WindowsCreateString((temp ~ '\0').ptr, cast(uint)temp.length, className);
  }

  override extern(Windows) HRESULT GetTrustLevel(TrustLevel* trustLevel) {
    *trustLevel = TrustLevel.BaseTrust;
    return S_OK;
  }

}

template flattenedNameOf(T...) {
  alias U = T[0];
  static if (isBasicType!U && !is(U == enum)) enum flattenedNameOf = U.stringof;
  else static if (U.stringof == "HSTRING__*") enum flattenedNameOf = "HSTRING";
  else                                        enum flattenedNameOf = flattenedSym!(true, U);
}

private template flattenedSym(bool start, alias T : U!A, alias U, A...) {

  template fqnTuple(T...) {
    import std.string : startsWith;

    static if (T.length == 0) {
      enum fqnTuple = "";
    } else static if (T.length == 1) {
      static if (isExpressions!T)
        enum fqnTuple = T[0].stringof;
      else static if (flattenedNameOf!(T[0]).startsWith("windows_"))
        enum fqnTuple = flattenedNameOf!(T[0]);
      else static if (T[0].stringof == "HSTRING__*")
        enum fqnTuple = "HSTRING";
      else
        enum fqnTuple = T[0].stringof;
    }
    else {
      enum fqnTuple = fqnTuple!(T[0]) ~ "_" ~ fqnTuple!(T[1 .. $]);
    }
  }

  static if (start)
    enum flattenedSym = __traits(identifier, U) ~ "_" ~ fqnTuple!A;
  else
    enum flattenedSym = flattenedSym!(false, __traits(parent, U)) ~ "_" ~ 
      __traits(identifier, U) ~ "_" ~ fqnTuple!A;
}

private template flattenedSym(bool start, alias T) {
  static if (__traits(compiles, __traits(parent, T)) && !__traits(isSame, T, __traits(parent, T)))
    enum parentPrefix = flattenedSym!(start, __traits(parent, T)) ~ "_";
  else
    enum parentPrefix = null;

  static string adjust(string s) {
    return (s.skipOver("package ") || s.skipOver("module ")) ? s : s.findSplit("(")[0];
  }

  enum flattenedSym = parentPrefix ~ adjust(__traits(identifier, T));
}

enum RuntimeClassName(T, string className) = "immutable " ~ flattenedNameOf!T ~ " = \"" ~ className ~ "\";";

auto factory(T = IActivationFactory)(string activatableClassId = null)
  if (is(T : IInspectable)) {
  if (activatableClassId == null) {
    enum moduleName = .moduleName!T;
    mixin("static import " ~ moduleName ~ ";");
    static if (is(typeof(mixin(moduleName ~ "." ~ flattenedNameOf!T))))
      activatableClassId = mixin(moduleName ~ "." ~ flattenedNameOf!T);
  }

  enforce(activatableClassId != null);

  HSTRING pClassId;
  HSTRING_HEADER rClassId;

  auto temp = activatableClassId.toUTF16();
  WindowsCreateStringReference((temp ~ '\0').ptr, cast(uint)temp.length, &rClassId, &pClassId);

  IInspectable factory;
  immutable iid = guidOf!T;
  if (RoGetActivationFactory(pClassId, &iid, cast(void**)&factory).isSuccessOK)
    return factory.asUnknownPtr!T();
  return UnknownPtr!T.init;
}

auto activate(T = IInspectable)(string activatableClassId = null)
  if (is(T : IInspectable)) {
  if (activatableClassId == null) {
    enum moduleName = .moduleName!T;
    mixin("static import " ~ moduleName ~ ";");
    static if (is(typeof(mixin(moduleName ~ "." ~ flattenedNameOf!T))))
      activatableClassId = mixin(moduleName ~ "." ~ flattenedNameOf!T);
  }

  enforce(activatableClassId != null);

  if (auto factory = factory(activatableClassId)) {
    IInspectable result;
    if (factory.ActivateInstance(&result).isSuccessOK)
      return result.asUnknownPtr!T();
  }
  return UnknownPtr!T.init;
}

auto activateWith(T, Args...)(auto ref Args args) {
  enum defaultOmittedArgs = "/+IInspectable baseInterface+/ null, /+IInspectable* innerInterface+/ null";
  enum omittedArgs = (Args.length == 0) ? defaultOmittedArgs : "args, " ~ defaultOmittedArgs;

  static foreach (m; __traits(derivedMembers, T)) {
    static if (m.startsWith("Create")) {
      static if (__traits(compiles, (UnknownPtr!T r) => mixin(mixin("r.$m(args)".inject)))) {
        static if (!is(typeof(comet_makeWith_factoryFound))) {
          enum comet_makeWith_factoryFound = true;

          auto factory = factory!T();
          return mixin(mixin("factory.$m(args)".inject));
        }
      } else static if (__traits(compiles, (UnknownPtr!T r) => mixin(mixin("r.$m($omittedArgs)".inject)))) {
        static if (!is(typeof(comet_makeWith_factoryFound))) {
          enum comet_makeWith_factoryFound = true;

          auto factory = factory!T();
          return mixin(mixin("factory.$m($omittedArgs)".inject));
        }
      }
    }
  }

  static if (!is(typeof(comet_makeWith_factoryFound))) {
    static if (Args.length == 0)
      return activate!T();
    else
      static assert(false, "`%s` has no factory method that can be called using argument types `%s`"
                    .format(nameOf!T, Args.stringof));
  }
}