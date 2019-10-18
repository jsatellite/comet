module comet.internal;

import std.traits,
  std.meta,
  std.range,
  std.string,
  std.array;
import comet.utils;

package bool isCamelCase(string input) {
  import std.uni : isLower;
  return input.front.isLower();
}

package string toPascalCase(string input) {
  import std.uni : toUpper;
  return cast(char)input.front.toUpper() ~ input.drop(1);
}

private enum enumeratorNames = AliasSeq!("_newEnum", "get__newEnum", "_NewEnum", "get__NewEnum"),
             getIndexerNames = AliasSeq!("item", "get_item", "Item", "get_Item"),
             setIndexerNames = AliasSeq!("put_item", "put_Item");

private template firstOrNull(T, names...) {
  enum hasMember(string name) = __traits(hasMember, T, name);

  enum matches = Filter!(hasMember, names);
  static if (matches.length != 0) enum firstOrNull = matches[0];
  else                            enum firstOrNull = null;
}

package template isEnumerator(T) {
  static if (__traits(hasMember, T, "Next")) {
    alias L = functionLinkage!(T.Next);
    alias P = Parameters!(T.Next);
    enum isEnumerator = P.length == 3 &&
      is(P[0] == uint) && is(P[1] : void*) && is(P[2] == uint*);
  } else {
    enum isEnumerator = false;
  }
}

package template EnumeratorElementType(T) {
  static if (isEnumerator!T) 
    alias EnumeratorElementType = PointerTarget!(Parameters!(T.Next)[1]);
  else
    alias EnumeratorElementType = void;
}

package template hasEnumerator(T) {
  static if (enumeratorName!T != null) {
    enum hasEnumerator = isEnumerator!(PointerTarget!(Parameters!(mixin("T." ~ enumeratorName!T))[0]));
  } else {
    enum hasEnumerator = false;
  }
}

package enum enumeratorName(T) = firstOrNull!(T, enumeratorNames);

package template hasGetIndexer(T) {

  template isGetIndexer(string name) {
    static if (__traits(hasMember, T, name)) {
      alias P = Parameters!(mixin("T." ~ name));
      enum isGetIndexer = P.length == 2 &&
        (is(P[0] == int) || is(P[0] == uint)) && is(P[1] : void*);
    } else {
      enum isGetIndexer = false;
    }
  }

  enum hasGetIndexer = anySatisfy!(isGetIndexer, getIndexerNames);
}

package enum getIndexerName(T) = firstOrNull!(T, getIndexerNames);

package template hasSetIndexer(T) {

  template isSetIndexer(string name) {
    static if (__traits(hasMember, T, name)) {
      alias P = Parameters!(mixin("T." ~ name));
      enum isSetIndexer = P.length == 2 &&
        (is(P[0] == int) || is(P[0] == uint));
    } else {
      enum isSetIndexer = false;
    }
  }

  enum hasSetIndexer = anySatisfy!(isSetIndexer, setIndexerNames);
}

package enum setIndexerName(T) = firstOrNull!(T, setIndexerNames);

package template formatParameters(alias member) {

  string formatParametersImpl() {
    import core.internal.traits : staticIota;

    alias types = Parameters!member;
    alias names = ParameterIdentifierTuple!member;
    alias defaults = ParameterDefaults!member;

    string params(size_t i)() {
      string s = types[i].stringof;
      if (!names[i].empty) s ~= " " ~ names[i];
      if (!is(defaults[i] == void)) s ~= " = " ~ defaults[i].stringof;
      return s;
    }

    static if (types.length == 0)
      return "()";
    else
      return "(" ~ only(staticMap!(params, staticIota!(0, types.length))).join(", ") ~ ")";
  }

  enum formatParameters = formatParametersImpl();
}

package string opDispatchError(string name, T, Args...)() {
  static if (__traits(hasMember, T, name)) {
    alias overloads = __traits(getOverloads, T, name);
    static if (overloads.length == 1) {
      return mixin("function `$(fullyQualifiedName!T).$(name, formatParameters!overloads)` is not callable using argument types `$(Args.stringof)`".inject);
    } else {
      enum candidate(alias member) = mixin("\t\t`$(fullyQualifiedName!T).$(name, formatParameters!member)`".inject);

      return mixin("none of the overloads of `$name` are callable using argument types `$(Args.stringof)`, candidates are:\n".inject) ~
        only(staticMap!(candidate, overloads)).join("\n");
    }
  } else {
    return mixin("no property `$name` for type `$(fullyQualifiedName!T)`".inject);
  }
}

import core.sys.windows.w32api,
  core.sys.windows.windef,
  core.sys.windows.wtypes,
  core.sys.windows.basetyps,
  core.sys.windows.unknwn,
  core.sys.windows.oaidl;
static if (_WIN32_WINNT >= 0x602) import core.sys.windows.winrt.hstring;

package enum opDispatchMarshalParam() = q{
  alias param$i = args[$i];
};

package enum opDispatchMarshalParam(T : BSTR) = q{
  auto temp$i = args[$i].toUTF16();
  auto param$i = SysAllocStringLen(temp$i.ptr, cast(uint)temp$i.length);
  scope(exit) SysFreeString(param$i);
};

package enum opDispatchMarshalParam(T : BSTR*) = q{
  auto temp$i = (*args[$i]).toUTF16();
  auto param$i_ = SysAllocStringLen(temp$i.ptr, cast(uint)temp$i.length);
  scope(exit) SysFreeString(*param$i_);
  auto param$i = &param$i_;
  scope(exit) if (!!*param$i) *args[$i] = (*param$i)[0 .. SysStringLen(*param$i)].to!(typeof(*args[$i]));
};

package enum opDispatchMarshalParam(T : LPCWSTR) = q{
  auto param$i = args[$i].toUTFz!(wchar*);
};

package enum opDispatchMarshalParam(T : LPCWSTR*) = q{
  auto temp$i = (*args[$i]).toUTFz!(wchar*);
  auto param$i = cast(LPCWSTR*)&temp$i;
  scope(exit) if (!!*param$i) *args[$i] = (*param$i)[0 .. wcslen(*param$i)].to!(typeof(*args[$i]));
};

static if (is(HSTRING))
package enum opDispatchMarshalParam(T : HSTRING) = q{
  auto temp$i = args[$i].toUTF16();
  HSTRING param$i;
  HSTRING_HEADER param$(i)Header;
  checkHR(WindowsCreateStringReference((temp$i ~ '\0').ptr, cast(uint)temp$i.length, &param$(i)Header, &param$i));
};

static if (is(HSTRING))
package enum opDispatchMarshalParam(T : HSTRING*) = q{
  auto temp$i = (*args[$i]).toUTF16();
  HSTRING param$i_;
  HSTRING_HEADER param$(i)Header;
  if (!(*args[$i]).empty)
    checkHR(WindowsCreateStringReference((temp$i ~ '\0').ptr, cast(uint)temp$i.length, &param$(i)Header, &param$i_));
  auto param$i = &param$i_;
  scope(exit) {
    WindowsDeleteString(*param$i);
    uint length$i;
    if (auto ptr$i = WindowsGetStringRawBuffer(*param$i, &length$i))
      *args[$i] = ptr$i[0 .. length$i].to!(typeof(*args[$i]));
  }
};

package enum opDispatchMarshalParam(T : BOOL) = q{
  auto param$i = args[$i] ? TRUE : FALSE;
};

package enum opDispatchMarshalParam(T : BOOL*) = q{
  auto temp$i = *args[$i] ? TRUE : FALSE;
  auto param$i = &temp$i;
  scope(exit) *args[$i] = *param$i != FALSE;
};

package enum opDispatchMarshalParam(T : VARIANT_BOOL) = q{
  auto param$i = args[$i] ? VARIANT_TRUE : VARIANT_FALSE;
};

package enum opDispatchMarshalParam(T : VARIANT_BOOL*) = q{
  auto temp$i = *args[$i] ? VARIANT_TRUE : VARIANT_FALSE;
  auto param$i = &temp$i;
  scope(exit) *args[$i] = *param$i != VARIANT_FALSE;
};

package enum opDispatchMarshalParam(T : VARIANT) = q{
  auto param$i = variantInit(args[$i]);
  scope(exit) variantClear(param$i);
};

package enum opDispatchMarshalParam(T : VARIANT*) = q{
  auto temp$i = variantInit(*args[$i]);
  auto param$i = &temp$i;
  scope(exit) variantClear(temp$i);
  scope(exit) *args[$i] = variantGet(*param$i, typeof(*args[$i]).init);
};

package enum opDispatchMarshalParam(T : SAFEARRAY*) = q{
  auto param$i = args[$i].safeArray;
};

package enum opDispatchMarshalParam(T : SAFEARRAY**) = q{
  auto temp$i = (*args[$i]).safeArray;
  auto param$i = &temp$i;
  scope(exit) if (temp$i !is null) {
    temp$i.dynamicArray(*args[$i]);
    SafeArrayDestroy(temp$i);
  }
};

package template opDispatchMarshalParam(T : T*)
  if (!(is(T == LARGE_INTEGER) || is(T == ULARGE_INTEGER))) {
  enum opDispatchMarshalParam = q{
    auto param$i = args[$i].ptr;
  };
}

package template opDispatchMarshalParam(T)
  if (is(T == LARGE_INTEGER) || is(T == ULARGE_INTEGER)) {
  enum opDispatchMarshalParam = q{
    Params[$i] param$i = {QuadPart: args[$i]};
  };
}

package template opDispatchMarshalParam(T)
  if (is(T == LARGE_INTEGER*) || is(T == ULARGE_INTEGER*)) {
  enum opDispatchMarshalParam = q{
    PointerTarget!(Params[$i]) temp$i = {QuadPart: *args[$i]};
    auto param$i = &temp$i;
    scope(exit) *args[$i] = (*param$i).QuadPart;
  };
}

package enum opDispatchMarshalParam(T : IID*) = q{
  auto param$i = &args[$i];
};

package enum opDispatchMarshalParam(T : IDispatch) = q{
  auto param$i = new class IDispatch {
    mixin Dispatch;

    @dispId(DISPID_VALUE)
    extern(D) ReturnType!(Args[$i]) invoke(Parameters!(Args[$i]) a) { return args[$i](a); }
  };
};

package template opDispatchMarshalParam(T)
  if (isInstanceOf!(TaskBuffer, T)) {
  enum opDispatchMarshalParam = q{
    auto param$i = args[$i].ptr;
  };
}

package enum opDispatchMarshalReturn(T : BSTR) = q{
  size_t returnParamLength;
  auto isReturnParamBSTR = validateBSTR(returnParam, returnParamLength);
  scope(exit) if (isReturnParamBSTR) SysFreeString(returnParam);
  return returnParam[0 .. returnParamLength].toUTF8();
};

static if (is(HSTRING))
package enum opDispatchMarshalReturn(T : HSTRING) = q{
  uint returnParamLength;
  auto p = WindowsGetStringRawBuffer(returnParam, &returnParamLength);
  return !!p ? p[0 .. returnParamLength].toUTF8() : null;
};

package template opDispatchMarshalReturn(T)
  if (is(T == LARGE_INTEGER) || is(T == ULARGE_INTEGER)) {
  enum opDispatchMarshalReturn = q{
    return returnParam.QuadPart;
  };
}

package enum opDispatchMarshalReturn(T : VARIANT) = q{
  scope(exit) variantClear(returnParam);
  return variantAs(returnParam, Variant(returnParam));
};

package enum opDispatchMarshalReturn(T : IUnknown) = q{
  return returnParam.asUnknownPtr();
};

package {

  import core.sys.windows.ocidl,
    core.sys.windows.oleauto,
    core.sys.windows.uuid,
    std.utf;
  import comet.core : checkHR, isSuccess;

  string typeName(ITypeInfo typeInfo) @property {
    wchar* typeName;
    checkHR(typeInfo.GetDocumentation(MEMBERID_NIL, &typeName, null, null, null));
    scope(exit) SysFreeString(typeName);
    return typeName[0 .. SysStringLen(typeName)].toUTF8();
  }

  string memberName(ITypeInfo typeInfo, int memberId) {
    wchar* memberName;
    uint count;
    checkHR(typeInfo.GetNames(memberId, &memberName, 1, &count));
    scope(exit) SysFreeString(memberName);
    return memberName[0 .. SysStringLen(memberName)].toUTF8();
  }

  class MethodDescription {

    string name;
    int dispId;
    INVOKEKIND invokeKind;
    short paramCount;

    this(ITypeInfo typeInfo, FUNCDESC* funcDesc) {
      name = typeInfo.memberName(funcDesc.memid);
      dispId = funcDesc.memid;
      invokeKind = funcDesc.invkind;
      paramCount = funcDesc.cParams;
    }

  }

  class EventDescription {

    int dispId;
    GUID sourceIid;

    this(int dispId, GUID sourceIid) {
      this.dispId = dispId;
      this.sourceIid = sourceIid;
    }

  }

  class TypeDescription {

    string name;
    GUID guid;
    MethodDescription[string] methods, setters;
    EventDescription[string] events;

    this(ITypeInfo typeInfo) { name = typeInfo.typeName; }

    static fromTypeInfo(ITypeInfo typeInfo, TYPEATTR* typeAttr) {
      switch (typeAttr.typekind) {
      case TYPEKIND.TKIND_COCLASS:
        return new ClassTypeDescription(typeInfo);
      case TYPEKIND.TKIND_INTERFACE, TYPEKIND.TKIND_DISPATCH:
        return new TypeDescription(typeInfo);
      default:
        assert(false);
      }
    }

  }

  class ClassTypeDescription : TypeDescription {

    import std.container : SList;
    import std.algorithm : canFind;

    SList!string sourceInterfaces, interfaces;

    this(ITypeInfo typeInfo) {
      super(typeInfo);

      TYPEATTR* typeAttr;
      checkHR(typeInfo.GetTypeAttr(&typeAttr));
      scope(exit) typeInfo.ReleaseTypeAttr(typeAttr);

      guid = typeAttr.guid;

      foreach (i; 0 .. typeAttr.cImplTypes) {
        uint refType;
        checkHR(typeInfo.GetRefTypeOfImplType(i, &refType));

        ITypeInfo refInfo;
        checkHR(typeInfo.GetRefTypeInfo(refType, &refInfo));
        scope(exit) refInfo.Release();

        auto interfaceName = refInfo.typeName;

        int implTypeFlags;
        checkHR(typeInfo.GetImplTypeFlags(i, &implTypeFlags));
        if ((implTypeFlags & IMPLTYPEFLAG_FSOURCE) != 0)
          sourceInterfaces.insertAfter(sourceInterfaces[], interfaceName);
        else
          interfaces.insertAfter(interfaces[], interfaceName);
      }
    }

    bool implements(string name, bool isSource) {
      return isSource ? sourceInterfaces[].canFind(name) : interfaces[].canFind(name);
    }

  }

  class TypeLibDescription {

    import std.container : SList;

    static TypeLibDescription[GUID] cache;
    SList!ClassTypeDescription classes;

    static fromTypeLib(ITypeLib typeLib) {
      TLIBATTR* libAttr;
      checkHR(typeLib.GetLibAttr(&libAttr));
      scope(exit) typeLib.ReleaseTLibAttr(libAttr);

      if (auto cached = cache.get(libAttr.guid, null))
        return cached;

      auto result = new TypeLibDescription;

      foreach (i; 0 .. typeLib.GetTypeInfoCount()) {
        TYPEKIND kind;
        checkHR(typeLib.GetTypeInfoType(i, &kind));

        if (kind == TYPEKIND.TKIND_COCLASS) {
          ITypeInfo typeInfo;
          checkHR(typeLib.GetTypeInfo(i, &typeInfo));
          scope(exit) typeInfo.Release();

          result.classes.insertAfter(result.classes[], new ClassTypeDescription(typeInfo));
        }
      }

      return cache[libAttr.guid] = result;
    }

  }

  auto scanForMethods(IDispatch obj, ref TypeDescription[GUID] cache) {
    ITypeInfo typeInfo;
    if (!obj.GetTypeInfo(0, 0, &typeInfo).isSuccess) return null;
    scope(exit) typeInfo.Release();

    TYPEATTR* typeAttr;
    checkHR(typeInfo.GetTypeAttr(&typeAttr));
    scope(exit) typeInfo.ReleaseTypeAttr(typeAttr);

    TypeDescription result;
    if (auto cached = cache.get(typeAttr.guid, null)) {
      if ((result = cached).methods != null)
        return result;
    }

    MethodDescription[string] methods_, setters_;

    foreach (i; 0 .. typeAttr.cFuncs) {
      FUNCDESC* funcDesc;
      checkHR(typeInfo.GetFuncDesc(i, &funcDesc));
      scope(exit) typeInfo.ReleaseFuncDesc(funcDesc);

      if ((funcDesc.wFuncFlags & FUNCFLAGS.FUNCFLAG_FRESTRICTED) != 0) continue;

      auto method = new MethodDescription(typeInfo, funcDesc);
      if ((funcDesc.invkind & (INVOKEKIND.INVOKE_PROPERTYPUT | INVOKEKIND.INVOKE_PROPERTYPUTREF)) != 0)
        setters_[method.name.toLower()] = method;
      else
        methods_[method.name.toLower()] = method;
    }

    if (methods_ != null || setters_ != null) {
      with (result = cache.require(typeAttr.guid, TypeDescription.fromTypeInfo(typeInfo, typeAttr))) {
        methods = methods_;
        setters = setters_;
      }
    }
    return result;
  }

  auto scanForEvents(IDispatch obj, ref TypeDescription[GUID] cache) {
    ITypeInfo sourceInfo;
    if (!obj.GetTypeInfo(0, 0, &sourceInfo).isSuccess) return null;
    scope(exit) sourceInfo.Release();

    TYPEATTR* typeAttr;
    checkHR(sourceInfo.GetTypeAttr(&typeAttr));
    scope(exit) sourceInfo.ReleaseTypeAttr(typeAttr);

    TypeDescription result;
    if (auto cached = cache.get(typeAttr.guid, null)) {
      if ((result = cached).events != null)
        return result;
    }

    auto coClassTypeInfo() {
      ITypeInfo typeInfo;

      IProvideClassInfo classInfo;
      if (sourceInfo.QueryInterface(&IID_IProvideClassInfo, cast(void**)&classInfo).isSuccess) {
        scope(exit) classInfo.Release();
        if (classInfo.GetClassInfo(&typeInfo).isSuccess)
          return typeInfo;
      }

      ITypeLib sourceLib;
      uint count;
      checkHR(sourceInfo.GetContainingTypeLib(&sourceLib, &count));
      scope(exit) sourceLib.Release();

      auto name = sourceInfo.typeName;
      auto typeLib = TypeLibDescription.fromTypeLib(sourceLib);

      ClassTypeDescription coClass;
      foreach (c; typeLib.classes[]) {
        if (c.implements(name, /*isSource*/ false)) {
          coClass = c;
          break;
        }
      }
      if (coClass is null) return null;

      checkHR(sourceLib.GetTypeInfoOfGuid(&coClass.guid, &typeInfo));
      return typeInfo;
    }

    EventDescription[string] events_;

    IConnectionPointContainer container;
    if (obj.QueryInterface(&IID_IConnectionPointContainer, cast(void**)&container).isSuccess) {
      scope(exit) container.Release();

      if (auto classInfo = coClassTypeInfo()) {
        scope(exit) classInfo.Release();

        TYPEATTR* classAttr;
        checkHR(classInfo.GetTypeAttr(&classAttr));
        scope(exit) classInfo.ReleaseTypeAttr(classAttr);

        foreach (i; 0 .. classAttr.cImplTypes) {
          uint interfaceType;
          checkHR(classInfo.GetRefTypeOfImplType(i, &interfaceType));

          ITypeInfo interfaceInfo;
          checkHR(classInfo.GetRefTypeInfo(interfaceType, &interfaceInfo));
          scope(exit) interfaceInfo.Release();

          int implTypeFlags;
          checkHR(classInfo.GetImplTypeFlags(i, &implTypeFlags));
          if ((implTypeFlags & IMPLTYPEFLAG_FSOURCE) == 0) continue;

          TYPEATTR* interfaceAttr;
          checkHR(interfaceInfo.GetTypeAttr(&interfaceAttr));
          scope(exit) interfaceInfo.ReleaseTypeAttr(interfaceAttr);

          foreach (j; 0 .. interfaceAttr.cFuncs) {
            FUNCDESC* funcDesc;
            checkHR(interfaceInfo.GetFuncDesc(j, &funcDesc));
            scope(exit) interfaceInfo.ReleaseFuncDesc(funcDesc);

            if ((funcDesc.wFuncFlags & FUNCFLAGS.FUNCFLAG_FRESTRICTED) != 0) continue;
            if ((funcDesc.wFuncFlags & FUNCFLAGS.FUNCFLAG_FHIDDEN) != 0) continue;

            auto name = interfaceInfo.memberName(funcDesc.memid).toLower();
            events_.require(name, new EventDescription(funcDesc.memid, interfaceAttr.guid));
          }
        }
      }
    }

    if (events_ != null) {
      with (result = cache.require(typeAttr.guid, TypeDescription.fromTypeInfo(sourceInfo, typeAttr))) {
        events = events_;
      }
    }
    return result;
  }

}