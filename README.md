# COMet

COMet is a library for the D programming language that aims to make COM/WinRT programming easy.

At the heart of COMet is a super-smart pointer, `UnknownPtr`, which manages the lifetime of IUnknown-based objects; generates helper methods at compile-time allowing you to use more a more natural programming style; throws exceptions instead of returning HRESULTs; connects delegates to event sources; supports foreach-style enumeration and indexing; lets you use normal type casts in place of QueryInterface.

In addition, the library provides default implementations of `IUnknown`, `IDispatch` and `IInspectable` for when you need to implement COM interfaces.

## UnknownPtr

### Marshaling

`UnknownPtr` hides much of the complexity of converting types to and from ones COM recognises such as VARIANTs and BSTRs. Most of the time, you can just use D's built-in types like bool, string, arrays, even delegates. For example, to call a method that receives a pointer to a BSTR, you can pass the address of a string instead:

```d
UnknownPtr!IUri uri = ...;
string absoluteUri;
uri.GetAbsoluteUri(&absoluteUri);
writeln(absoluteUri);
```

where `GetAbsoluteUri` is defined as an interface method in urlmon.d as:

```d
interface IUri : IUnknown {
  ... // other methods
  HRESULT GetAbsoluteUri(BSTR* pbstrAbsoluteUri);
  ... // other methods
}
```

### Methods

The convention in COM is to treat the last parameter in a method's argument list as the return value (because most COM methods actually return an HRESULT value indicating success or failure). So `UnknownPtr` uses D's metaprogramming abilities to determine if that's the case and generates a wrapper function that returns the last argument (the HRESULT gets converted into an exception). The example above can then be rewritten:

```d
string absoluteUri = uri.GetAbsoluteUri();
```

You can also use camel casing so the call feels more at home in D:

```d
string absoluteUri = uri.getAbsoluteUri();
```

### Properties

Another COM convention is declaring property getters and setters with **get_** and **put_** prefixes, but `UnknownPtr` allows you to omit these and call accessors using the property's name:

```d
import adoint.d;

UnknownPtr!_ADOCommand command = ...;
command.commandText = "select * from People"; // Calls _ADOCommand.put_CommandText
int recordsAffected;
UnknownPtr!_ADORecordSet recordSet = command.Execute(&recordsAffected, null, 0);
writeln("recordCount: ", recordSet.recordCount); // Calls _ADORecordSet.get_RecordCount
```

### Early vs late binding

Much of this magic is accomplished by `UnknownPtr` implementing **opDispatch**. And, as you might expect given the name, it also takes the pain out of working with IDispatch-based objects, which support late binding so no interface definitions are required. In plain COM you'd need to query `GetIDsOfNames` for the identifier of a member and use that when you `Invoke` the function (with much code to convert arguments into VARIANTs and DISPPARAMS). However, `UnknownPtr` generates all that code on your behalf. So to enable late binding, the example above could be rewritten with a couple of changes:

```d
// No need to import adoint.d for late binding
UnknownPtr!IDispatch command = ...;
command.commandText = "select * from People";
int recordsAffected;
auto recordSet = command.Execute(&recordsAffected, null, 0);
writeln("recordCount: ", cast(int)recordSet.recordCount);
```

(The cast is needed because late-bound methods return VARIANTs, which are wrapped by COMet into a struct that defines **opCast**.)

### Enumeration

Collections in COM tend to share a common interface, so `UnknownPtr` uses introspection to determine whether the contained interface either supports enumeration or is itself an enumerator, and generates the appropriate **opApply** functions and element conversions. Then you can simply foreach over the smart pointer:

```d
import spellcheck.d;

UnknownPtr!ISpellChecker spellChecker = ...;
foreach (suggestion; spellChecker.suggest("acheive")) {
  writeln("Did you mean ", suggestion, "?");
}
```

Similarly, **opIndex** will be generated to enable array-like access (currently limited to a simgle dimension).

### Hooked on events

Various event-handling mechanisms are supported: simple `IDispatch`-based handlers, connectable objects (i.e., `IConnectionPointContainer`), and specific event-handling interfaces you would subscribe and unsubscribe via **add_** and **remove_** prefixed methods. `UnknownPtr` unifies these so that you just supply a delegate without needing to write any plumbing code. For example, some objects support simple property assignment:

```d
import msxml6.d;

UnknownPtr!IXMLDOMDocument doc = ...;
doc.onReadyStateChange = () => writeln("readyState: " cast(int)doc.readyState);
```

But most of the time you'll be hooking up to event sources with the `+=` operator (or `~=` if you prefer):

```d
UnknownPtr!IDispatch doc = ...;
doc.onDataAvailable += () => writeln(cast(string)doc.xml);
doc.load("books.xml");
```

## Creating and authoring 

### Make

The counterpart to **new** in the COM world is `CoCreateInstance`. In COMet, you call `make`, specifiying the identifier (CLSID) associated with the class you want to instantiate:

```d
IUnknown doc1 = make(CLSID_DOMDocument60);
// or pass the CLSID in string form
IUnknown doc2 = make("88d96a05-f192-11d4-a65f-0040963251e5");
// or the ProgID
IUnknown doc3 = make("MSXML2.DOMDocument");
```

The `make` function returns IUnknown by default, although you can specify the type as a template parameter, eg `make!IXMLDOMDocument("MSXML2.DOMDocument")`. But this isn't necessary if you assign the result to an `UnknownPtr`, enabling all the magic it provides:

```d
Unknown!IXMLDOMDocument doc = make("MSXML2.DOMDocument");
```

### Implementing interfaces

In COMet, instead of deriving from a base class that implements IUnknown's required methods, you use composition by mixing in one of the provided templates: `Unknown`, `Dispatch`, or `Inspectable`. (COMet uses composition over inheritance because D's single inheritance model means you couldn't extend your classes and implement additional interfaces as well.)

```d
class XMLHTTPRequestCallback : IXMLHTTPRequest2Callback {
  mixin Unknown;

  HRESULT OnRedirect(IXMLHTTPRequest2 request, const(WCHAR)* redirectUrl) { return S_OK; }
  HRESULT OnHeadersAvailable(IXMLHTTPRequest2 request, uint status, const(WCHAR)* statusText) { return S_OK; }
  HRESULT OnDataAvailable(IXMLHTTPRequest2 request, ISequentialStream responseStream) { return S_OK; }
  HRESULT OnResponseReceived(IXMLHTTPRequest2 request, ISequentialStream responseStream) { return S_OK; }
  HRESULT OnError(IXMLHTTPRequest2 request, HRESULT error) { return S_OK; }
}

UnknownPtr!IXMLHTTPRequest2 request = ...;
request.open("GET", "https://dlang.org/index.html", new XMLHTTPRequestCallback, null, null, null, null);

```

The `Unknown` template contains all the code for QueryInterface, AddRef and Release. You just need to implement `IXMLHTTPRequest2Callback`'s methods - the ones you actually care about. However, it doesn't look especially nice, littered as it is with C types and HRESULTs.

Enter the `Marshaling` template: you write normal D-style code, mixin the template and it will generate methods so your class conforms to the interface, converting arguments to expected types and exceptions to HRESULT values. You can even omit methods you don't want to handle, leaving COMet to insert a stub.

```d
class XMLHTTPRequestCallback : IXMLHTTPRequest2Callback {
  mixin Unknown;
  mixin Marshaling;

  void onRedirect(IXMLHTTPRequest2 request, string redirectUrl) {}
  void onHeadersAvailable(IXMLHTTPRequest2 request, uint status, string statusText) {}
  void onDataAvailable(IXMLHTTPRequest2 request, ISequentialStream responseStream) {}
}
```

To round off this introduction, here's a few code samples.

## Code examples