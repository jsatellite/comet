module comet.utils;

import std.traits,
  std.ascii;

template nameOf(alias T) {
  static if (__traits(compiles, __traits(identifier, T)))
    enum nameOf = __traits(identifier, T);
  else
    enum nameOf = T.stringof;
}

// static pattern matching
auto match(alias T, cases...)() {
  bool conditionToSuppressStatementNotReachableWarning = true;

  static foreach (i, _; cases) {
    static if (i % 2 == 1) {
      static if (is(typeof(cases[i - 1]) == bool)) {
        static if (cases[i - 1]) {
          if (conditionToSuppressStatementNotReachableWarning) 
            return cases[i]();
        }
      } else static if (is(Unqual!T == cases[i - 1])) {
        if (conditionToSuppressStatementNotReachableWarning) 
          return cases[i]();
      } else static if (is(typeof(T) == typeof(cases[i - 1]))) {
        static if (T == cases[i - 1]) {
          if (conditionToSuppressStatementNotReachableWarning) 
            return cases[i]();
        }
      }
    }
  }

  static if (cases.length % 2 == 1) {
    static if (isFunctionPointer!(typeof(cases[$ - 1]()))) {
      return cases[$ - 1]()();
    } else {
      return cases[$ - 1]();
    }
  } else {
    static assert(false, "no matches for `" ~ T.stringof ~ "`");
  }
}

private string commaExpression(string s) {
  if (s.length == 0)
    return null;

  if (s[0] == '$' && (s.length < 2 || s[1] != '$'))
    return injectExpression(s[1 .. $]);

  string result = `"`;
  size_t mark = 0, index = 0;

  while (true) {
    auto c = s[index++];
    if (c == '$') {
      if (index >= s.length || s[index] != '$')
        return result ~ s[mark .. index - 1] ~ `", ` ~ injectExpression(s[index .. $]);

      result ~= s[mark .. index++];
      mark = index;
    } else if (c == '"' || c == '\\') {
      result ~= s[mark .. index - 1] ~ '\\';
      mark = index - 1;
    }

    if (index >= s.length)
      return result ~ s[mark .. index] ~ `"`;
  }
}

private string injectExpression(string s) {
  if (s[0] == '(') {
    uint depth = 1;
    size_t i = 1;
    for (;; i++) {
      if (i >= s.length)
        assert(false, "missing close delimiter `)` for injected expression started with `$(`");

      if (s[i] == ')') {
        depth--;
        if (depth == 0) break;
      } else if (s[i] == '(') {
        depth++;
      }
    }

    if (i == s.length - 1)
      return s[1 .. i];

    return s[1 .. i] ~ `,` ~ commaExpression(s[i + 1 .. $]);
  } else {
    //assert(false, "missing open delimiter `(` for injected expression");
    size_t i = 0;
    for (;; i++) {
      if (i >= s.length || !isAlpha(s[i]))
        break;
    }

    if (i == s.length)
      return s[0 .. i];

    return s[0 .. i] ~ `,` ~ commaExpression(s[i .. $]);
  }
}

string inject(string s) {
  // useage: mixin("Hello $(x + y)".inject);
  // or: mixin("Hello $x".inject);
  return "(){import std.conv:text;import std.typecons:tuple;return text(tuple(" ~ commaExpression(s) ~ ").expand);}()";
}

struct TaskBuffer(T) {

  import core.sys.windows.objbase,
    std.exception;

  private T* ptr_;
  private size_t size_;

  this(size_t size) {
    enforce((ptr_ = cast(T*)CoTaskMemAlloc((size_ = size) * T.sizeof)) !is null);
  }

  ~this() {
    if (ptr_ !is null) {
      CoTaskMemFree(ptr_);
      ptr_ = null;
    }
  }

  ref T* ptr() @property { return ptr_; }

  size_t size() @property { return size_; }
  alias opDollar = size;

  auto opIndex(size_t index) { return ptr_[index]; }
  auto opIndex() { return ptr_[0 .. size_]; }
  auto opSlice(size_t x, size_t y) { return ptr_[x .. y]; }

}