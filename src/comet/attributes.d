///
module comet.attributes;

import core.sys.windows.w32api,
  core.sys.windows.basetyps : GUID;

private immutable guidFormat_ = "Input string for GUID was not in a correct format";

/**
 * Initializes _a new GUID instance with the specified data.
 */
GUID guid(ubyte[16] b) {
  return GUID(b[3] << 24 | b[2] << 16 | b[1] << 8 | b[0],
              b[5] << 8 | b[4],
              b[7] << 8 | b[6],
              [b[8], b[9], b[10], b[11], b[12], b[13], b[14], b[15]]);
}

/// ditto
GUID guid(uint a, ushort b, ushort c, ubyte[8] d) { return GUID(a, b, c, d); }

private ulong parse(string s) {

  bool hexToInt(char c, out uint result) {
    if (c >= '0' && c <= '9') result = c - '0';
    else if (c >= 'A' && c <= 'F') result = c - 'A' + 10;
    else if (c >= 'a' && c <= 'f') result = c - 'a' + 10;
    else result = -1;
    return cast(int)result >= 0;
  }

  ulong result;
  uint value, index, length = cast(uint)s.length;
  while (index < length && hexToInt(s[index++], value)) {
    result = result * 16 + value;
  }
  return result;
}

/// ditto
GUID guid(string s) {
  import std.string, std.exception;

  s = s.strip();
  enforce(!s.empty, guidFormat_);

  if (s[0] == '{') {
    enforce(s.length == 38 && s[37] == '}');
    s = s[1 .. 37];
  } else if (s[0] == '(') {
    enforce(s.length == 38 && s[37] == ')');
    s = s[1 .. 37];
  }

  enforce(s[8] == '-' && s[13] == '-' && s[18] == '-' && s[23] == '-', guidFormat_);

  GUID g;

  with (g) {
    Data1 = cast(uint)parse(s[0 .. 8]);
    Data2 = cast(ushort)parse(s[9 .. 13]);
    Data3 = cast(ushort)parse(s[14 .. 18]);
    auto a = cast(uint)parse(s[19 .. 23]);
    Data4[0] = cast(ubyte)(a >> 8);
    Data4[1] = cast(ubyte)a;
    auto b = parse(s[24 .. $]);
    a = cast(uint)(b >> 32);
    Data4[2] = cast(ubyte)(a >> 8);
    Data4[3] = cast(ubyte)a;
    a = cast(uint)b;
    Data4[4] = cast(ubyte)(a >> 24);
    Data4[5] = cast(ubyte)(a >> 16);
    Data4[6] = cast(ubyte)(a >> 8);
    Data4[7] = cast(ubyte)a;
  }

  return g;
}

private void hexToString(ref char[] s, ref size_t index, uint a, uint b) {

  char hexToChar(uint a) {
    a = a & 0x0F;
    return cast(char)((a > 9) ? a - 10 + 0x61 : a + 0x30);
  }

  s[index++] = hexToChar(a >> 4);
  s[index++] = hexToChar(a);
  s[index++] = hexToChar(b >> 4);
  s[index++] = hexToChar(b);
}

/**
 Returns a string representation of the specified GUID value.
 Params:
   g = The GUID whose value is to be represented as a string.
   fmt = The format specifier, which can be "D", "B" or "P". If null or empty, then "D" is used.
 */
string toString(GUID g, string fmt = "D") {
  import std.exception : assumeUnique;

  if (fmt == null) fmt = "D";

  size_t index;
  char[] s;
  if (fmt == "D" || fmt == "d") {
    s.length = 36;
  } else if (fmt == "B" || fmt == "b") {
    s.length = 38;
    s[index++] = '{';
    s[$ - 1] = '}';
  } else if (fmt == "P" || fmt == "p") {
    s.length = 38;
    s[index++] = '(';
    s[$ - 1] = ')';
  } else {
    throw new Exception("Format string must be one of `D`, `d`, `B`, `b`, `P` or `p`");
  }

  hexToString(s, index, g.Data1 >> 24, g.Data1 >> 16);
  hexToString(s, index, g.Data1 >> 8, g.Data1);
  s[index++] = '-';
  hexToString(s, index, g.Data2 >> 8, g.Data2);
  s[index++] = '-';
  hexToString(s, index, g.Data3 >> 8, g.Data3);
  s[index++] = '-';
  hexToString(s, index, g.Data4[0], g.Data4[1]);
  s[index++] = '-';
  hexToString(s, index, g.Data4[2], g.Data4[3]);
  hexToString(s, index, g.Data4[4], g.Data4[5]);
  hexToString(s, index, g.Data4[6], g.Data4[7]);
  return s.assumeUnique();
}

/**
 Returns a 16-element array of bytes that contains the GUID value.
 */
ubyte[16] data(GUID g) @property {
  with (g) return [
    cast(ubyte)Data1, 
    cast(ubyte)(Data1 >> 8), 
    cast(ubyte)(Data1 >> 16), 
    cast(ubyte)(Data1 >> 24),
    cast(ubyte)Data2, 
    cast(ubyte)(Data2 >> 8), 
    cast(ubyte)Data3, 
    cast(ubyte)(Data3 >> 8),
    Data4[0], 
    Data4[1],
    Data4[2], 
    Data4[3],
    Data4[4], 
    Data4[5], 
    Data4[6], 
    Data4[7]
  ];
}

struct overload {
  string method;
}

/**
 Specifies the dispatch identifier (DISPID) of a method.
*/
struct dispId {
  int value;
}

struct notMarshaled {}

/**
 Indicates that the HRESULT signature transformation that occurs during COM calls should be suppressed.
*/
struct keepReturn {}

/**
 Identifies how to marshal parameters to and from COM.
 */
enum MarshalingType {
  bool_ = 1,  /// A 4-byte Win32 BOOL type.
  bstr,       /// A Unicode string, the default COM string type.
  lpstr,      /// A single-byte _null-terminated ANSI string.
  lpwstr,     /// A 2-byte _null-terminated Unicode string.
  variantBool /// A 2-byte VARIANT_BOOL type.
}

/**
 Indicates how to marshal the data to and from COM.
 */
struct marshalAs {
  /// Indicates the type the data is to be marshaled as.
  MarshalingType value;
}

/// ditto
struct marshalReturnAs {
  MarshalingType value;
}