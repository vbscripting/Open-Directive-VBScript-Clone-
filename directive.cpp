// ============================================================================
//  Directive  (Open Directive / Directive Script)
//  A small scripting language with VBScript-like syntax, in one C++17 file.
//
//  - Lexer + recursive-descent parser + tree-walking evaluator
//  - Variant-style dynamic value type (Empty/Null/Bool/Int/Double/String/
//    Object/Array) with VBScript-style coercion rules
//  - ~70 built-in functions, classes, error handling (On Error / Err),
//    Sub/Function with ByRef/ByVal, Select Case, Do/For/While, With, etc.
//  - Native objects created with New:  Dictionary, List (linked list),
//    RegExp, FileSystem (alias FSO) + File/Folder objects.  They work on
//    every platform.
//  - CreateObject("...") is reserved for REAL ActiveX/COM automation via
//    IDispatch on Windows (see #ifdef _WIN32 section): CreateObject("Excel.Application").
//  - The host object is "Directive": Directive.Echo, Directive.StdOut.Write,
//    Directive.StdIn.ReadLine, Directive.Quit, ...
//
//  Build (console, cscript-style stdio):
//    Linux/macOS : g++ -std=c++17 -O2 directive.cpp -o directive
//    Windows MSVC: cl /std:c++17 /EHsc directive.cpp
//    Windows MinGW: g++ -std=c++17 directive.cpp -o directive.exe -lole32 -loleaut32 -luuid
//  Build (GUI, wscript-style dialogs) - add -DDIRECTIVE_GUI and a windowed subsystem:
//    Windows MSVC : cl /std:c++17 /EHsc /DDIRECTIVE_GUI directive.cpp /Fe:directivew.exe /link /SUBSYSTEM:WINDOWS
//    Windows MinGW: g++ -std=c++17 -DDIRECTIVE_GUI directive.cpp -o directivew.exe -mwindows -lole32 -loleaut32 -luuid
//
//  Run:  ./directive script.directive      (or:  ./directive -e "Directive.Echo 1+2")
//
//  License: MIT.

// ============================================================================
#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include <map>
#include <chrono>
#include <list>
#include <iterator>
#include <functional>
#include <variant>
#include <cstdint>
#include <cmath>
#include <cctype>
#include <sstream>
#include <iostream>
#include <fstream>
#include <algorithm>
#include <stdexcept>
#include <ctime>
#include <regex>
#include <set>
#include <filesystem>
namespace fs = std::filesystem;
#include <cerrno>
#include <iomanip>

// ------------------------------------------------------------------ platform libs
#if defined(_WIN32) && defined(_MSC_VER)
  #pragma comment(lib, "ole32.lib")
  #pragma comment(lib, "oleaut32.lib")
  #pragma comment(lib, "user32.lib")
  #pragma comment(lib, "gdi32.lib")
  #pragma comment(lib, "winmm.lib")
  #ifdef DIRECTIVE_GUI
    #pragma comment(lib, "shell32.lib")
  #endif
#endif

// ------------------------------------------------------------------ ui layer
// Console builds route MsgBox/InputBox/Echo through stdio (like cscript.exe).
// GUI builds (compiled with -DDIRECTIVE_GUI on Windows) show real dialogs (like wscript.exe).
// Definitions live at the bottom of the file, next to the Win32 code.
namespace ui {
    void echo(const std::string& text);                                  // WScript.Echo
    long msgBox(const std::string& prompt, long flags, const std::string& title);   // returns button code
    bool inputBox(const std::string& prompt, const std::string& title,
                  const std::string& def, std::string& out);             // false == cancelled
    void error(const std::string& text);                                 // fatal error surface
}

// ------------------------------------------------------------------ utilities
static std::string toLower(const std::string& s) {
    std::string r = s;
    std::transform(r.begin(), r.end(), r.begin(),
                   [](unsigned char c){ return (char)std::tolower(c); });
    return r;
}

// ------------------------------------------------------------------ forwards
class Interpreter;
struct IObject;
using ObjectPtr = std::shared_ptr<IObject>;

struct ArrayData;
using ArrayPtr = std::shared_ptr<ArrayData>;

// -------------------------------------------------------------- error / signal
// A VBScript runtime error. Number/description map onto the Err object.
struct VbError {
    long number;
    std::string source;
    std::string description;
    long line = 0;             // source line where the error occurred (0 = unknown)
    long col  = 0;             // source column (0 = unknown)
    VbError(long n, std::string d, std::string src = "Directive runtime error")
        : number(n), source(std::move(src)), description(std::move(d)) {}
};
[[noreturn]] static void raiseErr(long n, const std::string& d) { throw VbError(n, d); }

// Non-local control flow used to implement Exit Sub/Function/For/Do/Property.
struct ExitSignal { enum K { Sub, Function, For, Do, Property, DoAll } kind; };

// ---------------------------------------------------------------- Value type
// The dynamic "Variant". Kept as plain members (not std::variant) so the rest
// of the code reads clearly. Object==null pointer with type Object encodes the
// VBScript value "Nothing".
class Value {
public:
    enum class Type { Empty, Null, Bool, Int, Double, String, Object, Array };
    Type type = Type::Empty;

    bool        b = false;
    int32_t     i = 0;
    double      d = 0.0;
    std::string s;
    ObjectPtr   obj;   // Object; nullptr => Nothing
    ArrayPtr    arr;

    Value() = default;

    static Value empty()            { return Value(); }
    static Value null()             { Value v; v.type = Type::Null; return v; }
    static Value nothing()          { Value v; v.type = Type::Object; v.obj = nullptr; return v; }
    static Value boolean(bool x)    { Value v; v.type = Type::Bool; v.b = x; return v; }
    static Value integer(int32_t x) { Value v; v.type = Type::Int; v.i = x; return v; }
    static Value number(double x)   { Value v; v.type = Type::Double; v.d = x; return v; }
    static Value str(std::string x) { Value v; v.type = Type::String; v.s = std::move(x); return v; }
    static Value object(ObjectPtr o){ Value v; v.type = Type::Object; v.obj = std::move(o); return v; }
    static Value array(ArrayPtr a)  { Value v; v.type = Type::Array; v.arr = std::move(a); return v; }

    bool isEmpty()  const { return type == Type::Empty; }
    bool isNull()   const { return type == Type::Null; }
    bool isObject() const { return type == Type::Object; }
    bool isArray()  const { return type == Type::Array; }
    bool isNothing()const { return type == Type::Object && !obj; }

    // ---- coercions (VBScript semantics) ----
    bool toBool() const;
    double toDouble() const;
    int32_t toInt() const;            // round-half-to-even, then int32
    long long toI64() const;
    std::string toStr() const;        // used by CStr / display
    std::string toConcatStr() const;  // used by & (Null -> "")
    bool looksNumeric() const;        // for IsNumeric
    std::string vbTypeName() const;   // for TypeName()
    int varType() const;              // for VarType()
};

// -------------------------------------------------------------- ArrayData
// N-dimensional array with 0-based lower bounds (VBScript arrays are 0-based).
// `uppers[k]` is the inclusive upper bound of dimension k.
struct ArrayData {
    std::vector<int> uppers;
    std::vector<Value> data;

    size_t total() const {
        size_t n = 1;
        for (int u : uppers) n *= (size_t)(u + 1);
        return n;
    }
    size_t flatIndex(const std::vector<int>& idx) const {
        if (idx.size() != uppers.size())
            raiseErr(9, "Subscript out of range (wrong number of dimensions)");
        size_t flat = 0;
        for (size_t k = 0; k < uppers.size(); ++k) {
            if (idx[k] < 0 || idx[k] > uppers[k])
                raiseErr(9, "Subscript out of range");
            flat = flat * (size_t)(uppers[k] + 1) + (size_t)idx[k];
        }
        return flat;
    }
    void redim(const std::vector<int>& newUppers, bool preserve) {
        if (!preserve) {
            uppers = newUppers;
            data.assign(total(), Value::empty());
            return;
        }
        // Preserve: VBScript only allows resizing the last dimension. We copy
        // element-by-element for the overlapping region (works for 1-D, which
        // is the overwhelmingly common case).
        ArrayData old = *this;
        uppers = newUppers;
        std::vector<Value> nd(total(), Value::empty());
        if (uppers.size() == old.uppers.size()) {
            std::vector<int> idx(uppers.size(), 0);
            std::function<void(size_t)> rec = [&](size_t dim) {
                if (dim == uppers.size()) {
                    bool inOld = true;
                    for (size_t k = 0; k < idx.size(); ++k)
                        if (idx[k] > old.uppers[k]) { inOld = false; break; }
                    if (inOld) nd[flatIndex(idx)] = old.data[old.flatIndex(idx)];
                    return;
                }
                for (int v = 0; v <= uppers[dim]; ++v) { idx[dim] = v; rec(dim + 1); }
            };
            if (total() > 0) rec(0);
        }
        data.swap(nd);
    }
};

// -------------------------------------------------------------- IObject
// Everything you can put a "." after implements this. Both script-defined
// classes and COM/native objects use it, so member access is uniform.
struct IObject : std::enable_shared_from_this<IObject> {
    virtual ~IObject() = default;
    virtual std::string typeName() const { return "Object"; }

    // Property-get OR method-call-returning-value (COM merges these; we do too).
    virtual Value get(Interpreter& in, const std::string& name, std::vector<Value>& args) = 0;

    // Property assignment: `obj.Name = v` (isSet=false, "Let") or
    // `Set obj.Name = v` (isSet=true).
    virtual void set(Interpreter& in, const std::string& name,
                     std::vector<Value>& args, const Value& val, bool isSet) {
        (void)in;(void)args;(void)val;(void)isSet;
        raiseErr(438, "Object doesn't support this property or method: " + name);
    }

    // Default property/value, for `x = obj` or coercions. Return false if none.
    virtual bool tryDefault(Interpreter& in, Value& out) { (void)in;(void)out; return false; }

    // For..Each enumeration. Fill `out` with the elements. Return false if the
    // object is not enumerable.
    virtual bool enumerate(Interpreter& in, std::vector<Value>& out) {
        (void)in;(void)out; return false;
    }
};

// -------------------------------------------------- number formatting helpers
static std::string numToStr(double v) {
    if (std::isnan(v)) return "NaN";
    if (std::isinf(v)) return v < 0 ? "-1.#INF" : "1.#INF";
    if (std::floor(v) == v && std::fabs(v) < 1e15)
        return std::to_string((long long)v);
    std::ostringstream os;
    os << std::setprecision(15) << v;
    return os.str();
}
static double roundHalfEven(double v) {
    double fl = std::floor(v);
    double diff = v - fl;
    if (diff < 0.5) return fl;
    if (diff > 0.5) return fl + 1.0;
    return (std::fmod(fl, 2.0) == 0.0) ? fl : fl + 1.0;   // .5 -> nearest even
}
// Parse a numeric prefix the way VBScript coercion does; strict=true requires
// the WHOLE string to be numeric (used by IsNumeric).
static bool parseNumber(const std::string& in, double& out, bool strict) {
    std::string s = in;
    size_t a = s.find_first_not_of(" \t");
    if (a == std::string::npos) { if (strict) return false; out = 0; return true; }
    size_t z = s.find_last_not_of(" \t");
    s = s.substr(a, z - a + 1);
    if (s.empty()) { if (strict) return false; out = 0; return true; }
    // hex / octal literals
    if (s.size() > 2 && (s[0] == '&')) {
        char t = (char)std::tolower(s[1]);
        try {
            if (t == 'h') { out = (double)std::stoll(s.substr(2), nullptr, 16); return true; }
            if (t == 'o') { out = (double)std::stoll(s.substr(2), nullptr, 8);  return true; }
        } catch (...) { return false; }
    }
    const char* p = s.c_str();
    char* end = nullptr;
    errno = 0;
    double val = std::strtod(p, &end);
    if (end == p) return false;
    if (strict) { while (*end == ' ' || *end == '\t') ++end; if (*end != '\0') return false; }
    out = val;
    return true;
}

// ----------------------------------------------------------- Value coercions
double Value::toDouble() const {
    switch (type) {
        case Type::Empty:  return 0.0;
        case Type::Null:   raiseErr(94, "Invalid use of Null");
        case Type::Bool:   return b ? -1.0 : 0.0;   // True == -1 in VBScript
        case Type::Int:    return (double)i;
        case Type::Double: return d;
        case Type::String: { double v; if (!parseNumber(s, v, false)) raiseErr(13, "Type mismatch"); return v; }
        case Type::Object: { /* default property */ return 0.0; }
        case Type::Array:  raiseErr(13, "Type mismatch");
    }
    return 0.0;
}
long long Value::toI64() const { return (long long)roundHalfEven(toDouble()); }
int32_t Value::toInt() const {
    long long v = toI64();
    return (int32_t)v;
}
bool Value::toBool() const {
    switch (type) {
        case Type::Empty:  return false;
        case Type::Null:   raiseErr(94, "Invalid use of Null");
        case Type::Bool:   return b;
        case Type::Int:    return i != 0;
        case Type::Double: return d != 0.0;
        case Type::String: {
            std::string t = toLower(s);
            if (t == "true")  return true;
            if (t == "false") return false;
            double v; if (!parseNumber(s, v, false)) raiseErr(13, "Type mismatch");
            return v != 0.0;
        }
        case Type::Object: if (!obj) return false; return true;
        case Type::Array:  raiseErr(13, "Type mismatch");
    }
    return false;
}
std::string Value::toStr() const {
    switch (type) {
        case Type::Empty:  return "";
        case Type::Null:   raiseErr(94, "Invalid use of Null");
        case Type::Bool:   return b ? "True" : "False";
        case Type::Int:    return std::to_string(i);
        case Type::Double: return numToStr(d);
        case Type::String: return s;
        case Type::Object: if (!obj) raiseErr(13, "Type mismatch"); return "[object]";
        case Type::Array:  raiseErr(13, "Type mismatch");
    }
    return "";
}
std::string Value::toConcatStr() const {   // & operator: Null becomes ""
    if (type == Type::Null || type == Type::Empty) return "";
    return toStr();
}
bool Value::looksNumeric() const {
    if (type == Type::Int || type == Type::Double || type == Type::Bool) return true;
    if (type == Type::String) { double v; return parseNumber(s, v, true); }
    return false;
}
std::string Value::vbTypeName() const {
    switch (type) {
        case Type::Empty:  return "Empty";
        case Type::Null:   return "Null";
        case Type::Bool:   return "Boolean";
        case Type::Int:    return "Long";
        case Type::Double: return "Double";
        case Type::String: return "String";
        case Type::Array:  return "Variant()";
        case Type::Object: return obj ? obj->typeName() : "Nothing";
    }
    return "Empty";
}
int Value::varType() const {  // subset of VbVarType constants
    switch (type) {
        case Type::Empty:  return 0;    // vbEmpty
        case Type::Null:   return 1;    // vbNull
        case Type::Int:    return 3;    // vbLong
        case Type::Double: return 5;    // vbDouble
        case Type::String: return 8;    // vbString
        case Type::Bool:   return 11;   // vbBoolean
        case Type::Object: return obj ? 9 : 9; // vbObject
        case Type::Array:  return 8192 + 12;   // vbArray + vbVariant
    }
    return 0;
}

// ================================================================== LEXER
enum class Tok {
    End, Eol, Colon, Ident, Number, String, DateLit,
    Plus, Minus, Star, Slash, Backslash, Caret, Amp,
    Eq, Ne, Lt, Gt, Le, Ge,
    LParen, RParen, Comma, Dot
};
struct Token {
    Tok kind;
    std::string text;   // identifiers keep original case; lower is precomputed
    std::string lower;
    Value num;          // for Number
    std::string str;    // for String
    int line;
    int col = 1;        // 1-based column where the token starts
};

class Lexer {
    const std::string& src;
    size_t pos = 0;
    int line = 1;
    size_t lineStart = 0;   // byte offset of the current line's first char
    int tokCol = 1;         // column of the token currently being read
public:
    explicit Lexer(const std::string& s) : src(s) {}

    std::vector<Token> tokenize() {
        std::vector<Token> out;
        while (true) {
            skipInline();
            if (pos >= src.size()) { push(out, Tok::End); break; }
            char c = src[pos];
            tokCol = (int)(pos - lineStart) + 1;   // column where this token begins

            // comments: ' ... to EOL
            if (c == '\'') { while (pos < src.size() && src[pos] != '\n') pos++; continue; }

            // end of statement
            if (c == '\n') { pos++; lineStart = pos; push(out, Tok::Eol); line++; continue; }
            if (c == '\r') { pos++; continue; }
            if (c == ':')  { pos++; push(out, Tok::Colon); continue; }

            // identifiers / keywords / Rem-comments
            if (std::isalpha((unsigned char)c) || c == '_') {
                std::string id;
                while (pos < src.size() &&
                       (std::isalnum((unsigned char)src[pos]) || src[pos] == '_'))
                    id += src[pos++];
                std::string lo = toLower(id);
                if (lo == "rem") { while (pos < src.size() && src[pos] != '\n') pos++; continue; }
                Token t; t.kind = Tok::Ident; t.text = id; t.lower = lo; t.line = line; t.col = tokCol;
                out.push_back(t);
                continue;
            }

            // numbers (decimal + &H hex + &O octal)
            if (std::isdigit((unsigned char)c) || (c == '.' && pos + 1 < src.size() &&
                                                   std::isdigit((unsigned char)src[pos + 1]))) {
                out.push_back(readNumber());
                continue;
            }
            if (c == '&' && pos + 1 < src.size() &&
                (std::tolower(src[pos + 1]) == 'h' || std::tolower(src[pos + 1]) == 'o')) {
                out.push_back(readNumber());
                continue;
            }

            // strings
            if (c == '"') { out.push_back(readString()); continue; }

            // date literals: #2020-03-15#, #1/15/2020 3:45:00 PM#, #13:45:00#
            if (c == '#') { out.push_back(readDate()); continue; }

            // operators / punctuation
            switch (c) {
                case '+': pos++; push(out, Tok::Plus); break;
                case '-': pos++; push(out, Tok::Minus); break;
                case '*': pos++; push(out, Tok::Star); break;
                case '/': pos++; push(out, Tok::Slash); break;
                case '\\':pos++; push(out, Tok::Backslash); break;
                case '^': pos++; push(out, Tok::Caret); break;
                case '&': pos++; push(out, Tok::Amp); break;
                case '(': pos++; push(out, Tok::LParen); break;
                case ')': pos++; push(out, Tok::RParen); break;
                case ',': pos++; push(out, Tok::Comma); break;
                case '.': pos++; push(out, Tok::Dot); break;
                case '=': pos++; push(out, Tok::Eq); break;
                case '<':
                    pos++;
                    if (peek() == '>') { pos++; push(out, Tok::Ne); }
                    else if (peek() == '=') { pos++; push(out, Tok::Le); }
                    else push(out, Tok::Lt);
                    break;
                case '>':
                    pos++;
                    if (peek() == '=') { pos++; push(out, Tok::Ge); }
                    else push(out, Tok::Gt);
                    break;
                default:
                    raiseErr(1057, std::string("Unexpected character '") + c + "'");
            }
        }
        return out;
    }
private:
    char peek() const { return pos < src.size() ? src[pos] : '\0'; }
    void push(std::vector<Token>& out, Tok k) {
        Token t; t.kind = k; t.line = line; t.col = tokCol; out.push_back(t);
    }
    // Skip spaces/tabs; honor the "_" line-continuation (underscore then EOL).
    void skipInline() {
        while (pos < src.size()) {
            char c = src[pos];
            if (c == ' ' || c == '\t') { pos++; continue; }
            if (c == '_') {
                size_t q = pos + 1;
                while (q < src.size() && (src[q] == ' ' || src[q] == '\t' || src[q] == '\r')) q++;
                if (q < src.size() && src[q] == '\n') { pos = q + 1; lineStart = pos; line++; continue; }
                if (q >= src.size()) { pos = q; continue; }
            }
            break;
        }
    }
    Token readNumber() {
        Token t; t.kind = Tok::Number; t.line = line; t.col = tokCol;
        std::string n;
        if (src[pos] == '&') {
            n += src[pos++];              // &
            n += src[pos++];              // h / o
            while (pos < src.size() && std::isalnum((unsigned char)src[pos])) n += src[pos++];
            double v; parseNumber(n, v, false);
            t.num = Value::integer((int32_t)v);
            return t;
        }
        bool isFloat = false;
        while (pos < src.size()) {
            char c = src[pos];
            if (std::isdigit((unsigned char)c)) n += src[pos++];
            else if (c == '.') { isFloat = true; n += src[pos++]; }
            else if (c == 'e' || c == 'E') {
                isFloat = true; n += src[pos++];
                if (pos < src.size() && (src[pos] == '+' || src[pos] == '-')) n += src[pos++];
            } else break;
        }
        if (isFloat) t.num = Value::number(std::stod(n));
        else {
            try { long long v = std::stoll(n);
                  if (v >= INT32_MIN && v <= INT32_MAX) t.num = Value::integer((int32_t)v);
                  else t.num = Value::number((double)v); }
            catch (...) { t.num = Value::number(std::stod(n)); }
        }
        return t;
    }
    Token readString() {
        Token t; t.kind = Tok::String; t.line = line; t.col = tokCol;
        pos++; // opening quote
        std::string s;
        while (pos < src.size()) {
            char c = src[pos++];
            if (c == '"') {
                if (pos < src.size() && src[pos] == '"') { s += '"'; pos++; }  // "" -> "
                else break;
            } else s += c;
        }
        t.str = s;
        return t;
    }
    Token readDate() {
        Token t; t.kind = Tok::DateLit; t.line = line; t.col = tokCol;
        pos++; // opening #
        std::string s;
        while (pos < src.size() && src[pos] != '#' && src[pos] != '\n') s += src[pos++];
        if (pos < src.size() && src[pos] == '#') pos++;   // closing #
        t.str = s;   // raw contents; parser converts via parseDateToSerial
        return t;
    }
};

// =================================================================== AST
struct Expr;
struct Stmt;
using ExprP = std::shared_ptr<Expr>;
using StmtP = std::shared_ptr<Stmt>;

struct Expr {
    enum K { Lit, Var, Unary, Binary, Index, Member, New } k;
    Value lit;                 // Lit
    std::string name;          // Var name / Member name / New classname / operator
    std::string lname;         // precomputed lowercase of name (Var/Member/New) - avoids per-access toLower
    ExprP a, b;                // Unary:a  Binary:a,b  Index:a(target)  Member:a(target)
    std::vector<ExprP> args;   // Index args
    int line = 0;
};
static ExprP mkLit(Value v)                       { auto e=std::make_shared<Expr>(); e->k=Expr::Lit; e->lit=std::move(v); return e; }
static ExprP mkVar(std::string n)                 { auto e=std::make_shared<Expr>(); e->k=Expr::Var; e->lname=toLower(n); e->name=std::move(n); return e; }
static ExprP mkUnary(std::string op, ExprP a)     { auto e=std::make_shared<Expr>(); e->k=Expr::Unary; e->name=std::move(op); e->a=std::move(a); return e; }
static ExprP mkBinary(std::string op,ExprP a,ExprP b){ auto e=std::make_shared<Expr>(); e->k=Expr::Binary; e->name=std::move(op); e->a=std::move(a); e->b=std::move(b); return e; }
static ExprP mkMember(ExprP t, std::string n)     { auto e=std::make_shared<Expr>(); e->k=Expr::Member; e->a=std::move(t); e->lname=toLower(n); e->name=std::move(n); return e; }
static ExprP mkIndex(ExprP t, std::vector<ExprP> a){ auto e=std::make_shared<Expr>(); e->k=Expr::Index; e->a=std::move(t); e->args=std::move(a); return e; }
static ExprP mkNew(std::string cls)               { auto e=std::make_shared<Expr>(); e->k=Expr::New; e->lname=toLower(cls); e->name=std::move(cls); return e; }

struct Param { std::string name; std::string lname; bool byVal = false; };

// One "Case" label: a plain value, an "Is <op> value", or "lo To hi".
struct CaseTest {
    enum K { Val, Is, Range } k = Val;
    ExprP e1, e2;
    std::string op;   // for Is
};
struct CaseClause {
    std::vector<CaseTest> tests;   // empty => Case Else
    std::vector<StmtP> body;
};

struct Stmt {
    enum K {
        DimS, ReDimS, Assign, SetAssign, IfS, ForS, ForEachS, DoLoopS, WhileS,
        SelectS, SubDecl, FuncDecl, ClassDecl, CallS, ExitS, OnErrorS, ConstS,
        OptionS, PropDecl, WithS, EraseS, MidS
    } k;
    int line = 0;
    int col = 0;

    // Dim / Const
    std::vector<std::pair<std::string, std::vector<ExprP>>> decls; // name + optional dims
    std::vector<std::pair<std::string, ExprP>> consts;

    // ReDim
    bool preserve = false;
    std::string name;                 // ReDim target / For var / class or proc name / Exit-what
    std::string lname;                // precomputed lowercase of name (For/ForEach loop var)
    std::vector<ExprP> dims;

    // Assign / SetAssign
    ExprP target, value;

    // If
    ExprP cond;
    std::vector<StmtP> body;                                  // then / loop body / else-less
    std::vector<std::pair<ExprP, std::vector<StmtP>>> elifs;
    std::vector<StmtP> elseBody;

    // For
    ExprP fromE, toE, stepE;

    // ForEach / With / Select
    ExprP coll;   // For Each collection, With object, Select expr

    // DoLoop
    int testPos = 0;      // 0 none, 1 pre (Do While/Until), 2 post (Loop While/Until)
    bool isUntil = false;

    // Select
    std::vector<CaseClause> cases;

    // Sub/Function/Property
    std::vector<Param> params;
    bool isFunction = false;
    int propKind = 0;     // 0 Get, 1 Let, 2 Set
    bool isDefault = false;   // marked with the Default keyword (class default member)

    // Class
    std::vector<StmtP> members;

    // Call
    ExprP callee;
    std::vector<ExprP> callArgs;

    // OnError: 1 = Resume Next, 0 = Goto 0
    int errMode = 0;
    // Option Explicit
    bool optExplicit = false;
};

// ================================================================== PARSER
// parseDateToSerial is defined in the interpreter section; the parser needs it for #date# literals.
static bool parseDateToSerial(const std::string& in, double& out);

class Parser {
    std::vector<Token> t;
    size_t p = 0;
    bool inProc = false;    // currently inside a Sub/Function/Property body
    bool inClass = false;   // currently inside a Class body
    std::string baseDir_;                                  // directory of the file being parsed (for relative Include)
    std::shared_ptr<std::set<std::string>> included_;      // canonical paths already included (guard against dupes/cycles)
public:
    explicit Parser(std::vector<Token> toks, std::string baseDir = std::string(),
                    std::shared_ptr<std::set<std::string>> included = nullptr)
        : t(std::move(toks)), baseDir_(std::move(baseDir)),
          included_(included ? included : std::make_shared<std::set<std::string>>()) {}
    ExprP parseOneExpr() { return parseExpr(); }   // used by Eval()
    std::vector<StmtP> parseProgram() {
        auto prog = parseStatements({});
        if (cur().kind != Tok::End)
            err("Unexpected '" + cur().text + "'");
        return prog;
    }
private:
    // ---- token helpers ----
    const Token& cur() const { return t[p]; }
    const Token& peek(size_t n = 1) const { return t[std::min(p + n, t.size() - 1)]; }
    void adv() { if (p + 1 < t.size()) p++; }
    bool is(Tok k) const { return cur().kind == k; }
    bool isKw(const char* kw) const { return cur().kind == Tok::Ident && cur().lower == kw; }
    bool peekKw(const char* kw) const { return peek().kind == Tok::Ident && peek().lower == kw; }
    [[noreturn]] void err(const std::string& m) {
        VbError e(1025, m, "Directive compilation error");
        e.line = cur().line;
        e.col  = cur().col;
        throw e;
    }
    void expect(Tok k, const char* what) { if (!is(k)) err(std::string("Expected ") + what); adv(); }
    void expectKw(const char* kw) { if (!isKw(kw)) err(std::string("Expected '") + kw + "'"); adv(); }
    void skipSeps() { while (is(Tok::Eol) || is(Tok::Colon)) adv(); }
    void skipEol()  { while (is(Tok::Eol)) adv(); }
    std::string ident(const char* what) {
        if (!is(Tok::Ident)) err(std::string("Expected ") + what);
        std::string s = cur().text; adv(); return s;
    }
    static bool canStartExpr(const Token& tk) {
        switch (tk.kind) {
            case Tok::Number: case Tok::String: case Tok::DateLit: case Tok::LParen:
            case Tok::Minus: case Tok::Plus: case Tok::Ident: case Tok::Dot:
                return true;
            default: return false;
        }
    }

    // ---- block of statements until a terminator keyword ----
    std::vector<StmtP> parseStatements(const std::set<std::string>& stop) {
        std::vector<StmtP> out;
        while (true) {
            skipSeps();
            if (is(Tok::End)) break;
            if (is(Tok::Ident) && stop.count(cur().lower)) break;
            if (is(Tok::Ident) && cur().lower == "include") { parseIncludeInto(out); continue; }
            out.push_back(parseStatement());
        }
        return out;
    }
    // Include "file" — parse the referenced file at parse time and splice its
    // top-level statements in here (so its Sub/Function/Class/Const hoist normally).
    void parseIncludeInto(std::vector<StmtP>& out) {
        if (inProc || inClass) err("Include is only allowed at the top level");
        adv();   // 'include'
        if (!is(Tok::String)) err("Include expects a quoted file path");
        std::string rel = cur().str; adv();

        fs::path pth(rel);
        std::error_code ec;
        if (pth.is_relative() && !baseDir_.empty()) pth = fs::path(baseDir_) / pth;
        fs::path canon = fs::weakly_canonical(pth, ec);
        if (ec || canon.empty()) canon = fs::absolute(pth, ec);
        std::string key = canon.string();

        if (included_->count(key)) return;         // include guard: skip already-included files
        included_->insert(key);

        std::ifstream f(canon, std::ios::binary);
        if (!f) err("Include: cannot open file '" + rel + "'");
        std::string src((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
        if (src.size() >= 3 && (unsigned char)src[0] == 0xEF &&
            (unsigned char)src[1] == 0xBB && (unsigned char)src[2] == 0xBF)
            src = src.substr(3);                   // strip UTF-8 BOM

        Lexer lx(src);
        Parser sub(lx.tokenize(), canon.parent_path().string(), included_);
        std::vector<StmtP> inc;
        try { inc = sub.parseProgram(); }
        catch (VbError& e) { e.description = "in included file '" + rel + "': " + e.description; throw; }
        for (auto& s : inc) out.push_back(std::move(s));
    }
    // ---- statements separated by ':' on a single line (for one-line If) ----
    std::vector<StmtP> parseInlineStmts(const std::set<std::string>& stop) {
        std::vector<StmtP> out;
        while (true) {
            if (is(Tok::Eol) || is(Tok::End)) break;
            if (is(Tok::Ident) && stop.count(cur().lower)) break;
            out.push_back(parseStatement());
            if (is(Tok::Colon)) { adv(); continue; }
            break;
        }
        return out;
    }

    // ---- dispatch a single statement ----
    StmtP parseStatement() {
        int ln = cur().line, col = cur().col;
        StmtP s = parseStatementInner();
        if (s) { if (!s->line) s->line = ln; if (!s->col) s->col = col; }
        return s;
    }
    StmtP parseStatementInner() {
        int ln = cur().line;
        if (is(Tok::Dot)) return parseAssignOrCall();   // .Member inside With
        if (is(Tok::Ident)) {
            const std::string& kw = cur().lower;
            if (kw == "dim")      return parseDim(false);
            if (kw == "redim")    return parseReDim();
            if (kw == "set")      return parseSet();
            if (kw == "if")       return parseIf();
            if (kw == "for")      return parseFor();
            if (kw == "do")       return parseDo();
            if (kw == "while")    return parseWhile();
            if (kw == "select")   return parseSelect();
            if (kw == "sub")      { if (inProc) err("a Sub cannot be defined inside another procedure"); return parseProc(false); }
            if (kw == "function") { if (inProc) err("a Function cannot be defined inside another procedure"); return parseProc(true); }
            if (kw == "property") { if (inProc) err("a Property cannot be defined inside another procedure"); if (!inClass) err("Property Get/Let/Set is only valid inside a Class"); return parseProperty(); }
            if (kw == "class")    { if (inProc || inClass) err("a Class cannot be defined inside a procedure or another Class"); return parseClass(); }
            if (kw == "with")     return parseWith();
            if (kw == "call")     { adv(); auto e = parsePostfix();
                                    auto s = mk(Stmt::CallS, ln); s->callee = e; return s; }
            if (kw == "exit")     { adv(); auto s = mk(Stmt::ExitS, ln); s->name = toLower(ident("Sub/Function/For/Do/Property")); return s; }
            if (kw == "const")    return parseConst();
            if (kw == "option")   { adv(); expectKw("explicit"); auto s = mk(Stmt::OptionS, ln); s->optExplicit = true; return s; }
            if (kw == "on")       return parseOnError();
            if (kw == "erase")    { adv(); auto s = mk(Stmt::EraseS, ln);
                                    while (true) { s->callArgs.push_back(parsePostfix()); if (is(Tok::Comma)) { adv(); continue; } break; }
                                    return s; }
            if (kw == "public" || kw == "private") {
                adv();
                bool isDef = false;
                if (isKw("default")) { adv(); isDef = true; }   // Public Default Property/Function/Sub
                if (isKw("sub"))      { if (inProc) err("a Sub cannot be defined inside another procedure"); return parseProc(false, isDef); }
                if (isKw("function")) { if (inProc) err("a Function cannot be defined inside another procedure"); return parseProc(true, isDef); }
                if (isKw("property")) { if (inProc) err("a Property cannot be defined inside another procedure"); if (!inClass) err("Property Get/Let/Set is only valid inside a Class"); return parseProperty(isDef); }
                if (isKw("class"))    { if (inProc || inClass) err("a Class cannot be defined inside a procedure or another Class"); return parseClass(); }
                if (isKw("const"))    return parseConst();
                return parseDim(false);   // Public/Private x, y  (treated like Dim)
            }
            if (kw == "default") {   // Default [on its own] implies Public
                adv();
                if (isKw("property")) { if (inProc) err("a Property cannot be defined inside another procedure"); if (!inClass) err("Property Get/Let/Set is only valid inside a Class"); return parseProperty(true); }
                if (isKw("function")) { if (inProc) err("a Function cannot be defined inside another procedure"); return parseProc(true, true); }
                if (isKw("sub"))      { if (inProc) err("a Sub cannot be defined inside another procedure"); return parseProc(false, true); }
                err("Default must be followed by Property, Function, or Sub");
            }
        }
        return parseAssignOrCall();
    }

    StmtP mk(Stmt::K k, int ln) { auto s = std::make_shared<Stmt>(); s->k = k; s->line = ln; return s; }

    // ---- Dim ----
    StmtP parseDim(bool /*unused*/) {
        int ln = cur().line; if (isKw("dim")) adv();   // no keyword when reached via Public/Private
        auto s = mk(Stmt::DimS, ln);
        while (true) {
            std::string nm = ident("variable name");
            std::vector<ExprP> dims;
            if (is(Tok::LParen)) { adv(); dims = parseArgList(Tok::RParen); }
            s->decls.push_back({nm, dims});
            if (is(Tok::Comma)) { adv(); continue; }
            break;
        }
        return s;
    }
    StmtP parseReDim() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::ReDimS, ln);
        if (isKw("preserve")) { adv(); s->preserve = true; }
        s->name = ident("array name");
        expect(Tok::LParen, "(");
        s->dims = parseArgList(Tok::RParen);
        return s;
    }
    StmtP parseSet() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::SetAssign, ln);
        s->target = parsePostfix();
        expect(Tok::Eq, "=");
        s->value = parseExpr();
        return s;
    }
    StmtP parseConst() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::ConstS, ln);
        while (true) {
            std::string nm = ident("const name");
            expect(Tok::Eq, "=");
            s->consts.push_back({nm, parseExpr()});
            if (is(Tok::Comma)) { adv(); continue; }
            break;
        }
        return s;
    }
    StmtP parseOnError() {
        int ln = cur().line; adv();          // On
        expectKw("error");
        auto s = mk(Stmt::OnErrorS, ln);
        if (isKw("resume")) { adv(); expectKw("next"); s->errMode = 1; }
        else { expectKw("goto"); if (is(Tok::Number)) adv(); s->errMode = 0; }  // GoTo 0
        return s;
    }
    StmtP parseIf() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::IfS, ln);
        s->cond = parseExpr();
        expectKw("then");
        if (is(Tok::Eol)) {                          // block form
            skipEol();
            s->body = parseStatements({"elseif", "else", "end"});
            while (isKw("elseif")) {
                adv(); ExprP c = parseExpr(); expectKw("then"); skipEol();
                s->elifs.push_back({c, parseStatements({"elseif", "else", "end"})});
            }
            if (isKw("else")) { adv(); skipEol(); s->elseBody = parseStatements({"end"}); }
            expectKw("end"); expectKw("if");
        } else {                                     // single-line form
            s->body = parseInlineStmts({"else"});
            if (isKw("else")) { adv(); s->elseBody = parseInlineStmts({}); }
        }
        return s;
    }
    StmtP parseFor() {
        int ln = cur().line; adv();
        if (isKw("each")) {
            adv();
            auto s = mk(Stmt::ForEachS, ln);
            s->name = ident("loop variable");
            s->lname = toLower(s->name);
            expectKw("in");
            s->coll = parseExpr();
            skipEol();
            s->body = parseStatements({"next"});
            expectKw("next");
            return s;
        }
        auto s = mk(Stmt::ForS, ln);
        s->name = ident("loop variable");
        s->lname = toLower(s->name);
        expect(Tok::Eq, "=");
        s->fromE = parseExpr();
        expectKw("to");
        s->toE = parseExpr();
        if (isKw("step")) { adv(); s->stepE = parseExpr(); }
        skipEol();
        s->body = parseStatements({"next"});
        expectKw("next");
        if (is(Tok::Ident)) adv();    // optional loop var after Next
        return s;
    }
    StmtP parseDo() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::DoLoopS, ln);
        if (isKw("while")) { adv(); s->testPos = 1; s->isUntil = false; s->cond = parseExpr(); }
        else if (isKw("until")) { adv(); s->testPos = 1; s->isUntil = true; s->cond = parseExpr(); }
        skipEol();
        s->body = parseStatements({"loop"});
        expectKw("loop");
        if (s->testPos == 0) {
            if (isKw("while")) { adv(); s->testPos = 2; s->isUntil = false; s->cond = parseExpr(); }
            else if (isKw("until")) { adv(); s->testPos = 2; s->isUntil = true; s->cond = parseExpr(); }
        }
        return s;
    }
    StmtP parseWhile() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::WhileS, ln);
        s->cond = parseExpr();
        skipEol();
        s->body = parseStatements({"wend"});
        expectKw("wend");
        return s;
    }
    StmtP parseSelect() {
        int ln = cur().line; adv();
        expectKw("case");
        auto s = mk(Stmt::SelectS, ln);
        s->coll = parseExpr();
        skipSeps();
        while (isKw("case")) {
            adv();
            CaseClause cc;
            if (isKw("else")) { adv(); }                 // Case Else
            else {
                while (true) {
                    CaseTest ct;
                    if (isKw("is")) {
                        adv(); ct.k = CaseTest::Is;
                        ct.op = readRelOp();
                        ct.e1 = parseConcat();
                    } else {
                        ct.e1 = parseExpr();
                        if (isKw("to")) { adv(); ct.k = CaseTest::Range; ct.e2 = parseExpr(); }
                        else ct.k = CaseTest::Val;
                    }
                    cc.tests.push_back(ct);
                    if (is(Tok::Comma)) { adv(); continue; }
                    break;
                }
            }
            skipSeps();
            cc.body = parseStatements({"case", "end"});
            s->cases.push_back(std::move(cc));
        }
        expectKw("end"); expectKw("select");
        return s;
    }
    std::string readRelOp() {
        switch (cur().kind) {
            case Tok::Eq: adv(); return "=";
            case Tok::Ne: adv(); return "<>";
            case Tok::Lt: adv(); return "<";
            case Tok::Gt: adv(); return ">";
            case Tok::Le: adv(); return "<=";
            case Tok::Ge: adv(); return ">=";
            default: err("Expected comparison operator"); return "";
        }
    }
    std::vector<Param> parseParams() {
        std::vector<Param> ps;
        expect(Tok::LParen, "(");
        if (!is(Tok::RParen)) {
            while (true) {
                Param pm;
                if (isKw("byval")) { adv(); pm.byVal = true; }
                else if (isKw("byref")) { adv(); pm.byVal = false; }
                pm.name = ident("parameter name");
                pm.lname = toLower(pm.name);
                if (is(Tok::LParen)) { adv(); expect(Tok::RParen, ")"); }  // arr() param
                ps.push_back(pm);
                if (is(Tok::Comma)) { adv(); continue; }
                break;
            }
        }
        expect(Tok::RParen, ")");
        return ps;
    }
    StmtP parseProc(bool isFunc, bool isDefault = false) {
        int ln = cur().line; adv();
        auto s = mk(isFunc ? Stmt::FuncDecl : Stmt::SubDecl, ln);
        s->isFunction = isFunc;
        s->isDefault = isDefault;
        s->name = ident("procedure name");
        s->lname = toLower(s->name);
        if (is(Tok::LParen)) s->params = parseParams();
        skipEol();
        bool savedProc = inProc; inProc = true;
        s->body = parseStatements({"end"});
        inProc = savedProc;
        expectKw("end");
        expectKw(isFunc ? "function" : "sub");
        return s;
    }
    StmtP parseProperty(bool isDefault = false) {
        int ln = cur().line; adv();
        auto s = mk(Stmt::PropDecl, ln);
        s->isDefault = isDefault;
        if (isKw("get")) { s->propKind = 0; adv(); }
        else if (isKw("let")) { s->propKind = 1; adv(); }
        else if (isKw("set")) { s->propKind = 2; adv(); }
        else err("Expected Get/Let/Set");
        s->name = ident("property name");
        s->lname = toLower(s->name);
        if (is(Tok::LParen)) s->params = parseParams();
        skipEol();
        bool savedProc = inProc; inProc = true;
        s->body = parseStatements({"end"});
        inProc = savedProc;
        expectKw("end"); expectKw("property");
        return s;
    }
    StmtP parseClass() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::ClassDecl, ln);
        s->name = ident("class name");
        skipSeps();
        bool savedClass = inClass; inClass = true;
        while (!isKw("end") && !is(Tok::End)) {
            s->members.push_back(parseStatement());
            skipSeps();
        }
        inClass = savedClass;
        expectKw("end"); expectKw("class");
        return s;
    }
    StmtP parseWith() {
        int ln = cur().line; adv();
        auto s = mk(Stmt::WithS, ln);
        s->coll = parseExpr();
        skipEol();
        s->body = parseStatements({"end"});
        expectKw("end"); expectKw("with");
        return s;
    }

    // ---- assignment or call (statement starting with lvalue/callee) ----
    StmtP parseAssignOrCall() {
        int ln = cur().line;
        ExprP lhs = parsePostfix();
        if (is(Tok::Eq)) {
            adv();
            // Mid(strVar, start[, length]) = value  — VBScript's in-place Mid statement.
            if (lhs->k == Expr::Index && lhs->a && lhs->a->k == Expr::Var &&
                lhs->a->lname == "mid" && lhs->args.size() >= 2 && lhs->args.size() <= 3) {
                auto s = mk(Stmt::MidS, ln);
                s->target = lhs->args[0];                    // the string variable (an lvalue)
                s->callArgs.push_back(lhs->args[1]);         // start (1-based)
                if (lhs->args.size() == 3) s->callArgs.push_back(lhs->args[2]);  // optional length
                s->value = parseExpr();                      // replacement text
                return s;
            }
            auto s = mk(Stmt::Assign, ln);
            s->target = lhs; s->value = parseExpr();
            return s;
        }
        // parenless sub call:  Foo a, b   /   obj.M a, b
        if ((lhs->k == Expr::Var || lhs->k == Expr::Member) && canStartExpr(cur())) {
            auto s = mk(Stmt::CallS, ln);
            s->callee = lhs;
            while (true) {
                if (is(Tok::Comma)) { s->callArgs.push_back(mkLit(Value::empty())); adv(); continue; }
                s->callArgs.push_back(parseExpr());
                if (is(Tok::Comma)) { adv(); continue; }
                break;
            }
            return s;
        }
        auto s = mk(Stmt::CallS, ln);   // bare call:  Foo   /   obj.M   /   Foo(1,2)
        s->callee = lhs;
        return s;
    }

    // ---- expressions (precedence climbing) ----
    ExprP parseExpr() { return parseImp(); }
    ExprP parseImp() { auto l = parseEqv(); while (isKw("imp")) { adv(); l = mkBinary("Imp", l, parseEqv()); } return l; }
    ExprP parseEqv() { auto l = parseXor(); while (isKw("eqv")) { adv(); l = mkBinary("Eqv", l, parseXor()); } return l; }
    ExprP parseXor() { auto l = parseOr();  while (isKw("xor")) { adv(); l = mkBinary("Xor", l, parseOr()); } return l; }
    ExprP parseOr()  { auto l = parseAnd(); while (isKw("or"))  { adv(); l = mkBinary("Or", l, parseAnd()); } return l; }
    ExprP parseAnd() { auto l = parseNot(); while (isKw("and")) { adv(); l = mkBinary("And", l, parseNot()); } return l; }
    ExprP parseNot() { if (isKw("not")) { adv(); return mkUnary("Not", parseNot()); } return parseCmp(); }
    ExprP parseCmp() {
        auto l = parseConcat();
        while (true) {
            std::string op;
            switch (cur().kind) {
                case Tok::Eq: op = "="; break;   case Tok::Ne: op = "<>"; break;
                case Tok::Lt: op = "<"; break;   case Tok::Gt: op = ">"; break;
                case Tok::Le: op = "<="; break;  case Tok::Ge: op = ">="; break;
                default: if (isKw("is")) op = "Is"; break;
            }
            if (op.empty()) break;
            adv();
            l = mkBinary(op, l, parseConcat());
        }
        return l;
    }
    ExprP parseConcat() { auto l = parseAdd(); while (is(Tok::Amp)) { adv(); l = mkBinary("&", l, parseAdd()); } return l; }
    ExprP parseAdd() {
        auto l = parseModL();
        while (is(Tok::Plus) || is(Tok::Minus)) {
            std::string op = is(Tok::Plus) ? "+" : "-"; adv();
            l = mkBinary(op, l, parseModL());
        }
        return l;
    }
    ExprP parseModL() { auto l = parseIntDiv(); while (isKw("mod")) { adv(); l = mkBinary("Mod", l, parseIntDiv()); } return l; }
    ExprP parseIntDiv() { auto l = parseMul(); while (is(Tok::Backslash)) { adv(); l = mkBinary("\\", l, parseMul()); } return l; }
    ExprP parseMul() {
        auto l = parseNeg();
        while (is(Tok::Star) || is(Tok::Slash)) {
            std::string op = is(Tok::Star) ? "*" : "/"; adv();
            l = mkBinary(op, l, parseNeg());
        }
        return l;
    }
    ExprP parseNeg() {
        if (is(Tok::Minus)) { adv(); return mkUnary("-", parseNeg()); }
        if (is(Tok::Plus))  { adv(); return parseNeg(); }
        return parsePow();
    }
    ExprP parsePow() {
        auto base = parsePostfix();
        if (is(Tok::Caret)) { adv(); return mkBinary("^", base, parseNeg()); }
        return base;
    }
    ExprP parsePostfix() {
        ExprP e = parsePrimary();
        while (true) {
            if (is(Tok::Dot)) { adv(); e = mkMember(e, ident("member name")); }
            else if (is(Tok::LParen)) { adv(); e = mkIndex(e, parseArgList(Tok::RParen)); }
            else break;
        }
        return e;
    }
    ExprP parsePrimary() {
        if (is(Tok::Number)) { Value v = cur().num; adv(); return mkLit(v); }
        if (is(Tok::String)) { std::string s = cur().str; adv(); return mkLit(Value::str(s)); }
        if (is(Tok::DateLit)) { double ser; if (!parseDateToSerial(cur().str, ser)) err("Invalid date literal '#" + cur().str + "#'"); adv(); return mkLit(Value::number(ser)); }
        if (is(Tok::Dot))    { adv(); return mkMember(nullptr, ident("member name")); }  // With-member
        if (is(Tok::LParen)) { adv(); auto e = parseExpr(); expect(Tok::RParen, ")"); return e; }
        if (is(Tok::Ident)) {
            const std::string& lo = cur().lower;
            if (lo == "true")    { adv(); return mkLit(Value::boolean(true)); }
            if (lo == "false")   { adv(); return mkLit(Value::boolean(false)); }
            if (lo == "nothing") { adv(); return mkLit(Value::nothing()); }
            if (lo == "empty")   { adv(); return mkLit(Value::empty()); }
            if (lo == "null")    { adv(); return mkLit(Value::null()); }
            if (lo == "new")     { adv(); return mkNew(ident("class name")); }
            std::string nm = cur().text; adv(); return mkVar(nm);
        }
        err("Expected expression");
        return nullptr;
    }
    std::vector<ExprP> parseArgList(Tok close) {
        std::vector<ExprP> args;
        if (is(close)) { adv(); return args; }
        while (true) {
            if (is(Tok::Comma)) { args.push_back(mkLit(Value::empty())); adv(); continue; }
            if (is(close)) break;
            args.push_back(parseExpr());
            if (is(Tok::Comma)) { adv(); continue; }
            break;
        }
        expect(close, ")");
        return args;
    }
};

// ============================================================ NATIVE OBJECTS
struct QuitSignal { int code; };   // thrown by WScript.Quit

static Value makeArray(std::vector<Value> items) {
    auto ad = std::make_shared<ArrayData>();
    ad->uppers = { (int)items.size() - 1 };
    ad->data = std::move(items);
    return Value::array(ad);
}

// ---- Err ----
struct ErrObject : IObject {
    long number = 0;
    std::string source, description;
    std::string typeName() const override { return "ErrObject"; }
    void clear() { number = 0; source.clear(); description.clear(); }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "number")      return Value::integer((int32_t)number);
        if (n == "description") return Value::str(description);
        if (n == "source")      return Value::str(source);
        if (n == "clear")       { clear(); return Value::empty(); }
        if (n == "raise") {
            long num = args.size() > 0 ? args[0].toI64() : 0;
            std::string src = args.size() > 1 && !args[1].isEmpty() ? args[1].toStr() : "";
            std::string des = args.size() > 2 && !args[2].isEmpty() ? args[2].toStr()
                                                                    : ("Error " + std::to_string(num));
            throw VbError(num, des, src.empty() ? "Directive runtime error" : src);
        }
        raiseErr(438, "Err: unknown member " + name);
    }
    void set(Interpreter&, const std::string& name, std::vector<Value>&, const Value& v, bool) override {
        std::string n = toLower(name);
        if (n == "number")           number = v.toI64();
        else if (n == "description") description = v.toStr();
        else if (n == "source")      source = v.toStr();
        else raiseErr(438, "Err: cannot set " + name);
    }
};

// ---- Directive.StdOut / .StdErr / .StdIn ----
// Output streams always use std::cout/std::cerr; stdin uses std::cin. In a GUI
// build these are wired to the parent console (if any) via AttachConsole in WinMain.
struct ConsoleStream : IObject {
    int which;   // 0 = stdout, 1 = stderr, 2 = stdin
    explicit ConsoleStream(int w) : which(w) {}
    std::string typeName() const override { return "TextStream"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (which != 2) {
            std::ostream& os = (which == 1) ? std::cerr : std::cout;
            if (n == "write")           { os << (args.empty() ? "" : args[0].toConcatStr()); return Value::empty(); }
            if (n == "writeline")        { os << (args.empty() ? "" : args[0].toConcatStr()) << "\n"; return Value::empty(); }
            if (n == "writeblanklines")  { long k = args.empty() ? 1 : args[0].toI64(); for (long i = 0; i < k; ++i) os << "\n"; return Value::empty(); }
            if (n == "close" || n == "flush") { os.flush(); return Value::empty(); }
        } else {
            if (n == "readline")     { std::string l; if (!std::getline(std::cin, l)) raiseErr(62, "Input past end of file"); return Value::str(l); }
            if (n == "read")         { long k = args.empty() ? 1 : args[0].toI64(); if (k < 0) k = 0; std::string buf((size_t)k, '\0'); std::cin.read(&buf[0], k); buf.resize((size_t)std::cin.gcount()); return Value::str(buf); }
            if (n == "readall")      { std::ostringstream ss; ss << std::cin.rdbuf(); return Value::str(ss.str()); }
            if (n == "atendofstream")return Value::boolean(std::cin.peek() == EOF);
        }
        raiseErr(438, "TextStream: unknown member " + name);
    }
};

// ---- Directive (host object; replaces WScript) ----
// WScript.Arguments-style collection of the command-line arguments passed to the
// script (0-based, like WSH). Supports Count/Length, Item(i), the default index
// Arguments(i), and For Each iteration.
struct ArgumentsObject : IObject {
    std::vector<Value> items;
    ArgumentsObject() = default;
    explicit ArgumentsObject(const std::vector<std::string>& a) { for (auto& s : a) items.push_back(Value::str(s)); }
    std::string typeName() const override { return "Arguments"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "count" || n == "length") return Value::integer((int32_t)items.size());
        if (n.empty() || n == "item") {
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            long i = (long)args[0].toI64();
            if (i < 0 || i >= (long)items.size()) raiseErr(9, "Subscript out of range");
            return items[(size_t)i];
        }
        raiseErr(438, "Arguments: unknown member " + name);
    }
    bool enumerate(Interpreter&, std::vector<Value>& out) override { out = items; return true; }
};

struct DirectiveObject : IObject {
    std::string scriptName = "script.directive";
    std::string scriptFullName;                 // absolute path of the running script ("" for -e/stdin)
    std::string scriptDir;                       // directory containing the script (AutoIt @ScriptDir)
    std::vector<std::string> scriptArgs;         // command-line arguments after the script path
    std::string typeName() const override { return "Directive"; }
    Value get(Interpreter& in, const std::string& name, std::vector<Value>& args) override {
        (void)in;
        std::string n = toLower(name);
        if (n == "arguments" || n == "args") {
            auto ao = std::make_shared<ArgumentsObject>(scriptArgs);
            if (!args.empty()) return ao->get(in, "", args);   // Directive.Arguments(i) -> item i
            return Value::object(ao);
        }
        if (n == "echo") {
            std::string out;
            for (size_t k = 0; k < args.size(); ++k) { if (k) out += " "; out += args[k].toConcatStr(); }
            ui::echo(out);
            return Value::empty();
        }
        if (n == "stdout")     return Value::object(std::make_shared<ConsoleStream>(0));
        if (n == "stderr")     return Value::object(std::make_shared<ConsoleStream>(1));
        if (n == "stdin")      return Value::object(std::make_shared<ConsoleStream>(2));
        if (n == "sleep")      return Value::empty();
        if (n == "quit")       throw QuitSignal{ args.empty() ? 0 : (int)args[0].toI64() };
        if (n == "scriptname")     return Value::str(scriptName);
        if (n == "scriptfullname") return Value::str(scriptFullName);
        if (n == "scriptpath" || n == "scriptdir") return Value::str(scriptDir);
        raiseErr(438, "Directive: unknown member " + name);
    }
};

// ---- List (native doubly-linked list) ----
struct ListObject : IObject {
    std::list<Value> data;
    std::string typeName() const override { return "List"; }

    static bool eq(const Value& a, const Value& b) {
        if (a.isObject() || b.isObject())
            return a.isObject() && b.isObject() && a.obj.get() == b.obj.get();
        bool an = a.looksNumeric(), bn = b.looksNumeric();
        if (a.type == Value::Type::String || b.type == Value::Type::String)
            if (!(an && bn)) return a.toStr() == b.toStr();
        return a.toDouble() == b.toDouble();
    }
    std::list<Value>::iterator at(long i) {
        if (i < 0 || i >= (long)data.size()) raiseErr(9, "Subscript out of range");
        auto it = data.begin(); std::advance(it, i); return it;
    }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n.empty() || n == "item") {                       // default: Item(index)
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            return *at(args[0].toI64());
        }
        if (n == "add" || n == "append" || n == "push") { for (auto& a : args) data.push_back(a); return Value::empty(); }
        if (n == "insert") {                                  // Insert index, value
            if (args.size() < 2) raiseErr(450, "Wrong number of arguments");
            long i = args[0].toI64();
            if (i < 0 || i > (long)data.size()) raiseErr(9, "Subscript out of range");
            auto it = data.begin(); std::advance(it, i); data.insert(it, args[1]);
            return Value::empty();
        }
        if (n == "removeat") { if (args.empty()) raiseErr(450, "Wrong number of arguments"); data.erase(at(args[0].toI64())); return Value::empty(); }
        if (n == "remove") {                                  // Remove value -> bool found
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            for (auto it = data.begin(); it != data.end(); ++it)
                if (eq(*it, args[0])) { data.erase(it); return Value::boolean(true); }
            return Value::boolean(false);
        }
        if (n == "clear")    { data.clear(); return Value::empty(); }
        if (n == "count")    return Value::integer((int32_t)data.size());
        if (n == "contains") { for (auto& v : data) if (!args.empty() && eq(v, args[0])) return Value::boolean(true); return Value::boolean(false); }
        if (n == "indexof")  { long i = 0; for (auto& v : data) { if (!args.empty() && eq(v, args[0])) return Value::integer((int32_t)i); ++i; } return Value::integer(-1); }
        if (n == "swap") {                                    // Swap i, j
            if (args.size() < 2) raiseErr(450, "Wrong number of arguments");
            std::iter_swap(at(args[0].toI64()), at(args[1].toI64())); return Value::empty();
        }
        if (n == "reverse")  { data.reverse(); return Value::empty(); }
        if (n == "sort") {
            std::vector<Value> v(data.begin(), data.end());
            std::sort(v.begin(), v.end(), [](const Value& a, const Value& b) {
                bool an = a.looksNumeric(), bn = b.looksNumeric();
                if (an && bn) return a.toDouble() < b.toDouble();
                return a.toStr() < b.toStr();
            });
            data.assign(v.begin(), v.end()); return Value::empty();
        }
        if (n == "first")   { if (data.empty()) raiseErr(9, "List is empty"); return data.front(); }
        if (n == "last")    { if (data.empty()) raiseErr(9, "List is empty"); return data.back(); }
        if (n == "toarray") { return makeArray(std::vector<Value>(data.begin(), data.end())); }
        raiseErr(438, "List: unknown member " + name);
    }
    void set(Interpreter&, const std::string& name, std::vector<Value>& args, const Value& v, bool) override {
        std::string n = toLower(name);
        if (n.empty() || n == "item") { if (args.empty()) raiseErr(450, "Wrong number of arguments"); *at(args[0].toI64()) = v; return; }
        raiseErr(438, "List: cannot set " + name);
    }
    bool enumerate(Interpreter&, std::vector<Value>& out) override { for (auto& v : data) out.push_back(v); return true; }
};

// ---- Scripting.Dictionary ----
// VBScript Dictionary keys are typed: numeric 1 and string "1" are different keys,
// and object keys compare by identity. Encode type + value so they don't collide.
static std::string dictKey(const Value& v) {
    switch (v.type) {
        case Value::Type::Empty:  return "\x01""e";
        case Value::Type::Null:   return "\x01""z";
        case Value::Type::Bool:   return std::string("\x01""b") + (v.b ? "1" : "0");
        case Value::Type::Int:
        case Value::Type::Double: return "\x01""n" + v.toStr();          // 1 and 1.0 → same numeric key
        case Value::Type::String: return "\x01""s" + v.s;
        case Value::Type::Object: { char b[24]; std::snprintf(b, sizeof(b), "\x01o%p", (void*)v.obj.get()); return b; }
        default:                  return "\x01?" + v.toStr();
    }
}

struct DictionaryObject : IObject {
    std::vector<std::pair<Value, Value>> entries;      // insertion order
    std::unordered_map<std::string, size_t> index;
    std::string typeName() const override { return "Dictionary"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n.empty() || n == "item") {                // default property
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            auto it = index.find(dictKey(args[0]));
            return it == index.end() ? Value::empty() : entries[it->second].second;
        }
        if (n == "add") {
            if (args.size() < 2) raiseErr(450, "Wrong number of arguments");
            std::string k = dictKey(args[0]);
            if (index.count(k)) raiseErr(457, "This key is already associated with an element of this collection");
            index[k] = entries.size();
            entries.push_back({ args[0], args[1] });
            return Value::empty();
        }
        if (n == "exists") return Value::boolean(!args.empty() && index.count(dictKey(args[0])) > 0);
        if (n == "keys")   { std::vector<Value> v; for (auto& e : entries) v.push_back(e.first);  return makeArray(v); }
        if (n == "items")  { std::vector<Value> v; for (auto& e : entries) v.push_back(e.second); return makeArray(v); }
        if (n == "count")  return Value::integer((int32_t)entries.size());
        if (n == "remove") {
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            std::string k = dictKey(args[0]);
            auto it = index.find(k);
            if (it == index.end()) raiseErr(32811, "Element not found");
            entries.erase(entries.begin() + it->second);
            reindex();
            return Value::empty();
        }
        if (n == "removeall") { entries.clear(); index.clear(); return Value::empty(); }
        if (n == "comparemode") return Value::integer(0);
        raiseErr(438, "Dictionary: unknown member " + name);
    }
    void set(Interpreter&, const std::string& name, std::vector<Value>& args, const Value& v, bool) override {
        std::string n = toLower(name);
        if (n.empty() || n == "item") {                // d(key) = v  or  d.Item(key) = v
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            std::string k = dictKey(args[0]);
            auto it = index.find(k);
            if (it == index.end()) { index[k] = entries.size(); entries.push_back({ args[0], v }); }
            else entries[it->second].second = v;
            return;
        }
        if (n == "comparemode") return;                // ignored
        raiseErr(438, "Dictionary: cannot set " + name);
    }
    bool enumerate(Interpreter&, std::vector<Value>& out) override {
        for (auto& e : entries) out.push_back(e.first);   // For Each iterates keys
        return true;
    }
private:
    void reindex() { index.clear(); for (size_t k = 0; k < entries.size(); ++k) index[dictKey(entries[k].first)] = k; }
};

// ---- RegExp (VBScript.RegExp) with a compact Match/Matches model ----
struct MatchObject : IObject {
    std::string value; long firstIndex = 0, length = 0;
    std::vector<Value> subs;
    std::string typeName() const override { return "Match"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n.empty() || n == "value") return Value::str(value);
        if (n == "firstindex")         return Value::integer((int32_t)firstIndex);
        if (n == "length")             return Value::integer((int32_t)length);
        if (n == "submatches") {
            if (args.empty()) return makeArray(subs);
            long i = args[0].toI64();
            if (i < 0 || i >= (long)subs.size()) raiseErr(9, "Subscript out of range");
            return subs[i];
        }
        raiseErr(438, "Match: unknown member " + name);
    }
    bool tryDefault(Interpreter&, Value& out) override { out = Value::str(value); return true; }
};
struct MatchCollection : IObject {
    std::vector<Value> items;
    std::string typeName() const override { return "MatchCollection"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "count") return Value::integer((int32_t)items.size());
        if (n.empty() || n == "item") {
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            long i = args[0].toI64();
            if (i < 0 || i >= (long)items.size()) raiseErr(9, "Subscript out of range");
            return items[i];
        }
        raiseErr(438, "MatchCollection: unknown member " + name);
    }
    bool enumerate(Interpreter&, std::vector<Value>& out) override { out = items; return true; }
};
struct RegExpObject : IObject {
    std::string pattern;
    bool global = false, ignoreCase = false, multiLine = false;
    std::string typeName() const override { return "RegExp"; }
    std::regex build() {
        auto f = std::regex::ECMAScript;
        if (ignoreCase) f |= std::regex::icase;
        if (multiLine)  f |= std::regex::multiline;
        try { return std::regex(pattern, f); }
        catch (const std::exception& e) { raiseErr(5017, std::string("Invalid regex pattern: ") + e.what()); }
    }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "pattern")    return Value::str(pattern);
        if (n == "global")     return Value::boolean(global);
        if (n == "ignorecase") return Value::boolean(ignoreCase);
        if (n == "multiline")  return Value::boolean(multiLine);
        if (n == "test") {
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            std::regex re = build();
            return Value::boolean(std::regex_search(args[0].toStr(), re));
        }
        if (n == "replace") {
            if (args.size() < 2) raiseErr(450, "Wrong number of arguments");
            std::regex re = build();
            std::string input = args[0].toStr(), rep = args[1].toStr();
            auto fl = std::regex_constants::format_default;
            if (!global) fl |= std::regex_constants::format_first_only;
            return Value::str(std::regex_replace(input, re, rep, fl));
        }
        if (n == "execute") {
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            std::regex re = build();
            std::string input = args[0].toStr();
            auto mc = std::make_shared<MatchCollection>();
            auto begin = std::sregex_iterator(input.begin(), input.end(), re);
            auto end   = std::sregex_iterator();
            for (auto it = begin; it != end; ++it) {
                std::smatch m = *it;
                auto mo = std::make_shared<MatchObject>();
                mo->value = m.str(0);
                mo->firstIndex = (long)m.position(0);
                mo->length = (long)m.length(0);
                for (size_t g = 1; g < m.size(); ++g) mo->subs.push_back(Value::str(m.str(g)));
                mc->items.push_back(Value::object(mo));
                if (!global) break;
            }
            return Value::object(mc);
        }
        raiseErr(438, "RegExp: unknown member " + name);
    }
    void set(Interpreter&, const std::string& name, std::vector<Value>&, const Value& v, bool) override {
        std::string n = toLower(name);
        if (n == "pattern")         pattern = v.toStr();
        else if (n == "global")     global = v.toBool();
        else if (n == "ignorecase") ignoreCase = v.toBool();
        else if (n == "multiline")  multiLine = v.toBool();
        else raiseErr(438, "RegExp: cannot set " + name);
    }
};

// ---- Scripting.FileSystemObject + TextStream ----
struct TextStreamObject : IObject {
    std::vector<std::string> lines;   // for reading
    size_t linePos = 0;
    std::shared_ptr<std::ofstream> out;   // for writing
    std::string typeName() const override { return "TextStream"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "writeline") { if (out) { *out << (args.empty() ? "" : args[0].toConcatStr()) << "\n"; } return Value::empty(); }
        if (n == "write")     { if (out) { *out << (args.empty() ? "" : args[0].toConcatStr()); } return Value::empty(); }
        if (n == "readline")  { if (linePos < lines.size()) return Value::str(lines[linePos++]); raiseErr(62, "Input past end of file"); }
        if (n == "readall")   { std::string all; for (size_t k = linePos; k < lines.size(); ++k) { if (k > linePos) all += "\n"; all += lines[k]; } linePos = lines.size(); return Value::str(all); }
        if (n == "atendofstream") return Value::boolean(linePos >= lines.size());
        if (n == "close")     { if (out) { out->close(); out.reset(); } return Value::empty(); }
        raiseErr(438, "TextStream: unknown member " + name);
    }
};

// serialFromYMDHMS is defined later (interpreter section); declare it for the date helpers here.
static double serialFromYMDHMS(int y, int mo, int d, int h, int mi, int s);

#ifdef _WIN32
static std::string fsNorm(std::string p) { return p; }
#else
static std::string fsNorm(std::string p) { for (char& c : p) if (c == '\\') c = '/'; return p; }
#endif

// last-write-time -> VBScript serial (approximate C++17 file_clock -> system_clock conversion)
static double fsSerial(const fs::path& p) {
    std::error_code ec;
    auto ft = fs::last_write_time(p, ec);
    if (ec) return 0.0;
    auto sctp = std::chrono::time_point_cast<std::chrono::system_clock::duration>(
        ft - fs::file_time_type::clock::now() + std::chrono::system_clock::now());
    std::time_t tt = std::chrono::system_clock::to_time_t(sctp);
    std::tm lt = *std::localtime(&tt);
    return serialFromYMDHMS(lt.tm_year + 1900, lt.tm_mon + 1, lt.tm_mday, lt.tm_hour, lt.tm_min, lt.tm_sec);
}
static long fsAttrBits(const fs::path& p) {   // approximate VBScript attribute bits
    std::error_code ec; long a = 0;
    if (fs::is_directory(p, ec)) a |= 16;                 // Directory
    auto st = fs::status(p, ec);
    if (!ec && (st.permissions() & fs::perms::owner_write) == fs::perms::none) a |= 1;  // ReadOnly
    std::string fn = p.filename().string();
    if (!fn.empty() && fn[0] == '.') a |= 2;              // Hidden (Unix dotfile convention)
    return a;
}
static double fsSize(const fs::path& p) {
    std::error_code ec;
    if (fs::is_regular_file(p, ec)) return (double)fs::file_size(p, ec);
    if (fs::is_directory(p, ec)) {
        uintmax_t total = 0;
        for (auto& e : fs::recursive_directory_iterator(p, ec)) {
            std::error_code e2; if (e.is_regular_file(e2)) total += e.file_size(e2);
        }
        return (double)total;
    }
    return 0.0;
}
static Value fsOpenText(const fs::path& file, int mode) {   // 1 read, 2 write, 8 append
    auto ts = std::make_shared<TextStreamObject>();
    if (mode == 1) {
        std::ifstream f(file);
        if (!f) raiseErr(53, "File not found");
        std::string ln; while (std::getline(f, ln)) { if (!ln.empty() && ln.back() == '\r') ln.pop_back(); ts->lines.push_back(ln); }
    } else {
        auto flags = std::ios::out | (mode == 8 ? std::ios::app : std::ios::trunc);
        ts->out = std::make_shared<std::ofstream>(file, flags);
        if (!*ts->out) raiseErr(76, "Path not found");
    }
    return Value::object(ts);
}

static Value makeFileObj(const fs::path& p);
static Value makeFolderObj(const fs::path& p);

// ---- Files / SubFolders collections (enumerable, Count, Item(name)) ----
struct FilesCollection : IObject {
    fs::path dir;
    explicit FilesCollection(fs::path d) : dir(std::move(d)) {}
    std::string typeName() const override { return "Files"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "count") { long c = 0; std::error_code ec; for (auto& e : fs::directory_iterator(dir, ec)) { std::error_code e2; if (e.is_regular_file(e2)) ++c; } return Value::integer((int32_t)c); }
        if (n.empty() || n == "item") {
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            fs::path p = dir / fsNorm(args[0].toStr()); std::error_code ec;
            if (!fs::is_regular_file(p, ec)) raiseErr(53, "File not found");
            return makeFileObj(p);
        }
        raiseErr(438, "Files: unknown member " + name);
    }
    bool enumerate(Interpreter&, std::vector<Value>& out) override {
        std::error_code ec;
        for (auto& e : fs::directory_iterator(dir, ec)) { std::error_code e2; if (e.is_regular_file(e2)) out.push_back(makeFileObj(e.path())); }
        return true;
    }
};
struct FoldersCollection : IObject {
    fs::path dir;
    explicit FoldersCollection(fs::path d) : dir(std::move(d)) {}
    std::string typeName() const override { return "Folders"; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "count") { long c = 0; std::error_code ec; for (auto& e : fs::directory_iterator(dir, ec)) { std::error_code e2; if (e.is_directory(e2)) ++c; } return Value::integer((int32_t)c); }
        if (n.empty() || n == "item") {
            if (args.empty()) raiseErr(450, "Wrong number of arguments");
            fs::path p = dir / fsNorm(args[0].toStr()); std::error_code ec;
            if (!fs::is_directory(p, ec)) raiseErr(76, "Path not found");
            return makeFolderObj(p);
        }
        raiseErr(438, "Folders: unknown member " + name);
    }
    bool enumerate(Interpreter&, std::vector<Value>& out) override {
        std::error_code ec;
        for (auto& e : fs::directory_iterator(dir, ec)) { std::error_code e2; if (e.is_directory(e2)) out.push_back(makeFolderObj(e.path())); }
        return true;
    }
};

// ---- File object ----
struct FileObject : IObject {
    fs::path p;
    explicit FileObject(fs::path pp) : p(std::move(pp)) {}
    std::string typeName() const override { return "File"; }
    bool tryDefault(Interpreter&, Value& out) override { out = Value::str(p.string()); return true; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "path")                     return Value::str(p.string());
        if (n == "name" || n == "shortname") return Value::str(p.filename().string());
        if (n == "parentfolder")             return makeFolderObj(p.parent_path());
        if (n == "drive")                    return Value::str(p.root_name().string());
        if (n == "size")                     return Value::number(fsSize(p));
        if (n == "type")                     return Value::str("File");
        if (n == "attributes")               return Value::integer((int32_t)fsAttrBits(p));
        if (n == "datecreated" || n == "datelastmodified" || n == "datelastaccessed") return Value::number(fsSerial(p));
        if (n == "delete")                   { std::error_code ec; fs::remove(p, ec); if (ec) raiseErr(70, "Permission denied"); return Value::empty(); }
        if (n == "copy")                     { if (args.empty()) raiseErr(450, "Wrong number of arguments"); std::error_code ec; auto opt = (args.size() > 1 && !args[1].toBool()) ? fs::copy_options::none : fs::copy_options::overwrite_existing; fs::copy_file(p, fs::path(fsNorm(args[0].toStr())), opt, ec); if (ec) raiseErr(53, "Copy failed"); return Value::empty(); }
        if (n == "move")                     { if (args.empty()) raiseErr(450, "Wrong number of arguments"); fs::path d = fsNorm(args[0].toStr()); std::error_code ec; fs::rename(p, d, ec); if (ec) raiseErr(53, "Move failed"); p = d; return Value::empty(); }
        if (n == "openastextstream")         { int mode = args.empty() ? 1 : (int)args[0].toI64(); return fsOpenText(p, mode); }
        raiseErr(438, "File: unknown member " + name);
    }
};

// ---- Folder object ----
struct FolderObject : IObject {
    fs::path p;
    explicit FolderObject(fs::path pp) : p(std::move(pp)) {}
    std::string typeName() const override { return "Folder"; }
    bool tryDefault(Interpreter&, Value& out) override { out = Value::str(p.string()); return true; }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        if (n == "path")                     return Value::str(p.string());
        if (n == "name" || n == "shortname") { std::string f = p.filename().string(); return Value::str(f.empty() ? p.string() : f); }
        if (n == "parentfolder")             return makeFolderObj(p.parent_path());
        if (n == "drive")                    return Value::str(p.root_name().string());
        if (n == "size")                     return Value::number(fsSize(p));
        if (n == "type")                     return Value::str("File Folder");
        if (n == "attributes")               return Value::integer((int32_t)fsAttrBits(p));
        if (n == "datecreated" || n == "datelastmodified" || n == "datelastaccessed") return Value::number(fsSerial(p));
        if (n == "isrootfolder")             return Value::boolean(p == p.root_path());
        if (n == "files")                    return Value::object(std::make_shared<FilesCollection>(p));
        if (n == "subfolders")               return Value::object(std::make_shared<FoldersCollection>(p));
        if (n == "delete")                   { std::error_code ec; fs::remove_all(p, ec); if (ec) raiseErr(70, "Permission denied"); return Value::empty(); }
        if (n == "copy")                     { if (args.empty()) raiseErr(450, "Wrong number of arguments"); std::error_code ec; auto opt = fs::copy_options::recursive | ((args.size() > 1 && !args[1].toBool()) ? fs::copy_options::none : fs::copy_options::overwrite_existing); fs::copy(p, fs::path(fsNorm(args[0].toStr())), opt, ec); if (ec) raiseErr(76, "Copy failed"); return Value::empty(); }
        if (n == "move")                     { if (args.empty()) raiseErr(450, "Wrong number of arguments"); fs::path d = fsNorm(args[0].toStr()); std::error_code ec; fs::rename(p, d, ec); if (ec) raiseErr(76, "Move failed"); p = d; return Value::empty(); }
        if (n == "createtextfile")           { if (args.empty()) raiseErr(450, "Wrong number of arguments"); fs::path f = p / fsNorm(args[0].toStr()); auto ts = std::make_shared<TextStreamObject>(); ts->out = std::make_shared<std::ofstream>(f, std::ios::out | std::ios::trunc); if (!*ts->out) raiseErr(76, "Path not found"); return Value::object(ts); }
        raiseErr(438, "Folder: unknown member " + name);
    }
};

static Value makeFileObj(const fs::path& p)   { return Value::object(std::make_shared<FileObject>(p)); }
static Value makeFolderObj(const fs::path& p) { return Value::object(std::make_shared<FolderObject>(p)); }

struct FileSystemObject : IObject {
    std::string typeName() const override { return "FileSystemObject"; }
    // VBScript paths use '\'. On non-Windows, treat it as the separator too.
#ifdef _WIN32
    static std::string norm(std::string p) { return p; }
#else
    static std::string norm(std::string p) { for (char& c : p) if (c == '\\') c = '/'; return p; }
#endif
    // If dst names an existing directory (or ends in a separator), append src's file name.
    static std::string resolveDest(const std::string& src, std::string dst) {
        std::error_code ec;
        bool dir = fs::is_directory(dst, ec) || (!dst.empty() && (dst.back() == '/' || dst.back() == '\\'));
        if (dir) return (fs::path(dst) / fs::path(src).filename()).string();
        return dst;
    }
    static std::string baseName(const std::string& p) {
        size_t s = p.find_last_of("/\\"); std::string f = s == std::string::npos ? p : p.substr(s + 1);
        size_t d = f.find_last_of('.'); return d == std::string::npos ? f : f.substr(0, d);
    }
    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        std::string n = toLower(name);
        auto need = [&](size_t k){ if (args.size() < k) raiseErr(450, "Wrong number of arguments"); };
        if (n == "fileexists")   { need(1); std::error_code ec; return Value::boolean(fs::is_regular_file(norm(args[0].toStr()), ec)); }
        if (n == "folderexists") { need(1); std::error_code ec; return Value::boolean(fs::is_directory(norm(args[0].toStr()), ec)); }
        if (n == "createfolder") {
            need(1); std::string p = norm(args[0].toStr()); std::error_code ec;
            if (fs::exists(p, ec)) raiseErr(58, "File already exists");
            fs::create_directory(p, ec);
            if (ec) raiseErr(76, "Path not found");
            return makeFolderObj(p);
        }
        if (n == "getfolder") {
            need(1); std::string p = norm(args[0].toStr()); std::error_code ec;
            if (!fs::is_directory(p, ec)) raiseErr(76, "Path not found");
            return makeFolderObj(p);
        }
        if (n == "getfile") {
            need(1); std::string p = norm(args[0].toStr()); std::error_code ec;
            if (!fs::is_regular_file(p, ec)) raiseErr(53, "File not found");
            return makeFileObj(p);
        }
        if (n == "getspecialfolder") {
            long k = args.empty() ? 0 : args[0].toI64(); std::error_code ec;
            if (k == 2) return makeFolderObj(fs::temp_directory_path(ec));   // 2 = TemporaryFolder
            return makeFolderObj(fs::current_path(ec));                       // 0/1 best-effort
        }
        if (n == "deletefile") {
            need(1); std::string p = norm(args[0].toStr()); std::error_code ec;
            if (!fs::exists(p, ec)) raiseErr(53, "File not found");
            fs::remove(p, ec);
            if (ec) raiseErr(70, "Permission denied");
            return Value::empty();
        }
        if (n == "deletefolder") {
            need(1); std::string p = norm(args[0].toStr()); std::error_code ec;
            if (!fs::exists(p, ec)) raiseErr(76, "Path not found");
            fs::remove_all(p, ec);
            if (ec) raiseErr(70, "Permission denied");
            return Value::empty();
        }
        if (n == "copyfile") {
            need(2); std::error_code ec;
            std::string src = norm(args[0].toStr()), dst = resolveDest(src, norm(args[1].toStr()));
            auto opt = (args.size() > 2 && !args[2].toBool()) ? fs::copy_options::none
                                                              : fs::copy_options::overwrite_existing;
            fs::copy_file(src, dst, opt, ec);
            if (ec) raiseErr(53, "File not found or copy failed");
            return Value::empty();
        }
        if (n == "copyfolder") {
            need(2); std::error_code ec;
            auto opt = fs::copy_options::recursive |
                       ((args.size() > 2 && !args[2].toBool()) ? fs::copy_options::none
                                                               : fs::copy_options::overwrite_existing);
            fs::copy(norm(args[0].toStr()), norm(args[1].toStr()), opt, ec);
            if (ec) raiseErr(76, "Path not found or copy failed");
            return Value::empty();
        }
        if (n == "movefile" || n == "movefolder") {
            need(2); std::error_code ec;
            std::string src = norm(args[0].toStr()), dst = resolveDest(src, norm(args[1].toStr()));
            fs::rename(src, dst, ec);
            if (ec) raiseErr(53, "Move failed (source not found or cross-device)");
            return Value::empty();
        }
        if (n == "getabsolutepathname") {
            need(1); std::error_code ec; auto p = fs::absolute(norm(args[0].toStr()), ec);
            return Value::str(ec ? norm(args[0].toStr()) : p.string());
        }
        if (n == "gettempname") {
            static unsigned seed = (unsigned)std::time(nullptr);
            seed = seed * 1103515245u + 12345u;
            char buf[32]; std::snprintf(buf, sizeof(buf), "dir%05u.tmp", (seed >> 8) % 100000u);
            return Value::str(buf);
        }
        if (n == "createtextfile") {
            need(1);
            auto ts = std::make_shared<TextStreamObject>();
            ts->out = std::make_shared<std::ofstream>(norm(args[0].toStr()), std::ios::out | std::ios::trunc);
            if (!*ts->out) raiseErr(76, "Path not found");
            return Value::object(ts);
        }
        if (n == "opentextfile") {
            need(1);
            int mode = args.size() > 1 ? (int)args[1].toI64() : 1;   // 1 read, 2 write, 8 append
            auto ts = std::make_shared<TextStreamObject>();
            if (mode == 1) {
                std::ifstream f(norm(args[0].toStr()));
                if (!f) raiseErr(53, "File not found");
                std::string ln; while (std::getline(f, ln)) { if (!ln.empty() && ln.back() == '\r') ln.pop_back(); ts->lines.push_back(ln); }
            } else {
                auto flags = std::ios::out | (mode == 8 ? std::ios::app : std::ios::trunc);
                ts->out = std::make_shared<std::ofstream>(norm(args[0].toStr()), flags);
                if (!*ts->out) raiseErr(76, "Path not found");
            }
            return Value::object(ts);
        }
        if (n == "getbasename")      { need(1); return Value::str(baseName(norm(args[0].toStr()))); }
        if (n == "getfilename")      { need(1); std::string p = norm(args[0].toStr()); size_t s = p.find_last_of("/\\"); return Value::str(s == std::string::npos ? p : p.substr(s + 1)); }
        if (n == "getextensionname") { need(1); std::string p = norm(args[0].toStr()); size_t d = p.find_last_of('.'); return Value::str(d == std::string::npos ? "" : p.substr(d + 1)); }
        if (n == "getparentfoldername"){ need(1); std::string p = norm(args[0].toStr()); size_t s = p.find_last_of("/\\"); return Value::str(s == std::string::npos ? "" : p.substr(0, s)); }
        if (n == "buildpath")        { need(2); std::string a = norm(args[0].toStr()), b = norm(args[1].toStr()); if (!a.empty() && a.back() != '/' && a.back() != '\\') a += "/"; return Value::str(a + b); }
        raiseErr(438, "FileSystemObject: unknown member " + name);
    }
};

// ================================================================ INTERPRETER
#include <cstdlib>

// Date <-> VBScript serial (days since 1899-12-30). Hinnant's civil algorithms.
static long daysFromCivil(int y, unsigned m, unsigned d) {
    y -= m <= 2;
    long era = (y >= 0 ? y : y - 399) / 400;
    unsigned yoe = (unsigned)(y - era * 400);
    unsigned doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
    unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097L + (long)doe - 719468L;
}
static void civilFromDays(long z, int& y, unsigned& m, unsigned& d) {
    z += 719468;
    long era = (z >= 0 ? z : z - 146096) / 146097;
    unsigned doe = (unsigned)(z - era * 146097);
    unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    y = (int)(yoe) + (int)(era * 400);
    unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    unsigned mp = (5 * doy + 2) / 153;
    d = doy - (153 * mp + 2) / 5 + 1;
    m = mp + (mp < 10 ? 3 : -9);
    y += (m <= 2);
}
static const long kBaseDay = daysFromCivil(1899, 12, 30);
static double serialFromYMDHMS(int y, int mo, int d, int h, int mi, int s) {
    long dd = daysFromCivil(y, (unsigned)mo, (unsigned)d) - kBaseDay;
    return (double)dd + (h * 3600 + mi * 60 + s) / 86400.0;
}
static void ymdhmsFromSerial(double serial, int& y, int& mo, int& d, int& h, int& mi, int& s) {
    long dd = (long)std::floor(serial);
    double frac = serial - (double)dd;
    unsigned um, ud; int yy; civilFromDays(dd + kBaseDay, yy, um, ud);
    y = yy; mo = (int)um; d = (int)ud;
    long secs = (long)std::llround(frac * 86400.0);
    h = (int)(secs / 3600); mi = (int)((secs % 3600) / 60); s = (int)(secs % 60);
}

// COM factory (defined in the Windows section; returns ok=false off-Windows).
static ObjectPtr createComObject(const std::string& progid, bool& ok);
// Windows-only automation objects (real impl in the platform section; stubs raise on other OSes):
static ObjectPtr makeMouseObj();
static ObjectPtr makeScreenObj();
static ObjectPtr makeSoundObj();
// Clipboard text access (ClipPut/ClipGet); Windows-only, stubs raise elsewhere.
static bool clipboardSet(const std::string& text);
static bool clipboardGet(std::string& out);

// ---- WAV combiner: concatenates same-format PCM WAV files into one output ----
// Pure file I/O, so this works on every platform (no audio device needed).
struct WavFmt { uint16_t channels = 0; uint32_t rate = 0; uint16_t bits = 0; };

static bool readWavFile(const std::string& path, WavFmt& fmt, std::string& data) {
    std::ifstream f(path, std::ios::binary);
    if (!f) return false;
    char id[4]; uint32_t sz = 0;
    if (!f.read(id, 4) || std::string(id, 4) != "RIFF") return false;
    f.read(reinterpret_cast<char*>(&sz), 4);
    if (!f.read(id, 4) || std::string(id, 4) != "WAVE") return false;
    bool gotFmt = false, gotData = false;
    while (f.read(id, 4) && f.read(reinterpret_cast<char*>(&sz), 4)) {
        std::string chunk(id, 4);
        if (chunk == "fmt ") {
            uint16_t audio = 0, ch = 0, ba = 0, bits = 0; uint32_t rate = 0, br = 0;
            f.read(reinterpret_cast<char*>(&audio), 2);
            f.read(reinterpret_cast<char*>(&ch), 2);
            f.read(reinterpret_cast<char*>(&rate), 4);
            f.read(reinterpret_cast<char*>(&br), 4);
            f.read(reinterpret_cast<char*>(&ba), 2);
            f.read(reinterpret_cast<char*>(&bits), 2);
            fmt.channels = ch; fmt.rate = rate; fmt.bits = bits;
            gotFmt = true;
            if (audio != 1) return false;              // only uncompressed PCM
            if (sz > 16) f.seekg(sz - 16, std::ios::cur);
        } else if (chunk == "data") {
            data.resize(sz);
            if (sz) f.read(&data[0], sz);
            gotData = true;
            break;
        } else {
            f.seekg(sz + (sz & 1), std::ios::cur);     // skip unknown chunk (+ pad byte)
        }
    }
    return gotFmt && gotData;
}

struct WavWriter : IObject {
    std::ofstream out;
    WavFmt fmt;
    uint32_t dataBytes = 0;
    bool haveFmt = false, isOpen = false;
    std::string path;

    ~WavWriter() override { if (isOpen) finalize(); }

    void writeHeader() {
        uint32_t byteRate   = fmt.rate * fmt.channels * (fmt.bits / 8);
        uint16_t blockAlign = (uint16_t)(fmt.channels * (fmt.bits / 8));
        uint32_t chunkSize  = 36 + dataBytes;
        uint16_t audio = 1;  uint32_t sc1 = 16;
        out.seekp(0);
        out.write("RIFF", 4); out.write(reinterpret_cast<char*>(&chunkSize), 4); out.write("WAVE", 4);
        out.write("fmt ", 4); out.write(reinterpret_cast<char*>(&sc1), 4);
        out.write(reinterpret_cast<char*>(&audio), 2);
        out.write(reinterpret_cast<char*>(&fmt.channels), 2);
        out.write(reinterpret_cast<char*>(&fmt.rate), 4);
        out.write(reinterpret_cast<char*>(&byteRate), 4);
        out.write(reinterpret_cast<char*>(&blockAlign), 2);
        out.write(reinterpret_cast<char*>(&fmt.bits), 2);
        out.write("data", 4); out.write(reinterpret_cast<char*>(&dataBytes), 4);
    }
    void finalize() {
        if (!haveFmt) { fmt.channels = 2; fmt.rate = 44100; fmt.bits = 16; }
        writeHeader();
        out.close();
        isOpen = false;
    }

    std::string typeName() const override { return "WavWriter"; }
    Value get(Interpreter&, const std::string& member, std::vector<Value>& args) override {
        std::string m = toLower(member);
        if (m == "create") {
            if (args.empty()) raiseErr(5, "WavWriter.Create requires a path");
            path = args[0].toStr();
            out.open(path, std::ios::binary | std::ios::trunc);
            if (!out) raiseErr(76, "WavWriter.Create: cannot open '" + path + "'");
            char zero[44] = {0};
            out.write(zero, 44);                       // placeholder, patched on Close
            dataBytes = 0; haveFmt = false; isOpen = true;
            return Value::empty();
        }
        if (m == "append" || m == "add") {
            if (!isOpen) raiseErr(5, "WavWriter: call Create before Append");
            if (args.empty()) raiseErr(5, "WavWriter.Append requires a path");
            WavFmt f; std::string data;
            if (!readWavFile(args[0].toStr(), f, data))
                raiseErr(76, "WavWriter.Append: cannot read PCM WAV '" + args[0].toStr() + "'");
            if (!haveFmt) { fmt = f; haveFmt = true; }
            else if (f.channels != fmt.channels || f.rate != fmt.rate || f.bits != fmt.bits)
                raiseErr(5, "WavWriter.Append: format mismatch (all files must share channels/sample-rate/bit-depth)");
            if (!data.empty()) out.write(data.data(), (std::streamsize)data.size());
            dataBytes += (uint32_t)data.size();
            return Value::empty();
        }
        if (m == "close") { if (isOpen) finalize(); return Value::empty(); }
        if (m == "duration") {                          // seconds written so far
            double bps = haveFmt ? (double)fmt.rate * fmt.channels * (fmt.bits / 8) : 0.0;
            return Value::number(bps > 0 ? dataBytes / bps : 0.0);
        }
        raiseErr(438, "WavWriter has no member '" + member + "'");
        return Value::empty();
    }
};

// ---- parsed class shape ----
struct ClassInfo {
    std::string name;
    std::vector<std::string> fields;                       // field names (original case)
    std::unordered_map<std::string, StmtP> methods;        // lower -> Sub/Function
    std::unordered_map<std::string, StmtP> propGet, propLet, propSet;
    std::string defaultMember;                             // lower name of the Default member ("" if none)
};

// ---- an instance of a script-defined class ----
struct ClassInstance : IObject {
    std::shared_ptr<ClassInfo> cls;
    std::unordered_map<std::string, Value> fields;         // lower -> value
    Interpreter* interp = nullptr;                         // set at creation; used to run Class_Terminate
    ~ClassInstance();                                      // runs Class_Terminate (defined after Interpreter)
    std::string typeName() const override { return cls ? cls->name : "Object"; }
    Value get(Interpreter& in, const std::string& name, std::vector<Value>& args) override;
    void set(Interpreter& in, const std::string& name, std::vector<Value>& args,
             const Value& val, bool isSet) override;
    bool tryDefault(Interpreter& in, Value& out) override;   // fires the Default member (defined below)
};

class Interpreter {
public:
    Interpreter() { frames.reserve(64); framePool.reserve(64); setupConstants(); setupObjects(); registerBuiltins(); }
    ~Interpreter() { shuttingDown = true; }   // stop terminators once members start being destroyed

    // Run a class instance's Class_Terminate (if any). Called from ClassInstance's
    // destructor, so it must never throw and must be inert during shutdown.
    void runTerminator(ClassInstance* self) {
        if (shuttingDown || !self || !self->cls) return;
        auto it = self->cls->methods.find("class_terminate");
        if (it == self->cls->methods.end()) return;
        try { std::vector<Value> none; invoke(it->second, none, self, nullptr); }
        catch (...) { /* a terminator must not escape a destructor */ }
    }

    void run(std::vector<StmtP>& program) {
        for (auto& s : program) {          // pass 1: hoist procs + classes
            if (s->k == Stmt::SubDecl || s->k == Stmt::FuncDecl) procs[toLower(s->name)] = s;
            else if (s->k == Stmt::ClassDecl) registerClass(s);
        }
        execBlock(program);                // pass 2: execute top-level in order
        globals.clear();                   // drop module-level objects now (interpreter still alive) so
                                           // their Class_Terminate runs before teardown
    }

    // Execute / ExecuteGlobal. Parses and runs a code string.
    //   global == false : runs in the current scope (Execute)
    //   global == true  : runs at module/global scope (ExecuteGlobal), even when
    //                     called from inside a Sub/Function or class method.
    void executeString(const std::string& src, bool global) {
        Lexer lx(src);
        Parser p(lx.tokenize());
        std::vector<StmtP> prog = p.parseProgram();
        for (auto& s : prog) {             // hoist any procs/classes it defines
            if (s->k == Stmt::SubDecl || s->k == Stmt::FuncDecl) procs[toLower(s->name)] = s;
            else if (s->k == Stmt::ClassDecl) registerClass(s);
        }
        if (global && (!frames.empty() || !instanceStack.empty())) {
            auto savedFrames = std::move(frames);
            auto savedInstances = std::move(instanceStack);
            frames.clear(); instanceStack.clear();
            try { execBlock(prog); }
            catch (...) { frames = std::move(savedFrames); instanceStack = std::move(savedInstances); throw; }
            frames = std::move(savedFrames); instanceStack = std::move(savedInstances);
        } else {
            execBlock(prog);
        }
    }

    // Eval("expr") — parse and evaluate a single expression in the current scope.
    Value evalExpression(const std::string& src) {
        Lexer lx(src);
        Parser p(lx.tokenize());
        ExprP e = p.parseOneExpr();
        return eval(e.get());
    }
    // Call a Sub/Function/builtin by (lowercased) name — used by function references.
    Value callByName(const std::string& low, std::vector<Value>& args) {
        auto it = procs.find(low);
        if (it != procs.end()) return invoke(it->second, args, nullptr, nullptr);
        auto b = builtins.find(low);
        if (b != builtins.end()) return b->second(*this, args);
        raiseErr(438, "Unknown procedure or function: " + low);
    }
    Value makeFuncRef(const std::string& name);   // defined after FuncRefObject

    // Public hooks used by ClassInstance:
    int errorLine() const { return curLine; }
    int errorCol()  const { return curCol; }
    // Record the running script's path so Include resolution and Directive.Script* work.
    void setScriptInfo(const std::string& fullPath, const std::string& dir) {
        if (!directiveObj) return;
        directiveObj->scriptFullName = fullPath;
        directiveObj->scriptDir = dir;
        if (!fullPath.empty()) directiveObj->scriptName = fs::path(fullPath).filename().string();
    }
    void setArguments(const std::vector<std::string>& a) { if (directiveObj) directiveObj->scriptArgs = a; }
    Value callClassMethod(ClassInstance* self, StmtP decl, std::vector<Value>& argv) {
        return invoke(decl, argv, self, nullptr);
    }
    Value arrayGetPub(Value& av, std::vector<Value>& args) { return arrayGet(av, args); }
    void arraySetPub(Value& av, std::vector<Value>& args, const Value& v) { arraySet(av, args, v); }

private:
    std::unordered_map<std::string, Value> globals;
    std::vector<std::unordered_map<std::string, Value>> frames;
    std::vector<std::unordered_map<std::string, Value>> framePool;   // recycled frames (keep bucket allocations)
    std::unordered_map<std::string, StmtP> procs;
    std::unordered_map<std::string, std::shared_ptr<ClassInfo>> classes;
    std::unordered_map<std::string, std::function<Value(Interpreter&, std::vector<Value>&)>> builtins;
    std::shared_ptr<ErrObject> errObj;
    std::shared_ptr<DirectiveObject> directiveObj;   // host object; also carries script path info
    std::vector<ClassInstance*> instanceStack;
    std::vector<Value> withStack;
    bool onErrorResumeNext = false;
    bool optionExplicit = false;
    bool shuttingDown = false;        // set during teardown so Class_Terminate never runs against half-freed state
    int curLine = 0;                  // line of the statement currently executing (for error messages)
    int curCol = 0;                   // its starting column

    // ---------------------------------------------------------- helpers
    static bool truth(const Value& v) { return v.isNull() ? false : v.toBool(); }
    static Value cloneIfArray(const Value& v) {
        if (!v.isArray()) return v;
        auto nd = std::make_shared<ArrayData>(*v.arr);
        return Value::array(nd);
    }
    static Value fitInt(long long r) {
        if (r >= INT32_MIN && r <= INT32_MAX) return Value::integer((int32_t)r);
        return Value::number((double)r);
    }
    static Value numResult(double r, const Value& a, const Value& b) {
        if (a.type == Value::Type::Int && b.type == Value::Type::Int &&
            std::floor(r) == r && r >= INT32_MIN && r <= INT32_MAX)
            return Value::integer((int32_t)r);
        return Value::number(r);
    }
    static int cmp(const Value& a, const Value& b) {
        bool as = a.type == Value::Type::String, bs = b.type == Value::Type::String;
        if (as && bs) { int c = a.s.compare(b.s); return c < 0 ? -1 : (c > 0 ? 1 : 0); }
        if (as || bs) {
            const Value& sv = as ? a : b;
            if (!sv.looksNumeric()) {
                std::string x = a.toStr(), y = b.toStr();
                int c = x.compare(y); return c < 0 ? -1 : (c > 0 ? 1 : 0);
            }
        }
        double x = a.toDouble(), y = b.toDouble();
        return x < y ? -1 : (x > y ? 1 : 0);
    }
    std::vector<Value> evalArgs(std::vector<ExprP>& es) {
        std::vector<Value> v; v.reserve(es.size());
        for (auto& e : es) v.push_back(eval(e.get()));
        return v;
    }
    StmtP classCallable(ClassInstance* self, const std::string& low) {
        auto& ci = *self->cls;
        auto g = ci.propGet.find(low); if (g != ci.propGet.end()) return g->second;
        auto m = ci.methods.find(low); if (m != ci.methods.end()) return m->second;
        return StmtP();
    }
    Value* varPtr(const std::string& low) {
        if (!frames.empty()) { auto it = frames.back().find(low); if (it != frames.back().end()) return &it->second; }
        if (!instanceStack.empty()) { auto& f = instanceStack.back()->fields; auto it = f.find(low); if (it != f.end()) return &it->second; }
        auto it = globals.find(low); if (it != globals.end()) return &it->second;
        return nullptr;
    }

    // ---------------------------------------------------------- arrays
    Value arrayGet(Value& av, std::vector<Value>& args) {
        std::vector<int> idx; for (auto& a : args) idx.push_back((int)a.toI64());
        return av.arr->data[av.arr->flatIndex(idx)];
    }
    void arraySet(Value& av, std::vector<Value>& args, const Value& v) {
        std::vector<int> idx; for (auto& a : args) idx.push_back((int)a.toI64());
        av.arr->data[av.arr->flatIndex(idx)] = v;
    }

    // ---------------------------------------------------------- scope R/W
    // Fast paths take an already-lowercased name (precomputed on the AST node).
    Value getVarValueL(const std::string& low) {
        if (!frames.empty()) { auto it = frames.back().find(low); if (it != frames.back().end()) return it->second; }
        if (!instanceStack.empty()) { auto& f = instanceStack.back()->fields; auto it = f.find(low); if (it != f.end()) return it->second; }
        { auto it = globals.find(low); if (it != globals.end()) return it->second; }
        if (!instanceStack.empty()) { StmtP d = classCallable(instanceStack.back(), low); if (d) { std::vector<Value> none; return callClassMethod(instanceStack.back(), d, none); } }
        if (procs.count(low))    { std::vector<ExprP> none; return callProc(procs[low], none, nullptr); }
        if (builtins.count(low)) { std::vector<Value> none; return builtins[low](*this, none); }
        if (optionExplicit) raiseErr(500, "Variable is undefined: '" + low + "'");
        return Value::empty();
    }
    void setVarValueL(const std::string& low, const Value& val) {
        if (!frames.empty()) { auto it = frames.back().find(low); if (it != frames.back().end()) { it->second = val; return; } }
        if (!instanceStack.empty()) { auto& f = instanceStack.back()->fields; auto it = f.find(low); if (it != f.end()) { it->second = val; return; } }
        { auto it = globals.find(low); if (it != globals.end()) { it->second = val; return; } }
        if (optionExplicit) raiseErr(500, "Variable is undefined: '" + low + "'");
        if (!frames.empty()) frames.back()[low] = val; else globals[low] = val;
    }
    Value getVarValue(const std::string& name) { return getVarValueL(toLower(name)); }
    void setVarValue(const std::string& name, const Value& val) { setVarValueL(toLower(name), val); }
    void declareVar(const std::string& name) {
        std::string low = toLower(name);
        if (!frames.empty()) frames.back()[low] = Value::empty(); else globals[low] = Value::empty();
    }
    void putVar(const std::string& name, const Value& v) {   // force-create in current scope
        std::string low = toLower(name);
        if (!frames.empty()) frames.back()[low] = v; else globals[low] = v;
    }

    // ---------------------------------------------------------- expressions
    Value eval(Expr* e) {
        switch (e->k) {
            case Expr::Lit: return e->lit;
            case Expr::Var:
                if (e->lname == "me") {
                    if (instanceStack.empty()) raiseErr(91, "'Me' is only valid inside a class");
                    return Value::object(instanceStack.back()->shared_from_this());
                }
                return getVarValueL(e->lname);
            case Expr::Unary: return unaryOp(e->name, eval(e->a.get()));
            case Expr::Binary: return binOp(e->name, eval(e->a.get()), eval(e->b.get()));
            case Expr::New: return evalNew(e->name);
            case Expr::Member: {
                Value tgt = e->a ? eval(e->a.get()) : withTarget();
                if (tgt.isNothing()) raiseErr(91, "Object variable not set");
                if (tgt.isObject() && tgt.obj) { std::vector<Value> none; return tgt.obj->get(*this, e->name, none); }
                raiseErr(424, "Object required");
            }
            case Expr::Index: return evalIndex(e);
        }
        return Value::empty();
    }
    Value withTarget() {
        if (withStack.empty()) raiseErr(91, "Invalid use of '.' outside With block");
        return withStack.back();
    }
    Value evalNew(const std::string& className) {
        std::string low = toLower(className);
        // Native engine objects (COM is reserved for CreateObject):
        if (low == "dictionary")                             return Value::object(std::make_shared<DictionaryObject>());
        if (low == "list" || low == "linkedlist")            return Value::object(std::make_shared<ListObject>());
        if (low == "regexp")                                 return Value::object(std::make_shared<RegExpObject>());
        if (low == "filesystem" || low == "fso" || low == "filesystemobject")
                                                             return Value::object(std::make_shared<FileSystemObject>());
        if (low == "wavwriter" || low == "wavfile")          return Value::object(std::make_shared<WavWriter>());
        if (low == "mouse")                                  return Value::object(makeMouseObj());
        if (low == "screen")                                 return Value::object(makeScreenObj());
        if (low == "sound")                                  return Value::object(makeSoundObj());
        auto it = classes.find(low);
        if (it == classes.end()) raiseErr(500, "Class not defined: " + className);
        auto inst = std::make_shared<ClassInstance>();
        inst->cls = it->second;
        inst->interp = this;
        for (auto& f : it->second->fields) inst->fields[toLower(f)] = Value::empty();
        auto ini = it->second->methods.find("class_initialize");
        if (ini != it->second->methods.end()) { std::vector<Value> none; invoke(ini->second, none, inst.get(), nullptr); }
        return Value::object(inst);
    }
    Value evalIndex(Expr* e) {
        Expr* tgt = e->a.get();
        if (tgt->k == Expr::Var) {
            const std::string& low = tgt->lname;
            Value* vp = varPtr(low);
            if (vp && vp->isArray())  { auto a = evalArgs(e->args); return arrayGet(*vp, a); }
            if (vp && vp->isObject()) { if (!vp->obj) raiseErr(91, "Object variable not set"); auto a = evalArgs(e->args); return vp->obj->get(*this, "", a); }
            if (!instanceStack.empty()) { StmtP d = classCallable(instanceStack.back(), low); if (d) { auto a = evalArgs(e->args); return callClassMethod(instanceStack.back(), d, a); } }
            if (procs.count(low))    { auto a = evalArgs(e->args); return callProcV(procs[low], e->args, a, nullptr); }
            if (builtins.count(low)) { auto a = evalArgs(e->args); return builtins[low](*this, a); }
            if (vp) raiseErr(13, "Type mismatch (not callable): " + tgt->name);
            raiseErr(500, "Undefined: '" + tgt->name + "'");
        }
        if (tgt->k == Expr::Member) {
            Value obj = tgt->a ? eval(tgt->a.get()) : withTarget();
            if (obj.isNothing()) raiseErr(91, "Object variable not set");
            if (obj.isObject() && obj.obj) { auto a = evalArgs(e->args); return obj.obj->get(*this, tgt->name, a); }
            raiseErr(424, "Object required");
        }
        Value tv = eval(tgt);
        auto a = evalArgs(e->args);
        if (tv.isArray()) return arrayGet(tv, a);
        if (tv.isObject() && tv.obj) return tv.obj->get(*this, "", a);
        raiseErr(13, "Type mismatch");
    }

    Value unaryOp(const std::string& op, const Value& v) {
        if (op == "-") {
            if (v.isNull()) return Value::null();
            if (v.type == Value::Type::Int) return fitInt(-(long long)v.i);
            return Value::number(-v.toDouble());
        }
        // Not
        if (v.isNull()) return Value::null();
        if (v.type == Value::Type::Bool) return Value::boolean(!v.b);
        return Value::integer((int32_t)(~v.toI64()));
    }
    // Resolve an object to its default property/value in a value context (VBScript
    // uses the class's Default member here). Non-objects pass through unchanged.
    Value defaultValue(const Value& v) {
        if (v.isObject() && v.obj) { Value out; if (v.obj->tryDefault(*this, out)) return out; }
        return v;
    }
    Value binOp(const std::string& op, const Value& aRaw, const Value& bRaw) {
        if (op == "Is") {
            IObject* pa = aRaw.isObject() ? aRaw.obj.get() : nullptr;
            IObject* pb = bRaw.isObject() ? bRaw.obj.get() : nullptr;
            return Value::boolean(pa == pb);
        }
        // Every other operator works on values: an object yields its default member.
        Value a = defaultValue(aRaw), b = defaultValue(bRaw);
        if (op == "&") return Value::str(a.toConcatStr() + b.toConcatStr());
        if (op == "And" || op == "Or" || op == "Xor" || op == "Eqv" || op == "Imp") {
            if (a.isNull() || b.isNull()) return Value::null();
            if (a.type == Value::Type::Bool && b.type == Value::Type::Bool) {
                bool x = a.b, y = b.b, r = false;
                if (op == "And") r = x && y; else if (op == "Or") r = x || y;
                else if (op == "Xor") r = x != y; else if (op == "Eqv") r = x == y; else r = (!x) || y;
                return Value::boolean(r);
            }
            long long x = a.toI64(), y = b.toI64(), r = 0;
            if (op == "And") r = x & y; else if (op == "Or") r = x | y;
            else if (op == "Xor") r = x ^ y; else if (op == "Eqv") r = ~(x ^ y); else r = (~x) | y;
            return Value::integer((int32_t)r);
        }
        if (op == "=" || op == "<>" || op == "<" || op == ">" || op == "<=" || op == ">=") {
            if (a.isNull() || b.isNull()) return Value::null();
            int c = cmp(a, b); bool r = false;
            if (op == "=") r = c == 0; else if (op == "<>") r = c != 0;
            else if (op == "<") r = c < 0; else if (op == ">") r = c > 0;
            else if (op == "<=") r = c <= 0; else r = c >= 0;
            return Value::boolean(r);
        }
        if (a.isNull() || b.isNull()) return Value::null();
        if (op == "+") {
            if (a.type == Value::Type::String && b.type == Value::Type::String) return Value::str(a.s + b.s);
            return numResult(a.toDouble() + b.toDouble(), a, b);
        }
        double x = a.toDouble(), y = b.toDouble();
        if (op == "-") return numResult(x - y, a, b);
        if (op == "*") return numResult(x * y, a, b);
        if (op == "/") { if (y == 0) raiseErr(11, "Division by zero"); return Value::number(x / y); }
        if (op == "\\") { long long iy = (long long)roundHalfEven(y); if (iy == 0) raiseErr(11, "Division by zero"); return fitInt((long long)roundHalfEven(x) / iy); }
        if (op == "Mod") { long long iy = (long long)roundHalfEven(y); if (iy == 0) raiseErr(11, "Division by zero"); return fitInt((long long)roundHalfEven(x) % iy); }
        if (op == "^") return Value::number(std::pow(x, y));
        raiseErr(1, "Unknown operator " + op);
    }

    // ---------------------------------------------------------- statements
    void execBlock(std::vector<StmtP>& b) { for (auto& s : b) execOne(s); }
    void execOne(StmtP& s) {
        if (!onErrorResumeNext) { exec(s.get()); return; }
        try { exec(s.get()); }
        catch (VbError& e) { errObj->number = e.number; errObj->source = e.source; errObj->description = e.description; }
        // ExitSignal / QuitSignal propagate
    }
    void exec(Stmt* s) {
        if (s->line) curLine = s->line;
        if (s->col)  curCol = s->col;
        switch (s->k) {
            case Stmt::OptionS: optionExplicit = true; break;
            case Stmt::OnErrorS: onErrorResumeNext = (s->errMode == 1); break;
            case Stmt::SubDecl: case Stmt::FuncDecl: case Stmt::ClassDecl:
            case Stmt::PropDecl: break;   // handled at registration
            case Stmt::DimS: {
                for (auto& d : s->decls) {
                    if (d.second.empty()) { declareVar(d.first); }
                    else if (d.second.size() == 1 && !d.second[0]) {   // Dim a()  -> dynamic
                        auto ad = std::make_shared<ArrayData>(); ad->uppers = { -1 };
                        putVar(d.first, Value::array(ad));
                    } else {
                        auto ad = std::make_shared<ArrayData>();
                        for (auto& de : d.second) ad->uppers.push_back((int)eval(de.get()).toI64());
                        ad->data.assign(ad->total(), Value::empty());
                        putVar(d.first, Value::array(ad));
                    }
                }
                break;
            }
            case Stmt::ReDimS: {
                std::vector<int> up; for (auto& de : s->dims) up.push_back((int)eval(de.get()).toI64());
                std::string low = toLower(s->name);
                Value* vp = varPtr(low);
                if (vp && vp->isArray()) vp->arr->redim(up, s->preserve);
                else { auto ad = std::make_shared<ArrayData>(); ad->redim(up, false); setVarValue(s->name, Value::array(ad)); }
                break;
            }
            case Stmt::ConstS:
                for (auto& c : s->consts) putVar(c.first, cloneIfArray(eval(c.second.get())));
                break;
            case Stmt::Assign:    assignTo(s->target.get(), cloneIfArray(eval(s->value.get())), false); break;
            case Stmt::SetAssign: assignTo(s->target.get(), eval(s->value.get()), true); break;
            case Stmt::CallS:     execCall(s); break;
            case Stmt::ExitS: {
                const std::string& w = s->name;
                if (w == "for") throw ExitSignal{ ExitSignal::For };
                if (w == "do")  throw ExitSignal{ ExitSignal::Do };
                if (w == "function") throw ExitSignal{ ExitSignal::Function };
                if (w == "property") throw ExitSignal{ ExitSignal::Property };
                throw ExitSignal{ ExitSignal::Sub };
            }
            case Stmt::IfS: {
                if (truth(eval(s->cond.get()))) { execBlock(s->body); break; }
                bool done = false;
                for (auto& ei : s->elifs) if (truth(eval(ei.first.get()))) { execBlock(ei.second); done = true; break; }
                if (!done) execBlock(s->elseBody);
                break;
            }
            case Stmt::WhileS:
                while (truth(eval(s->cond.get()))) execBlock(s->body);
                break;
            case Stmt::DoLoopS: execDo(s); break;
            case Stmt::ForS: execFor(s); break;
            case Stmt::ForEachS: execForEach(s); break;
            case Stmt::SelectS: execSelect(s); break;
            case Stmt::WithS: {
                withStack.push_back(eval(s->coll.get()));
                try { execBlock(s->body); } catch (...) { withStack.pop_back(); throw; }
                withStack.pop_back();
                break;
            }
            case Stmt::EraseS: {
                for (auto& tgt : s->callArgs) {
                    auto ad = std::make_shared<ArrayData>(); ad->uppers = { -1 };
                    assignTo(tgt.get(), Value::array(ad), false);
                }
                break;
            }
            case Stmt::MidS: {
                // Mid(strVar, start[, length]) = value  — replace in place, never changing length.
                std::string str = eval(s->target.get()).toStr();
                long start = eval(s->callArgs[0].get()).toI64();          // 1-based
                std::string repl = eval(s->value.get()).toStr();
                if (start < 1 || start > (long)str.size())
                    raiseErr(5, "Invalid procedure call or argument");
                long pos = start - 1;
                long avail = (long)str.size() - pos;                      // chars we may overwrite
                long n = (long)repl.size();
                if (s->callArgs.size() == 2) { long len = eval(s->callArgs[1].get()).toI64(); if (len < n) n = len; }
                if (n > avail) n = avail;
                for (long k = 0; k < n; ++k) str[pos + k] = repl[k];
                assignTo(s->target.get(), Value::str(str), false);
                break;
            }
        }
    }
    static Value forStep(double v) {
        if (std::floor(v) == v && v >= INT32_MIN && v <= INT32_MAX) return Value::integer((int32_t)v);
        return Value::number(v);
    }
    void execFor(Stmt* s) {
        double from = eval(s->fromE.get()).toDouble();
        double to   = eval(s->toE.get()).toDouble();
        double step = s->stepE ? eval(s->stepE.get()).toDouble() : 1.0;
        setVarValueL(s->lname, forStep(from));
        Value* slot = varPtr(s->lname);   // element address is stable across the loop (unordered_map guarantee)
        while (true) {
            double v = slot->toDouble();
            if (step >= 0 ? v > to : v < to) break;
            try { execBlock(s->body); }
            catch (ExitSignal& sig) { if (sig.kind == ExitSignal::For) break; throw; }
            *slot = forStep(slot->toDouble() + step);
        }
    }
    void execForEach(Stmt* s) {
        Value coll = eval(s->coll.get());
        std::vector<Value> items;
        if (coll.isArray()) items = coll.arr->data;
        else if (coll.isObject() && coll.obj) { if (!coll.obj->enumerate(*this, items)) raiseErr(451, "Object not a collection"); }
        else if (coll.isEmpty()) return;
        else raiseErr(451, "Object not a collection");
        if (items.empty()) return;
        setVarValueL(s->lname, Value::empty());
        Value* slot = varPtr(s->lname);
        for (auto& it : items) {
            *slot = it;
            try { execBlock(s->body); }
            catch (ExitSignal& sig) { if (sig.kind == ExitSignal::For) break; throw; }
        }
    }
    void execDo(Stmt* s) {
        auto test = [&]() { bool t = truth(eval(s->cond.get())); return s->isUntil ? !t : t; };
        if (s->testPos == 1) {
            while (test()) { try { execBlock(s->body); } catch (ExitSignal& g) { if (g.kind == ExitSignal::Do) break; throw; } }
        } else if (s->testPos == 2) {
            do { try { execBlock(s->body); } catch (ExitSignal& g) { if (g.kind == ExitSignal::Do) break; throw; } } while (test());
        } else {
            while (true) { try { execBlock(s->body); } catch (ExitSignal& g) { if (g.kind == ExitSignal::Do) break; throw; } }
        }
    }
    void execSelect(Stmt* s) {
        Value v = eval(s->coll.get());
        int elseIdx = -1;
        for (size_t ci = 0; ci < s->cases.size(); ++ci) {
            auto& cc = s->cases[ci];
            if (cc.tests.empty()) { elseIdx = (int)ci; continue; }
            for (auto& t : cc.tests) {
                bool match = false;
                if (t.k == CaseTest::Val)      match = cmp(v, eval(t.e1.get())) == 0;
                else if (t.k == CaseTest::Range){ Value lo = eval(t.e1.get()), hi = eval(t.e2.get()); match = cmp(v, lo) >= 0 && cmp(v, hi) <= 0; }
                else { Value r = eval(t.e1.get()); int c = cmp(v, r);
                       const std::string& o = t.op;
                       match = (o == "=" && c == 0) || (o == "<>" && c != 0) || (o == "<" && c < 0) ||
                               (o == ">" && c > 0) || (o == "<=" && c <= 0) || (o == ">=" && c >= 0); }
                if (match) { execBlock(cc.body); return; }
            }
        }
        if (elseIdx >= 0) execBlock(s->cases[elseIdx].body);
    }

    void execCall(Stmt* s) {
        Expr* callee; std::vector<ExprP>* argExprs; std::vector<ExprP> empty;
        if (!s->callArgs.empty()) { callee = s->callee.get(); argExprs = &s->callArgs; }
        else if (s->callee->k == Expr::Index) { callee = s->callee->a.get(); argExprs = &s->callee->args; }
        else { callee = s->callee.get(); argExprs = &empty; }

        if (callee->k == Expr::Var) {
            const std::string& low = callee->lname;
            if (procs.count(low)) { auto a = evalArgs(*argExprs); callProcV(procs[low], *argExprs, a, nullptr); return; }
            if (!instanceStack.empty()) { StmtP d = classCallable(instanceStack.back(), low); if (d) { auto a = evalArgs(*argExprs); callClassMethod(instanceStack.back(), d, a); return; } }
            if (builtins.count(low)) { auto a = evalArgs(*argExprs); builtins[low](*this, a); return; }
            Value* vp = varPtr(low);
            if (vp && vp->isObject() && vp->obj) { auto a = evalArgs(*argExprs); vp->obj->get(*this, "", a); return; }
            raiseErr(13, "Type mismatch: '" + callee->name + "'");
        }
        if (callee->k == Expr::Member) {
            Value obj = callee->a ? eval(callee->a.get()) : withTarget();
            if (obj.isNothing()) raiseErr(91, "Object variable not set");
            if (obj.isObject() && obj.obj) { auto a = evalArgs(*argExprs); obj.obj->get(*this, callee->name, a); return; }
            raiseErr(424, "Object required");
        }
        eval(s->callee.get());   // fallback: just evaluate
    }

    // ---------------------------------------------------------- assignment
    void assignTo(Expr* lv, const Value& val, bool isSet) {
        if (lv->k == Expr::Var) { setVarValueL(lv->lname, val); return; }
        if (lv->k == Expr::Index) {
            Expr* t = lv->a.get();
            auto idx = evalArgs(lv->args);
            if (t->k == Expr::Var) {
                std::string low = toLower(t->name);
                Value* vp = varPtr(low);
                if (vp && vp->isArray())  { arraySet(*vp, idx, val); return; }
                if (vp && vp->isObject() && vp->obj) { vp->obj->set(*this, "", idx, val, isSet); return; }
                raiseErr(9, "Subscript out of range / not an array: " + t->name);
            }
            if (t->k == Expr::Member) {
                Value obj = t->a ? eval(t->a.get()) : withTarget();
                if (obj.isObject() && obj.obj) { obj.obj->set(*this, t->name, idx, val, isSet); return; }
                raiseErr(424, "Object required");
            }
            Value tv = eval(lv->a.get());
            if (tv.isArray()) { arraySet(tv, idx, val); return; }
            if (tv.isObject() && tv.obj) { tv.obj->set(*this, "", idx, val, isSet); return; }
            raiseErr(13, "Type mismatch");
        }
        if (lv->k == Expr::Member) {
            Value obj = lv->a ? eval(lv->a.get()) : withTarget();
            if (obj.isNothing()) raiseErr(91, "Object variable not set");
            if (obj.isObject() && obj.obj) { std::vector<Value> none; obj.obj->set(*this, lv->name, none, val, isSet); return; }
            raiseErr(424, "Object required");
        }
        raiseErr(1, "Cannot assign to this expression");
    }

    // ---------------------------------------------------------- proc calls
    Value invoke(StmtP decl, std::vector<Value>& argv, ClassInstance* self,
                 std::unordered_map<std::string, Value>* outFrame) {
        // Guard against runaway recursion: raise a catchable error instead of
        // overflowing the C++ stack and crashing the process. (frames.size() is
        // the current call depth.)
        if (frames.size() >= 2000) raiseErr(28, "Out of stack space");
        // recycle a frame map from the pool (keeps its bucket allocation) instead of allocating anew
        if (!framePool.empty()) { frames.push_back(std::move(framePool.back())); framePool.pop_back(); frames.back().clear(); }
        else frames.push_back({});
        size_t fi = frames.size() - 1;
        frames[fi].reserve(decl->params.size() + 2);
        for (size_t i = 0; i < decl->params.size(); ++i) {
            const std::string& pn = decl->params[i].lname;
            if (decl->params[i].byVal && i < argv.size() && argv[i].isArray())
                frames[fi][pn] = cloneIfArray(argv[i]);
            else
                frames[fi][pn] = i < argv.size() ? argv[i] : Value::empty();
        }
        bool retByName = decl->isFunction || (decl->k == Stmt::PropDecl && decl->propKind == 0);
        if (retByName) frames[fi][decl->lname] = Value::empty();
        if (self) instanceStack.push_back(self);
        bool saved = onErrorResumeNext; onErrorResumeNext = false;
        try { execBlock(decl->body); }
        catch (ExitSignal&) { }
        catch (...) {
            onErrorResumeNext = saved; if (self) instanceStack.pop_back();
            framePool.push_back(std::move(frames.back())); frames.pop_back(); throw;
        }
        Value ret = retByName ? std::move(frames[fi][decl->lname]) : Value::empty();
        if (outFrame) *outFrame = std::move(frames[fi]);   // only when ByRef writeback is needed
        onErrorResumeNext = saved; if (self) instanceStack.pop_back();
        // Pop the frame OFF the stack before destroying its locals: destroying an
        // object may run Class_Terminate, which can call functions that push onto
        // `frames` (and reallocate it). Clearing a detached map keeps that safe.
        std::unordered_map<std::string, Value> done = std::move(frames.back());
        frames.pop_back();
        done.clear();                                      // fires Class_Terminate for locals now
        framePool.push_back(std::move(done));              // empty map (keeps buckets) back to the pool
        return ret;
    }
    Value callProcV(StmtP decl, std::vector<ExprP>& argExprs, std::vector<Value>& argv, ClassInstance* self) {
        // Only capture the frame for ByRef write-back if some argument is actually an
        // lvalue (Var/Index/Member). For expression/literal args (the common case) this
        // avoids copying the whole frame map on every call.
        size_t np = decl->params.size();
        bool needWriteback = false;
        for (size_t i = 0; i < np && i < argExprs.size(); ++i) {
            if (decl->params[i].byVal) continue;
            Expr::K k = argExprs[i]->k;
            if (k == Expr::Var || k == Expr::Index || k == Expr::Member) { needWriteback = true; break; }
        }
        if (!needWriteback) return invoke(decl, argv, self, nullptr);

        std::unordered_map<std::string, Value> outFrame;
        Value ret = invoke(decl, argv, self, &outFrame);
        for (size_t i = 0; i < np && i < argExprs.size(); ++i) {
            if (decl->params[i].byVal) continue;
            Expr* ax = argExprs[i].get();
            if (ax->k == Expr::Var || ax->k == Expr::Index || ax->k == Expr::Member) {
                auto it = outFrame.find(decl->params[i].lname);
                if (it != outFrame.end()) assignTo(ax, it->second, false);
            }
        }
        return ret;
    }
    Value callProc(StmtP decl, std::vector<ExprP>& argExprs, ClassInstance* self) {
        auto argv = evalArgs(argExprs);
        return callProcV(decl, argExprs, argv, self);
    }

    // ---------------------------------------------------------- class reg
    void registerClass(StmtP c) {
        auto ci = std::make_shared<ClassInfo>();
        ci->name = c->name;
        for (auto& m : c->members) {
            if (m->k == Stmt::DimS) { for (auto& d : m->decls) ci->fields.push_back(d.first); }
            else if (m->k == Stmt::SubDecl || m->k == Stmt::FuncDecl) ci->methods[toLower(m->name)] = m;
            else if (m->k == Stmt::PropDecl) {
                if (m->propKind == 0) ci->propGet[toLower(m->name)] = m;
                else if (m->propKind == 1) ci->propLet[toLower(m->name)] = m;
                else ci->propSet[toLower(m->name)] = m;
            }
            if (m->isDefault) ci->defaultMember = toLower(m->name);
        }
        classes[toLower(c->name)] = ci;
    }

    // defined out-of-line below
    void setupConstants();
    void setupObjects();
    void registerBuiltins();

    // expose to builtins that need object creation
public:
    Value createObjectByName(const std::string& progid) {
        // CreateObject is reserved for real COM/ActiveX. Native objects use New.
        bool ok = false;
        ObjectPtr o = createComObject(progid, ok);
        if (ok) return Value::object(o);
        raiseErr(429, "ActiveX component can't create object: '" + progid +
                      "'  (CreateObject makes real COM calls, available on Windows only; "
                      "for engine objects use New, e.g. Set d = New Dictionary)");
    }
    std::shared_ptr<ErrObject> err() { return errObj; }
};

// ---- ClassInstance member access (needs full Interpreter) ----
// ---- function reference (GetRef) ----
struct FuncRefObject : IObject {
    std::string name;   // lowercased proc/builtin name
    explicit FuncRefObject(std::string n) : name(std::move(n)) {}
    std::string typeName() const override { return "Function"; }
    Value get(Interpreter& in, const std::string& member, std::vector<Value>& args) override {
        if (member.empty()) return in.callByName(name, args);   // f(args)
        raiseErr(438, "Function reference has no member '" + member + "'");
    }
};
Value Interpreter::makeFuncRef(const std::string& name) {
    std::string low = toLower(name);
    if (!procs.count(low) && !builtins.count(low))
        raiseErr(5, "GetRef: '" + name + "' is not a defined Sub, Function, or built-in");
    return Value::object(std::make_shared<FuncRefObject>(low));
}

ClassInstance::~ClassInstance() { if (interp) interp->runTerminator(this); }

bool ClassInstance::tryDefault(Interpreter& in, Value& out) {
    if (!cls || cls->defaultMember.empty()) return false;
    std::vector<Value> none;
    out = get(in, std::string(), none);   // "" routes to the default member
    return true;
}

Value ClassInstance::get(Interpreter& in, const std::string& name, std::vector<Value>& args) {    std::string low = toLower(name);
    auto& ci = *cls;
    // Empty member name => the class's Default member (for `o(args)` / value coercion).
    if (low.empty() && !ci.defaultMember.empty()) low = ci.defaultMember;
    auto pg = ci.propGet.find(low);
    auto m  = ci.methods.find(low);
    if (pg != ci.propGet.end() || m != ci.methods.end()) {
        StmtP decl = (pg != ci.propGet.end()) ? pg->second : m->second;
        size_t np = decl->params.size();
        if (args.size() > np) {
            // More arguments than the member declares: call it with the args it
            // expects, then apply the leftover args to the RESULT (VBScript treats
            // `obj.Prop(i)` on a no-arg getter as indexing the returned value).
            std::vector<Value> head(args.begin(), args.begin() + np);
            Value r = in.callClassMethod(this, decl, head);
            std::vector<Value> tail(args.begin() + np, args.end());
            if (r.isArray())                    return in.arrayGetPub(r, tail);
            if (r.isObject() && r.obj)          return r.obj->get(in, "", tail);
            raiseErr(13, "Type mismatch");
        }
        return in.callClassMethod(this, decl, args);
    }
    auto f  = fields.find(low);
    if (f != fields.end()) {
        if (args.empty()) return f->second;
        if (f->second.isArray())  return in.arrayGetPub(f->second, args);
        if (f->second.isObject() && f->second.obj) return f->second.obj->get(in, "", args);
        raiseErr(13, "Type mismatch");
    }
    raiseErr(438, "Object doesn't support this property or method: " + name);
}
void ClassInstance::set(Interpreter& in, const std::string& name, std::vector<Value>& args,
                        const Value& val, bool isSet) {
    std::string low = toLower(name);
    auto& ci = *cls;
    StmtP prop;
    if (isSet) { auto s = ci.propSet.find(low); if (s != ci.propSet.end()) prop = s->second;
                 else { auto l = ci.propLet.find(low); if (l != ci.propLet.end()) prop = l->second; } }
    else       { auto l = ci.propLet.find(low); if (l != ci.propLet.end()) prop = l->second;
                 else { auto s = ci.propSet.find(low); if (s != ci.propSet.end()) prop = s->second; } }
    if (prop) { std::vector<Value> argv = args; argv.push_back(val); in.callClassMethod(this, prop, argv); return; }
    auto f = fields.find(low);
    if (f != fields.end()) {
        if (args.empty()) { f->second = val; return; }
        if (f->second.isArray()) { in.arraySetPub(f->second, args, val); return; }
        if (f->second.isObject() && f->second.obj) { f->second.obj->set(in, "", args, val, isSet); return; }
        raiseErr(13, "Type mismatch");
    }
    fields[low] = val;   // lenient: allow late field creation
}

// ---- constants ----
void Interpreter::setupConstants() {
    auto C = [&](const char* n, Value v) { globals[toLower(n)] = v; };
    C("vbCr", Value::str("\r")); C("vbLf", Value::str("\n")); C("vbCrLf", Value::str("\r\n"));
    C("vbNewLine", Value::str("\r\n")); C("vbTab", Value::str("\t")); C("vbNullString", Value::str(""));
    C("vbNullChar", Value::str(std::string(1, '\0'))); C("vbBack", Value::str("\b"));
    C("vbFormFeed", Value::str("\f")); C("vbVerticalTab", Value::str("\v"));
    C("vbTrue", Value::integer(-1)); C("vbFalse", Value::integer(0));
    C("vbBinaryCompare", Value::integer(0)); C("vbTextCompare", Value::integer(1));
    // VbVarType + MsgBox + weekday + FormatDateTime constants
    const std::pair<const char*, int> ints[] = {
        {"vbEmpty",0},{"vbNull",1},{"vbInteger",2},{"vbLong",3},{"vbSingle",4},{"vbDouble",5},
        {"vbCurrency",6},{"vbDate",7},{"vbString",8},{"vbObject",9},{"vbError",10},{"vbBoolean",11},
        {"vbVariant",12},{"vbByte",17},{"vbArray",8192},
        {"vbOKOnly",0},{"vbOKCancel",1},{"vbAbortRetryIgnore",2},{"vbYesNoCancel",3},{"vbYesNo",4},
        {"vbRetryCancel",5},{"vbCritical",16},{"vbQuestion",32},{"vbExclamation",48},{"vbInformation",64},
        {"vbOK",1},{"vbCancel",2},{"vbAbort",3},{"vbRetry",4},{"vbIgnore",5},{"vbYes",6},{"vbNo",7},
        {"vbSunday",1},{"vbMonday",2},{"vbTuesday",3},{"vbWednesday",4},{"vbThursday",5},
        {"vbFriday",6},{"vbSaturday",7},{"vbUseSystemDayOfWeek",0},{"vbFirstJan1",1},
        {"vbGeneralDate",0},{"vbLongDate",1},{"vbShortDate",2},{"vbLongTime",3},{"vbShortTime",4},
    };
    for (auto& kv : ints) C(kv.first, Value::integer(kv.second));
    C("vbObjectError", Value::number((double)0x80040000u - 4294967296.0)); // -2147221504
}
void Interpreter::setupObjects() {
    errObj = std::make_shared<ErrObject>();
    globals["err"] = Value::object(errObj);
    directiveObj = std::make_shared<DirectiveObject>();
    globals["directive"] = Value::object(directiveObj);
}

// ---- small string helpers for builtins ----
static std::string upper(std::string s){ for(char&c:s)c=(char)std::toupper((unsigned char)c); return s; }
static std::string lower(std::string s){ for(char&c:s)c=(char)std::tolower((unsigned char)c); return s; }
static bool parseDateToSerial(const std::string& in, double& out) {
    std::string s = in; size_t a = s.find_first_not_of(" \t"); if (a==std::string::npos) return false;
    size_t z = s.find_last_not_of(" \t"); s = s.substr(a, z-a+1);
    int y,mo,d,h=0,mi=0,se=0;
    if (std::sscanf(s.c_str(), "%d-%d-%d %d:%d:%d", &y,&mo,&d,&h,&mi,&se) >= 3) { out = serialFromYMDHMS(y,mo,d,h,mi,se); return true; }
    if (std::sscanf(s.c_str(), "%d/%d/%d %d:%d:%d", &mo,&d,&y,&h,&mi,&se) >= 3) { out = serialFromYMDHMS(y,mo,d,h,mi,se); return true; }
    if (std::sscanf(s.c_str(), "%d:%d:%d", &h,&mi,&se) >= 2) { out = serialFromYMDHMS(1899,12,30,h,mi,se); return true; }
    return false;
}

// ---- builtins ----
static const char* kMonths[12] = {"January","February","March","April","May","June",
                                  "July","August","September","October","November","December"};
static const char* kDays[7] = {"Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"};
static std::string fmtGrouped(double v, int digits, bool group) {
    char buf[64]; std::snprintf(buf, sizeof(buf), "%.*f", digits < 0 ? 2 : digits, v);
    std::string s = buf;
    bool neg = !s.empty() && s[0] == '-'; if (neg) s = s.substr(1);
    size_t dot = s.find('.');
    std::string ip = (dot == std::string::npos) ? s : s.substr(0, dot);
    std::string fp = (dot == std::string::npos) ? "" : s.substr(dot);
    if (group) {
        std::string g; int c = 0;
        for (int i = (int)ip.size() - 1; i >= 0; --i) { g.push_back(ip[i]); if (++c % 3 == 0 && i > 0) g.push_back(','); }
        std::reverse(g.begin(), g.end()); ip = g;
    }
    return (neg ? "-" : "") + ip + fp;
}

void Interpreter::registerBuiltins() {
    auto& B = builtins;
    auto S  = [](std::vector<Value>& a, size_t i)->std::string { return i < a.size() ? a[i].toStr() : std::string(); };
    auto D  = [](std::vector<Value>& a, size_t i)->double { return i < a.size() ? a[i].toDouble() : 0.0; };
    auto N  = [](std::vector<Value>& a, size_t i)->long long { return i < a.size() ? a[i].toI64() : 0; };
    auto has = [](std::vector<Value>& a, size_t i){ return i < a.size() && !a[i].isEmpty(); };

    // ---- strings ----
    B["len"]    = [=](Interpreter&, std::vector<Value>& a){ if(!a.empty()&&a[0].isNull()) return Value::null(); return Value::integer((int32_t)S(a,0).size()); };
    B["left"]   = [=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); long n=(long)N(a,1); if(n<0)n=0; return Value::str(s.substr(0,std::min((size_t)n,s.size()))); };
    B["right"]  = [=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); long n=(long)N(a,1); if(n<0)n=0; if((size_t)n>=s.size()) return Value::str(s); return Value::str(s.substr(s.size()-n)); };
    B["mid"]    = [=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); long st=(long)N(a,1); if(st<1)st=1; if((size_t)st>s.size()) return Value::str(""); size_t off=st-1; size_t len = has(a,2)? (size_t)std::max(0LL,(long long)N(a,2)) : s.size()-off; return Value::str(s.substr(off, len)); };
    B["ucase"]  = [=](Interpreter&, std::vector<Value>& a){ return Value::str(upper(S(a,0))); };
    B["lcase"]  = [=](Interpreter&, std::vector<Value>& a){ return Value::str(lower(S(a,0))); };
    B["trim"]   = [=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); size_t b=s.find_first_not_of(" "); if(b==std::string::npos)return Value::str(""); size_t e=s.find_last_not_of(" "); return Value::str(s.substr(b,e-b+1)); };
    B["ltrim"]  = [=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); size_t b=s.find_first_not_of(" "); return Value::str(b==std::string::npos?"":s.substr(b)); };
    B["rtrim"]  = [=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); size_t e=s.find_last_not_of(" "); return Value::str(e==std::string::npos?"":s.substr(0,e+1)); };
    B["space"]  = [=](Interpreter&, std::vector<Value>& a){ long n=(long)N(a,0); if(n<0)n=0; return Value::str(std::string(n,' ')); };
    B["string"] = [=](Interpreter&, std::vector<Value>& a){ long n=(long)N(a,0); if(n<0)n=0; char c; if(a.size()>1&&a[1].type==Value::Type::String&&!a[1].s.empty()) c=a[1].s[0]; else c=(char)N(a,1); return Value::str(std::string(n,c)); };
    B["strreverse"]=[=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); std::reverse(s.begin(),s.end()); return Value::str(s); };
    B["instr"]  = [=](Interpreter&, std::vector<Value>& a){
        long start=1; std::string s1,s2;
        if(a.size()>=3 && a[0].looksNumeric()){ start=(long)a[0].toI64(); s1=a[1].toStr(); s2=a[2].toStr(); }
        else { s1=S(a,0); s2=S(a,1); }
        if(start<1){start=1;}
        if((size_t)(start-1)>s1.size()) return Value::integer(0);
        size_t pos=s1.find(s2,start-1); return Value::integer(pos==std::string::npos?0:(int32_t)(pos+1)); };
    B["instrrev"]=[=](Interpreter&, std::vector<Value>& a){
        std::string s1=S(a,0),s2=S(a,1); long start=has(a,2)?(long)N(a,2):-1;
        size_t from = start<0? std::string::npos : (size_t)(start-1);
        size_t pos=s1.rfind(s2, from); return Value::integer(pos==std::string::npos?0:(int32_t)(pos+1)); };
    B["replace"]=[=](Interpreter&, std::vector<Value>& a){
        std::string s=S(a,0),find=S(a,1),rep=S(a,2); if(find.empty()) return Value::str(s);
        long start = has(a,3)?(long)N(a,3):1; long count = has(a,4)?(long)N(a,4):-1;
        if(start<1){start=1;}
        std::string head = s.substr(0, std::min((size_t)start-1, s.size()));
        std::string body = s.substr(std::min((size_t)start-1, s.size()));
        std::string out; size_t pos=0; long done=0;
        while(true){ size_t f=body.find(find,pos); if(f==std::string::npos||(count>=0&&done>=count)){ out+=body.substr(pos); break;} out+=body.substr(pos,f-pos)+rep; pos=f+find.size(); done++; }
        return Value::str(head+out); };
    B["split"]  = [=](Interpreter&, std::vector<Value>& a){
        std::string s=S(a,0); std::string delim = has(a,1)?S(a,1):" ";
        long count = has(a,2) ? (long)N(a,2) : -1;      // max number of substrings (-1 = all)
        int cmp = has(a,3) ? (int)N(a,3) : 0;            // 0 binary, 1 text (case-insensitive)
        std::vector<Value> parts;
        if(count==0) return makeArray(parts);            // count 0 -> empty array
        if(delim.empty()){ parts.push_back(Value::str(s)); return makeArray(parts); }
        std::string hay = (cmp==1) ? lower(s) : s;
        std::string ndl = (cmp==1) ? lower(delim) : delim;
        size_t pos=0,f;
        while((f=hay.find(ndl,pos))!=std::string::npos){
            if(count>0 && (long)parts.size()==count-1) break;   // limit reached; remainder is final element
            parts.push_back(Value::str(s.substr(pos,f-pos))); pos=f+delim.size();
        }
        parts.push_back(Value::str(s.substr(pos))); return makeArray(parts); };
    B["join"]   = [=](Interpreter&, std::vector<Value>& a){
        if(a.empty()||!a[0].isArray()) raiseErr(13,"Type mismatch (Join expects an array)");
        std::string delim = has(a,1)?S(a,1):" "; std::string out;
        auto& d=a[0].arr->data; for(size_t k=0;k<d.size();++k){ if(k)out+=delim; out+=d[k].toConcatStr(); } return Value::str(out); };
    B["strcomp"]=[=](Interpreter&, std::vector<Value>& a){ std::string s1=S(a,0),s2=S(a,1); bool ci=(a.size()>2&&a[2].toI64()==1); if(ci){s1=lower(s1);s2=lower(s2);} int c=s1.compare(s2); return Value::integer(c<0?-1:(c>0?1:0)); };
    B["chr"]    = [=](Interpreter&, std::vector<Value>& a){ long n=(long)N(a,0); return Value::str(std::string(1,(char)(n&0xFF))); };
    B["chrw"]   = B["chr"];
    B["asc"]    = [=](Interpreter&, std::vector<Value>& a){ std::string s=S(a,0); if(s.empty()) raiseErr(5,"Invalid procedure call"); return Value::integer((unsigned char)s[0]); };
    B["ascw"]   = B["asc"];

    // ---- conversion ----
    B["cint"]   = [=](Interpreter&, std::vector<Value>& a){ double v=D(a,0); double r=roundHalfEven(v); if(r< -32768||r>32767) raiseErr(6,"Overflow"); return Value::integer((int32_t)r); };
    B["clng"]   = [=](Interpreter&, std::vector<Value>& a){ double v=D(a,0); double r=roundHalfEven(v); if(r<INT32_MIN||r>INT32_MAX) raiseErr(6,"Overflow"); return Value::integer((int32_t)r); };
    B["cbyte"]  = [=](Interpreter&, std::vector<Value>& a){ double r=roundHalfEven(D(a,0)); if(r<0||r>255) raiseErr(6,"Overflow"); return Value::integer((int32_t)r); };
    B["cdbl"]   = [=](Interpreter&, std::vector<Value>& a){ return Value::number(D(a,0)); };
    B["csng"]   = [=](Interpreter&, std::vector<Value>& a){ return Value::number((float)D(a,0)); };
    B["cstr"]   = [=](Interpreter&, std::vector<Value>& a){ return Value::str(S(a,0)); };
    B["cbool"]  = [=](Interpreter&, std::vector<Value>& a){ return Value::boolean(a.empty()?false:a[0].toBool()); };
    B["cdate"]  = [=](Interpreter&, std::vector<Value>& a){ if(a.empty())raiseErr(13,"Type mismatch"); if(a[0].looksNumeric())return Value::number(a[0].toDouble()); double o; if(parseDateToSerial(a[0].toStr(),o))return Value::number(o); raiseErr(13,"Type mismatch"); };
    B["int"]    = [=](Interpreter&, std::vector<Value>& a){ if(!a.empty()&&a[0].type==Value::Type::Int)return a[0]; return Value::number(std::floor(D(a,0))); };
    B["fix"]    = [=](Interpreter&, std::vector<Value>& a){ if(!a.empty()&&a[0].type==Value::Type::Int)return a[0]; return Value::number(std::trunc(D(a,0))); };
    B["round"]  = [=](Interpreter&, std::vector<Value>& a){ double v=D(a,0); int dg=has(a,1)?(int)N(a,1):0; double f=std::pow(10,dg); double r=roundHalfEven(v*f)/f; return Value::number(r); };
    B["abs"]    = [=](Interpreter&, std::vector<Value>& a){ if(!a.empty()&&a[0].isNull())return Value::null(); return Value::number(std::fabs(D(a,0))); };
    B["sgn"]    = [=](Interpreter&, std::vector<Value>& a){ double v=D(a,0); return Value::integer(v>0?1:(v<0?-1:0)); };
    B["sqr"]    = [=](Interpreter&, std::vector<Value>& a){ double v=D(a,0); if(v<0)raiseErr(5,"Invalid procedure call (Sqr of negative)"); return Value::number(std::sqrt(v)); };
    B["hex"]    = [=](Interpreter&, std::vector<Value>& a){ std::ostringstream o; o<<std::uppercase<<std::hex<<(uint32_t)(int32_t)N(a,0); return Value::str(o.str()); };
    B["oct"]    = [=](Interpreter&, std::vector<Value>& a){ std::ostringstream o; o<<std::oct<<(uint32_t)(int32_t)N(a,0); return Value::str(o.str()); };

    // ---- type info ----
    B["isnumeric"]=[=](Interpreter&, std::vector<Value>& a){ return Value::boolean(!a.empty()&&a[0].looksNumeric()); };
    B["isempty"] =[=](Interpreter&, std::vector<Value>& a){ return Value::boolean(!a.empty()&&a[0].isEmpty()); };
    B["isnull"]  =[=](Interpreter&, std::vector<Value>& a){ return Value::boolean(!a.empty()&&a[0].isNull()); };
    B["isarray"] =[=](Interpreter&, std::vector<Value>& a){ return Value::boolean(!a.empty()&&a[0].isArray()); };
    B["isobject"]=[=](Interpreter&, std::vector<Value>& a){ return Value::boolean(!a.empty()&&a[0].isObject()); };
    B["isdate"]  =[=](Interpreter&, std::vector<Value>& a){ if(a.empty()||a[0].type!=Value::Type::String)return Value::boolean(false); double o; return Value::boolean(parseDateToSerial(a[0].s,o)); };
    B["vartype"] =[=](Interpreter&, std::vector<Value>& a){ return Value::integer(a.empty()?0:a[0].varType()); };
    B["typename"]=[=](Interpreter&, std::vector<Value>& a){ return Value::str(a.empty()?"Empty":a[0].vbTypeName()); };

    // ---- arrays ----
    B["array"]  = [=](Interpreter&, std::vector<Value>& a){ return makeArray(a); };
    B["ubound"] = [=](Interpreter&, std::vector<Value>& a){ if(a.empty()||!a[0].isArray())raiseErr(13,"Type mismatch"); int dim=has(a,1)?(int)N(a,1):1; if(dim<1||dim>(int)a[0].arr->uppers.size())raiseErr(9,"Subscript out of range"); return Value::integer(a[0].arr->uppers[dim-1]); };
    B["lbound"] = [=](Interpreter&, std::vector<Value>& a){ if(a.empty()||!a[0].isArray())raiseErr(13,"Type mismatch"); return Value::integer(0); };
    B["filter"] = [=](Interpreter&, std::vector<Value>& a){ if(a.empty()||!a[0].isArray())raiseErr(13,"Type mismatch"); std::string m=S(a,1); bool inc=has(a,2)?a[2].toBool():true; std::vector<Value> out; for(auto&e:a[0].arr->data){ bool found=e.toStr().find(m)!=std::string::npos; if(found==inc)out.push_back(e);} return makeArray(out); };

    // ---- math ----
    B["rnd"]    = [=](Interpreter&, std::vector<Value>&){ return Value::number((double)std::rand()/((double)RAND_MAX+1.0)); };
    B["randomize"]=[=](Interpreter&, std::vector<Value>& a){ std::srand(a.empty()?(unsigned)std::time(nullptr):(unsigned)N(a,0)); return Value::empty(); };
    B["sin"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(std::sin(D(a,0))); };
    B["cos"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(std::cos(D(a,0))); };
    B["tan"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(std::tan(D(a,0))); };
    B["atn"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(std::atan(D(a,0))); };
    B["log"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(std::log(D(a,0))); };
    B["exp"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(std::exp(D(a,0))); };

    // ---- date / time ----
    auto nowSerial=[](){ std::time_t tt=std::time(nullptr); std::tm lt=*std::localtime(&tt); return serialFromYMDHMS(lt.tm_year+1900,lt.tm_mon+1,lt.tm_mday,lt.tm_hour,lt.tm_min,lt.tm_sec); };
    B["now"]  = [=](Interpreter&, std::vector<Value>&){ return Value::number(nowSerial()); };
    B["date"] = [=](Interpreter&, std::vector<Value>&){ return Value::number(std::floor(nowSerial())); };
    B["time"] = [=](Interpreter&, std::vector<Value>&){ double s=nowSerial(); return Value::number(s-std::floor(s)); };
    B["timer"]= [=](Interpreter&, std::vector<Value>&){ double s=nowSerial(); return Value::number((s-std::floor(s))*86400.0); };
    B["year"]  =[=](Interpreter&, std::vector<Value>& a){ int y,mo,d,h,mi,se; ymdhmsFromSerial(D(a,0),y,mo,d,h,mi,se); return Value::integer(y); };
    B["month"] =[=](Interpreter&, std::vector<Value>& a){ int y,mo,d,h,mi,se; ymdhmsFromSerial(D(a,0),y,mo,d,h,mi,se); return Value::integer(mo); };
    B["day"]   =[=](Interpreter&, std::vector<Value>& a){ int y,mo,d,h,mi,se; ymdhmsFromSerial(D(a,0),y,mo,d,h,mi,se); return Value::integer(d); };
    B["hour"]  =[=](Interpreter&, std::vector<Value>& a){ int y,mo,d,h,mi,se; ymdhmsFromSerial(D(a,0),y,mo,d,h,mi,se); return Value::integer(h); };
    B["minute"]=[=](Interpreter&, std::vector<Value>& a){ int y,mo,d,h,mi,se; ymdhmsFromSerial(D(a,0),y,mo,d,h,mi,se); return Value::integer(mi); };
    B["second"]=[=](Interpreter&, std::vector<Value>& a){ int y,mo,d,h,mi,se; ymdhmsFromSerial(D(a,0),y,mo,d,h,mi,se); return Value::integer(se); };
    B["weekday"]=[=](Interpreter&, std::vector<Value>& a){ long dd=(long)std::floor(D(a,0))+kBaseDay; int dow=(int)(((dd%7)+7+4)%7); return Value::integer(dow+1); };
    B["dateserial"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(serialFromYMDHMS((int)N(a,0),(int)N(a,1),(int)N(a,2),0,0,0)); };
    B["timeserial"]=[=](Interpreter&, std::vector<Value>& a){ return Value::number(serialFromYMDHMS(1899,12,30,(int)N(a,0),(int)N(a,1),(int)N(a,2))); };
    B["datevalue"]=[=](Interpreter&, std::vector<Value>& a){ double o; if(a.empty())raiseErr(13,"Type mismatch"); if(a[0].looksNumeric())o=a[0].toDouble(); else if(!parseDateToSerial(a[0].toStr(),o))raiseErr(13,"Type mismatch"); return Value::number(std::floor(o)); };   // date part only
    B["timevalue"]=[=](Interpreter&, std::vector<Value>& a){ double o; if(a.empty())raiseErr(13,"Type mismatch"); if(a[0].looksNumeric())o=a[0].toDouble(); else if(!parseDateToSerial(a[0].toStr(),o))raiseErr(13,"Type mismatch"); double f=o-std::floor(o); if(f<0)f+=1.0; return Value::number(f); };   // time-of-day fraction
    B["dateadd"]=[=](Interpreter&, std::vector<Value>& a){ std::string iv=lower(S(a,0)); double num=D(a,1); double ser=D(a,2); int y,mo,d,h,mi,se; ymdhmsFromSerial(ser,y,mo,d,h,mi,se);
        if(iv=="yyyy")y+=(int)num; else if(iv=="m"){int t=(mo-1)+(int)num; y+=t/12; mo=t%12+1; if(mo<1){mo+=12;y--;}} else if(iv=="d")ser+=num; else if(iv=="ww")ser+=num*7; else if(iv=="h")ser+=num/24.0; else if(iv=="n")ser+=num/1440.0; else if(iv=="s")ser+=num/86400.0; else raiseErr(5,"Invalid interval");
        if(iv=="yyyy"||iv=="m"){ return Value::number(serialFromYMDHMS(y,mo,d,h,mi,se)); }
        return Value::number(ser); };
    B["datediff"]=[=](Interpreter&, std::vector<Value>& a){ std::string iv=lower(S(a,0)); double s1=D(a,1),s2=D(a,2); double diff=s2-s1;
        if(iv=="d") return Value::integer((int32_t)std::floor(diff));
        if(iv=="ww") return Value::integer((int32_t)std::floor(diff/7));
        if(iv=="h") return Value::integer((int32_t)std::floor(diff*24));
        if(iv=="n") return Value::integer((int32_t)std::floor(diff*1440));
        if(iv=="s") return Value::integer((int32_t)std::floor(diff*86400));
        int y1,y2,mo1,mo2,d,h,mi,se; ymdhmsFromSerial(s1,y1,mo1,d,h,mi,se); ymdhmsFromSerial(s2,y2,mo2,d,h,mi,se); if(iv=="yyyy")return Value::integer(y2-y1); if(iv=="m")return Value::integer((y2-y1)*12+(mo2-mo1)); raiseErr(5,"Invalid interval"); };
    B["formatdatetime"]=[=](Interpreter&, std::vector<Value>& a){ int y,mo,d,h,mi,se; ymdhmsFromSerial(D(a,0),y,mo,d,h,mi,se); char buf[64]; int fmt=has(a,1)?(int)N(a,1):0;
        if(fmt==2){ std::snprintf(buf,64,"%d/%d/%04d",mo,d,y);} else if(fmt==3||fmt==4){ std::snprintf(buf,64,"%02d:%02d:%02d",h,mi,se);} else { std::snprintf(buf,64,"%d/%d/%04d %02d:%02d:%02d",mo,d,y,h,mi,se);} return Value::str(buf); };

    // ---- misc / IO ----
    B["msgbox"]  = [=](Interpreter&, std::vector<Value>& a){ long flags = has(a,1)?(long)N(a,1):0; std::string title = has(a,2)?S(a,2):""; return Value::integer((int32_t)ui::msgBox(S(a,0), flags, title)); };
    B["inputbox"]= [=](Interpreter&, std::vector<Value>& a){ std::string title = has(a,1)?S(a,1):""; std::string def = has(a,2)?S(a,2):""; std::string out; if(ui::inputBox(S(a,0), title, def, out)) return Value::str(out); return Value::str(""); };
    B["createobject"]=[=](Interpreter& I, std::vector<Value>& a){ return I.createObjectByName(S(a,0)); };
    B["getobject"]=[=](Interpreter&, std::vector<Value>& a){ raiseErr(429,"GetObject is not supported in this build: "+S(a,0)); return Value::empty(); };
    B["execute"]       = [=](Interpreter& I, std::vector<Value>& a){ I.executeString(S(a,0), false); return Value::empty(); };
    B["executeglobal"] = [=](Interpreter& I, std::vector<Value>& a){ I.executeString(S(a,0), true);  return Value::empty(); };
    B["rgb"]     = [=](Interpreter&, std::vector<Value>& a){ long r=N(a,0)&0xFF,g=N(a,1)&0xFF,b=N(a,2)&0xFF; return Value::integer((int32_t)(r | (g<<8) | (b<<16))); };

    // ---- added for VBScript parity ----
    B["eval"]    = [=](Interpreter& I, std::vector<Value>& a){ return I.evalExpression(S(a,0)); };
    B["getref"]  = [=](Interpreter& I, std::vector<Value>& a){ return I.makeFuncRef(S(a,0)); };
    B["erase"]   = [=](Interpreter&, std::vector<Value>&){ raiseErr(1, "Erase must be used as a statement: Erase arrayVar"); return Value::empty(); };
    B["monthname"] = [=](Interpreter&, std::vector<Value>& a){ long n=N(a,0); if(n<1||n>12) raiseErr(5,"Invalid procedure call"); std::string m=kMonths[n-1]; bool ab=a.size()>1&&a[1].toBool(); return Value::str(ab?m.substr(0,3):m); };
    B["weekdayname"] = [=](Interpreter&, std::vector<Value>& a){ long n=N(a,0); bool ab=a.size()>1&&a[1].toBool(); long fd=a.size()>2?N(a,2):1; if(fd<1)fd=1; long idx=((n-1)+(fd-1))%7; if(idx<0)idx+=7; std::string d=kDays[idx]; return Value::str(ab?d.substr(0,3):d); };
    B["formatnumber"]  = [=](Interpreter&, std::vector<Value>& a){ int dg=has(a,1)?(int)N(a,1):2; bool grp=has(a,4)?a[4].toBool():true; return Value::str(fmtGrouped(D(a,0),dg,grp)); };
    B["formatpercent"] = [=](Interpreter&, std::vector<Value>& a){ int dg=has(a,1)?(int)N(a,1):2; bool grp=has(a,4)?a[4].toBool():true; return Value::str(fmtGrouped(D(a,0)*100.0,dg,grp)+"%"); };
    B["formatcurrency"]= [=](Interpreter&, std::vector<Value>& a){ int dg=has(a,1)?(int)N(a,1):2; bool grp=has(a,4)?a[4].toBool():true; double v=D(a,0); std::string s=fmtGrouped(std::fabs(v),dg,grp); return Value::str((v<0?"($":"$")+s+(v<0?")":"")); };
    B["datepart"] = [=](Interpreter&, std::vector<Value>& a){ std::string iv=lower(S(a,0)); double ser=D(a,1); int y,mo,d,h,mi,se; ymdhmsFromSerial(ser,y,mo,d,h,mi,se);
        if(iv=="yyyy") return Value::integer(y);
        if(iv=="m")    return Value::integer(mo);
        if(iv=="d")    return Value::integer(d);
        if(iv=="h")    return Value::integer(h);
        if(iv=="n")    return Value::integer(mi);
        if(iv=="s")    return Value::integer(se);
        if(iv=="q")    return Value::integer((mo-1)/3+1);
        if(iv=="w"){ long dd=(long)std::floor(ser)+kBaseDay; int dow=(int)(((dd%7)+7+4)%7); return Value::integer(dow+1); }
        if(iv=="y"){ double jan1=serialFromYMDHMS(y,1,1,0,0,0); return Value::integer((int)(std::floor(ser)-std::floor(jan1))+1); }
        if(iv=="ww"){ double jan1=serialFromYMDHMS(y,1,1,0,0,0); return Value::integer((int)((std::floor(ser)-std::floor(jan1))/7)+1); }
        raiseErr(5,"Invalid interval"); return Value::empty(); };
    B["ccur"]    = [=](Interpreter&, std::vector<Value>& a){ return Value::number(D(a,0)); };
    B["scriptengine"] = [=](Interpreter&, std::vector<Value>&){ return Value::str("Directive"); };
    B["scriptenginemajorversion"] = [=](Interpreter&, std::vector<Value>&){ return Value::integer(1); };
    B["scriptengineminorversion"] = [=](Interpreter&, std::vector<Value>&){ return Value::integer(0); };
    B["scriptenginebuildversion"] = [=](Interpreter&, std::vector<Value>&){ return Value::integer(1); };
    B["getlocale"] = [=](Interpreter&, std::vector<Value>&){ return Value::integer(1033); };
    B["setlocale"] = [=](Interpreter&, std::vector<Value>&){ return Value::integer(1033); };
    B["clipput"]   = [=](Interpreter&, std::vector<Value>& a){ return Value::boolean(clipboardSet(S(a,0))); };
    B["clipget"]   = [=](Interpreter&, std::vector<Value>&){ std::string s; clipboardGet(s); return Value::str(s); };
}

// =============================================================== COM / ActiveX
#ifdef _WIN32
#include <windows.h>
#include <oleauto.h>
#include <mmsystem.h>    // PlaySound / SND_* (link winmm)

static std::wstring utf8ToWide(const std::string& s) {
    if (s.empty()) return std::wstring();
    int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
    std::wstring w(n, 0);
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), &w[0], n);
    return w;
}
static std::string wideToUtf8(const wchar_t* w, int len) {
    if (!w || len == 0) return std::string();
    int n = WideCharToMultiByte(CP_UTF8, 0, w, len, nullptr, 0, nullptr, nullptr);
    std::string s(n, 0);
    WideCharToMultiByte(CP_UTF8, 0, w, len, &s[0], n, nullptr, nullptr);
    return s;
}

struct ComObject : IObject {
    IDispatch* disp = nullptr;
    explicit ComObject(IDispatch* d) : disp(d) { if (disp) disp->AddRef(); }
    ~ComObject() override { if (disp) disp->Release(); }

    std::string typeName() const override { return "Object"; }

    DISPID nameToId(const std::string& name) {
        if (name.empty()) return DISPID_VALUE;
        std::wstring w = utf8ToWide(name);
        LPOLESTR p = (LPOLESTR)w.c_str();
        DISPID id = 0;
        HRESULT hr = disp->GetIDsOfNames(IID_NULL, &p, 1, LOCALE_USER_DEFAULT, &id);
        if (FAILED(hr)) raiseErr(438, "Object doesn't support this property or method: " + name);
        return id;
    }

    static Value variantToValue(VARIANT& v) {
        if (v.vt & VT_BYREF && v.vt != (VT_BYREF | VT_VARIANT)) {
            // dereference simple byref
        }
        switch (v.vt) {
            case VT_EMPTY: return Value::empty();
            case VT_NULL:  return Value::null();
            case VT_BOOL:  return Value::boolean(v.boolVal != VARIANT_FALSE);
            case VT_I1:    return Value::integer(v.cVal);
            case VT_UI1:   return Value::integer(v.bVal);
            case VT_I2:    return Value::integer(v.iVal);
            case VT_UI2:   return Value::integer(v.uiVal);
            case VT_I4:    return Value::integer(v.lVal);
            case VT_UI4:   return Value::number((double)v.ulVal);
            case VT_INT:   return Value::integer(v.intVal);
            case VT_UINT:  return Value::number((double)v.uintVal);
            case VT_I8:    return Value::number((double)v.llVal);
            case VT_UI8:   return Value::number((double)v.ullVal);
            case VT_R4:    return Value::number(v.fltVal);
            case VT_R8:    return Value::number(v.dblVal);
            case VT_DATE:  return Value::number(v.date);   // OLE date epoch == VBScript serial
            case VT_BSTR:  return Value::str(wideToUtf8(v.bstrVal, v.bstrVal ? (int)SysStringLen(v.bstrVal) : 0));
            case VT_DISPATCH: return v.pdispVal ? Value::object(std::make_shared<ComObject>(v.pdispVal)) : Value::nothing();
            case VT_UNKNOWN: {
                if (!v.punkVal) return Value::nothing();
                IDispatch* d = nullptr;
                if (SUCCEEDED(v.punkVal->QueryInterface(IID_IDispatch, (void**)&d)) && d) {
                    Value r = Value::object(std::make_shared<ComObject>(d)); d->Release(); return r;
                }
                return Value::nothing();
            }
            case (VT_BYREF | VT_VARIANT): return v.pvarVal ? variantToValue(*v.pvarVal) : Value::empty();
            default: return Value::empty();   // SAFEARRAY / unsupported
        }
    }
    static void valueToVariant(const Value& val, VARIANT& out) {
        VariantInit(&out);
        switch (val.type) {
            case Value::Type::Empty:  out.vt = VT_EMPTY; break;
            case Value::Type::Null:   out.vt = VT_NULL; break;
            case Value::Type::Bool:   out.vt = VT_BOOL; out.boolVal = val.b ? VARIANT_TRUE : VARIANT_FALSE; break;
            case Value::Type::Int:    out.vt = VT_I4;   out.lVal = val.i; break;
            case Value::Type::Double: out.vt = VT_R8;   out.dblVal = val.d; break;
            case Value::Type::String: { std::wstring w = utf8ToWide(val.s); out.vt = VT_BSTR; out.bstrVal = SysAllocStringLen(w.data(), (UINT)w.size()); break; }
            case Value::Type::Object: {
                if (val.obj) { ComObject* c = dynamic_cast<ComObject*>(val.obj.get());
                               if (c && c->disp) { out.vt = VT_DISPATCH; out.pdispVal = c->disp; c->disp->AddRef(); break; } }
                out.vt = VT_DISPATCH; out.pdispVal = nullptr; break;   // Nothing / non-COM native
            }
            default: out.vt = VT_EMPTY; break;   // arrays unsupported across COM boundary
        }
    }

    Value invoke(DISPID id, WORD flags, std::vector<Value>& args, const Value* putVal) {
        std::vector<VARIANTARG> vargs;
        // positional args go in reverse order
        for (size_t k = args.size(); k-- > 0; ) { VARIANTARG va; valueToVariant(args[k], va); vargs.push_back(va); }
        DISPID putId = DISPID_PROPERTYPUT;
        DISPPARAMS dp; ZeroMemory(&dp, sizeof(dp));
        VARIANTARG putVar; VariantInit(&putVar);
        if (putVal) {
            valueToVariant(*putVal, putVar);
            // property value must be rgvarg[0]
            vargs.insert(vargs.begin(), putVar);
            dp.cNamedArgs = 1; dp.rgdispidNamedArgs = &putId;
        }
        dp.cArgs = (UINT)vargs.size();
        dp.rgvarg = vargs.empty() ? nullptr : vargs.data();

        VARIANT result; VariantInit(&result);
        EXCEPINFO ei; ZeroMemory(&ei, sizeof(ei));
        UINT argErr = 0;
        HRESULT hr = disp->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &dp, &result, &ei, &argErr);
        for (auto& va : vargs) VariantClear(&va);
        if (FAILED(hr)) {
            std::string desc = "COM Invoke failed";
            long num = (long)(hr & 0xFFFF);
            if (hr == DISP_E_EXCEPTION && ei.bstrDescription)
                desc = wideToUtf8(ei.bstrDescription, (int)SysStringLen(ei.bstrDescription));
            if (ei.bstrSource) SysFreeString(ei.bstrSource);
            if (ei.bstrDescription) SysFreeString(ei.bstrDescription);
            if (ei.bstrHelpFile) SysFreeString(ei.bstrHelpFile);
            raiseErr(num ? num : 440, desc);
        }
        Value rv = variantToValue(result);
        VariantClear(&result);
        return rv;
    }

    Value get(Interpreter&, const std::string& name, std::vector<Value>& args) override {
        DISPID id = nameToId(name);
        return invoke(id, DISPATCH_METHOD | DISPATCH_PROPERTYGET, args, nullptr);
    }
    void set(Interpreter&, const std::string& name, std::vector<Value>& args, const Value& val, bool isSet) override {
        DISPID id = nameToId(name);
        invoke(id, isSet ? DISPATCH_PROPERTYPUTREF : DISPATCH_PROPERTYPUT, args, &val);
    }
    bool enumerate(Interpreter&, std::vector<Value>& out) override {
        DISPPARAMS dp; ZeroMemory(&dp, sizeof(dp));
        VARIANT res; VariantInit(&res);
        HRESULT hr = disp->Invoke(DISPID_NEWENUM, IID_NULL, LOCALE_USER_DEFAULT,
                                  DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, &res, nullptr, nullptr);
        if (FAILED(hr)) return false;
        IEnumVARIANT* e = nullptr;
        IUnknown* unk = (res.vt == VT_UNKNOWN) ? res.punkVal : (res.vt == VT_DISPATCH ? (IUnknown*)res.pdispVal : nullptr);
        if (unk) unk->QueryInterface(IID_IEnumVARIANT, (void**)&e);
        VariantClear(&res);
        if (!e) return false;
        VARIANT item; VariantInit(&item); ULONG fetched = 0;
        while (e->Next(1, &item, &fetched) == S_OK && fetched == 1) { out.push_back(variantToValue(item)); VariantClear(&item); }
        e->Release();
        return true;
    }
};

static ObjectPtr createComObject(const std::string& progid, bool& ok) {
    ok = false;
    std::wstring w = utf8ToWide(progid);
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(w.c_str(), &clsid);
    if (FAILED(hr)) return nullptr;
    IDispatch* disp = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                          IID_IDispatch, (void**)&disp);
    if (FAILED(hr) || !disp) return nullptr;
    auto obj = std::make_shared<ComObject>(disp);
    disp->Release();   // ComObject holds its own ref
    ok = true;
    return obj;
}

// ---- Mouse automation (Windows) ----
struct MouseObject : IObject {
    static void button(const std::string& btn, bool down) {
        std::string b = btn.empty() ? "left" : toLower(btn);
        DWORD f = 0;
        if      (b == "left")   f = down ? MOUSEEVENTF_LEFTDOWN   : MOUSEEVENTF_LEFTUP;
        else if (b == "right")  f = down ? MOUSEEVENTF_RIGHTDOWN  : MOUSEEVENTF_RIGHTUP;
        else if (b == "middle") f = down ? MOUSEEVENTF_MIDDLEDOWN : MOUSEEVENTF_MIDDLEUP;
        else raiseErr(5, "Mouse: unknown button '" + btn + "' (use left/right/middle)");
        mouse_event(f, 0, 0, 0, 0);
    }
    std::string typeName() const override { return "Mouse"; }
    Value get(Interpreter&, const std::string& member, std::vector<Value>& args) override {
        std::string m = toLower(member);
        auto N = [&](size_t i)->long { return i < args.size() ? args[i].toInt() : 0; };
        if (m == "move") {
            long x = N(0), y = N(1), speed = args.size() > 2 ? args[2].toInt() : 0;
            if (speed <= 0) SetCursorPos((int)x, (int)y);
            else {
                POINT p; GetCursorPos(&p);
                int steps = (int)std::min<long>(std::max<long>(speed, 1) * 10, 400);
                for (int s = 1; s <= steps; ++s) {
                    SetCursorPos((int)(p.x + (x - p.x) * s / steps),
                                 (int)(p.y + (y - p.y) * s / steps));
                    Sleep(5);
                }
            }
            return Value::empty();
        }
        if (m == "click") {                              // Click([button],[x],[y],[count])
            std::string btn = !args.empty() ? args[0].toStr() : "left";
            if (args.size() >= 3) SetCursorPos((int)N(1), (int)N(2));
            long count = args.size() > 3 ? args[3].toInt() : 1;
            for (long i = 0; i < (count < 1 ? 1 : count); ++i) { button(btn, true); button(btn, false); }
            return Value::empty();
        }
        if (m == "down") { button(!args.empty() ? args[0].toStr() : "left", true);  return Value::empty(); }
        if (m == "up")   { button(!args.empty() ? args[0].toStr() : "left", false); return Value::empty(); }
        if (m == "wheel") {                              // Wheel("up"|"down",[clicks])
            std::string dir = !args.empty() ? toLower(args[0].toStr()) : "up";
            long clicks = args.size() > 1 ? args[1].toInt() : 1;
            int delta = (int)(WHEEL_DELTA * clicks) * (dir == "down" ? -1 : 1);
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, (DWORD)delta, 0);
            return Value::empty();
        }
        if (m == "x" || m == "y" || m == "getpos") {
            POINT p; GetCursorPos(&p);
            if (m == "x") return Value::integer(p.x);
            if (m == "y") return Value::integer(p.y);
            auto ad = std::make_shared<ArrayData>();     // GetPos -> [x, y]
            ad->uppers = { 1 };
            ad->data = { Value::integer(p.x), Value::integer(p.y) };
            return Value::array(ad);
        }
        raiseErr(438, "Mouse has no member '" + member + "'");
        return Value::empty();
    }
};

// ---- Screen: pixel color + metrics (Windows) ----
struct ScreenObject : IObject {
    std::string typeName() const override { return "Screen"; }
    Value get(Interpreter&, const std::string& member, std::vector<Value>& args) override {
        std::string m = toLower(member);
        if (m == "pixelcolor" || m == "getpixel") {
            if (args.size() < 2) raiseErr(5, "Screen.PixelColor requires x, y");
            HDC dc = GetDC(NULL);
            COLORREF c = GetPixel(dc, (int)args[0].toInt(), (int)args[1].toInt());
            ReleaseDC(NULL, dc);
            if (c == CLR_INVALID) raiseErr(5, "Screen.PixelColor: coordinates off-screen");
            long rgb = ((long)GetRValue(c) << 16) | ((long)GetGValue(c) << 8) | (long)GetBValue(c);
            return Value::integer((int32_t)rgb);         // 0xRRGGBB (like AutoIt PixelGetColor)
        }
        if (m == "width")  return Value::integer(GetSystemMetrics(SM_CXSCREEN));
        if (m == "height") return Value::integer(GetSystemMetrics(SM_CYSCREEN));
        raiseErr(438, "Screen has no member '" + member + "'");
        return Value::empty();
    }
};

// ---- Sound playback (Windows, via winmm PlaySound) ----
struct SoundObject : IObject {
    std::string typeName() const override { return "Sound"; }
    Value get(Interpreter&, const std::string& member, std::vector<Value>& args) override {
        std::string m = toLower(member);
        if (m == "play") {                               // Play(file, [wait])  wait=True -> synchronous
            if (args.empty()) raiseErr(5, "Sound.Play requires a file path");
            bool wait = args.size() > 1 && args[1].toBool();
            std::wstring w = utf8ToWide(args[0].toStr());
            BOOL ok = PlaySoundW(w.c_str(), NULL, SND_FILENAME | (wait ? SND_SYNC : SND_ASYNC));
            return Value::boolean(ok ? true : false);
        }
        if (m == "stop") { PlaySoundW(NULL, NULL, 0); return Value::empty(); }
        if (m == "beep") {
            Beep(args.size() > 0 ? (DWORD)args[0].toInt() : 800,
                 args.size() > 1 ? (DWORD)args[1].toInt() : 200);
            return Value::empty();
        }
        raiseErr(438, "Sound has no member '" + member + "'");
        return Value::empty();
    }
};

static ObjectPtr makeMouseObj()  { return std::make_shared<MouseObject>(); }
static ObjectPtr makeScreenObj() { return std::make_shared<ScreenObject>(); }
static ObjectPtr makeSoundObj()  { return std::make_shared<SoundObject>(); }

// ---- Clipboard (Windows, CF_UNICODETEXT) ----
static bool clipboardSet(const std::string& text) {
    if (!OpenClipboard(NULL)) return false;
    EmptyClipboard();
    std::wstring w = utf8ToWide(text);
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE, (w.size() + 1) * sizeof(wchar_t));
    if (!h) { CloseClipboard(); return false; }
    wchar_t* dst = (wchar_t*)GlobalLock(h);
    if (!dst) { GlobalFree(h); CloseClipboard(); return false; }
    for (size_t i = 0; i < w.size(); ++i) dst[i] = w[i];
    dst[w.size()] = 0;
    GlobalUnlock(h);
    HANDLE r = SetClipboardData(CF_UNICODETEXT, h);   // clipboard takes ownership on success
    CloseClipboard();
    if (!r) { GlobalFree(h); return false; }
    return true;
}
static bool clipboardGet(std::string& out) {
    out.clear();
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) return false;
    if (!OpenClipboard(NULL)) return false;
    HANDLE h = GetClipboardData(CF_UNICODETEXT);
    if (!h) { CloseClipboard(); return false; }
    const wchar_t* w = (const wchar_t*)GlobalLock(h);
    if (!w) { CloseClipboard(); return false; }
    int len = 0; while (w[len]) ++len;
    out = wideToUtf8(w, len);
    GlobalUnlock(h);
    CloseClipboard();
    return true;
}
#else
static ObjectPtr createComObject(const std::string&, bool& ok) { ok = false; return nullptr; }
static ObjectPtr makeMouseObj()  { raiseErr(429, "Mouse control requires the Windows build of Directive"); }
static ObjectPtr makeScreenObj() { raiseErr(429, "Screen/pixel access requires the Windows build of Directive"); }
static ObjectPtr makeSoundObj()  { raiseErr(429, "Sound playback requires the Windows build of Directive"); }
static bool clipboardSet(const std::string&) { raiseErr(429, "ClipPut requires the Windows build of Directive"); }
static bool clipboardGet(std::string&)       { raiseErr(429, "ClipGet requires the Windows build of Directive"); }
#endif

// ============================================================================ MAIN
static std::string readAll(std::istream& in) {
    std::ostringstream ss; ss << in.rdbuf(); return ss.str();
}

// ---------------------------------------------------------------- ui layer impl
#if !(defined(_WIN32) && defined(DIRECTIVE_GUI))
// ---- console (cscript-style): everything goes through stdio ----
namespace ui {
    void echo(const std::string& t) { std::cout << t << "\n"; }
    long msgBox(const std::string& p, long, const std::string& title) {
        if (!title.empty()) std::cout << "[" << title << "] ";
        std::cout << p << "\n";
        return 1;   // vbOK (console can't offer other buttons without prompting)
    }
    bool inputBox(const std::string& p, const std::string& title, const std::string& def, std::string& out) {
        if (!title.empty()) std::cerr << "[" << title << "] ";
        std::cerr << p << " ";
        if (!std::getline(std::cin, out)) out = def;
        return true;
    }
    void error(const std::string& t) { std::cerr << t << "\n"; }
}
#else
// ---- Windows GUI (wscript-style): real dialogs. Build with -DDIRECTIVE_GUI ----
namespace ui {
    static std::wstring toW(const std::string& s) { return utf8ToWide(s); }

    void echo(const std::string& t) {
        std::wstring w = toW(t);
        MessageBoxW(NULL, w.c_str(), L"Directive", MB_OK | MB_SETFOREGROUND);
    }
    long msgBox(const std::string& p, long flags, const std::string& title) {
        std::wstring wp = toW(p);
        std::wstring wt = toW(title.empty() ? std::string("VBScript") : title);
        // MessageBox return values (IDOK=1..IDNO=7) equal the VBScript vbOK..vbNo codes.
        return MessageBoxW(NULL, wp.c_str(), wt.c_str(), (UINT)flags | MB_SETFOREGROUND);
    }
    void error(const std::string& t) {
        std::wstring w = toW(t);
        MessageBoxW(NULL, w.c_str(), L"Directive error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
    }

    // ---- InputBox: an in-memory dialog template lets the dialog manager
    //      handle Tab / Enter / Esc / focus for us. ----
    struct InputCtx { std::wstring text; bool ok; };
    static const WORD ID_EDIT = 1000;

    static INT_PTR CALLBACK inputProc(HWND h, UINT m, WPARAM w, LPARAM l) {
        if (m == WM_INITDIALOG) {
            InputCtx* c = (InputCtx*)l;
            SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)c);
            SetDlgItemTextW(h, ID_EDIT, c->text.c_str());
            HWND e = GetDlgItem(h, ID_EDIT);
            SetFocus(e); SendMessageW(e, EM_SETSEL, 0, -1);
            return FALSE;   // we set focus ourselves
        }
        if (m == WM_COMMAND) {
            InputCtx* c = (InputCtx*)GetWindowLongPtrW(h, GWLP_USERDATA);
            WORD id = LOWORD(w);
            if (id == IDOK) {
                wchar_t buf[8192]; GetDlgItemTextW(h, ID_EDIT, buf, 8192);
                if (c) { c->text = buf; c->ok = true; }
                EndDialog(h, IDOK); return TRUE;
            }
            if (id == IDCANCEL) { if (c) c->ok = false; EndDialog(h, IDCANCEL); return TRUE; }
        }
        return FALSE;
    }

    static void alignDword(std::vector<BYTE>& b) { while (b.size() & 3) b.push_back(0); }
    static void putW(std::vector<BYTE>& b, WORD v) { b.push_back((BYTE)(v & 0xFF)); b.push_back((BYTE)(v >> 8)); }
    static void putD(std::vector<BYTE>& b, DWORD v) { for (int i = 0; i < 4; i++) b.push_back((BYTE)((v >> (8 * i)) & 0xFF)); }
    static void putStr(std::vector<BYTE>& b, const wchar_t* s) { for (; *s; ++s) putW(b, (WORD)*s); putW(b, 0); }
    static void addItem(std::vector<BYTE>& b, DWORD style, short x, short y, short cx, short cy,
                        WORD id, WORD atom, const wchar_t* text) {
        alignDword(b);
        putD(b, style | WS_CHILD | WS_VISIBLE);
        putD(b, 0);                                   // exStyle
        putW(b, (WORD)x); putW(b, (WORD)y); putW(b, (WORD)cx); putW(b, (WORD)cy);
        putW(b, id);
        putW(b, 0xFFFF); putW(b, atom);               // predefined control class
        putStr(b, text);                              // caption
        putW(b, 0);                                   // no creation data
    }

    bool inputBox(const std::string& prompt, const std::string& title, const std::string& def, std::string& out) {
        InputCtx c; c.text = toW(def); c.ok = false;
        std::wstring wCaption = toW(title.empty() ? std::string("VBScript") : title);
        std::wstring wPrompt  = toW(prompt);

        std::vector<BYTE> t;
        putD(t, DS_MODALFRAME | DS_SETFONT | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU);
        putD(t, 0);                                   // exStyle
        putW(t, 4);                                   // number of controls
        putW(t, 0); putW(t, 0);                       // x, y
        putW(t, 240); putW(t, 100);                   // cx, cy (dialog units)
        putW(t, 0);                                   // no menu
        putW(t, 0);                                   // default dialog class
        putStr(t, wCaption.c_str());                  // caption
        putW(t, 8); putStr(t, L"MS Shell Dlg");       // DS_SETFONT: size + face

        addItem(t, SS_LEFT,                                7,  7, 226, 40, (WORD)-1, 0x0082, wPrompt.c_str()); // static
        addItem(t, ES_AUTOHSCROLL | WS_BORDER | WS_TABSTOP, 7, 50, 226, 14, ID_EDIT,  0x0081, L"");            // edit
        addItem(t, BS_DEFPUSHBUTTON | WS_TABSTOP,          128, 74, 50, 16, IDOK,     0x0080, L"OK");
        addItem(t, BS_PUSHBUTTON | WS_TABSTOP,             183, 74, 50, 16, IDCANCEL, 0x0080, L"Cancel");

        INT_PTR r = DialogBoxIndirectParamW(GetModuleHandleW(NULL), (LPCDLGTEMPLATEW)t.data(),
                                            NULL, inputProc, (LPARAM)&c);
        if (r == IDOK && c.ok) { out = wideToUtf8(c.text.c_str(), (int)c.text.size()); return true; }
        out.clear();
        return false;
    }
}
#endif

// ---------------------------------------------------------------- driver
static int runMain(int argc, char** argv) {
#ifdef _WIN32
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
#endif
    std::string source, path;
    bool inline_mode = false;
    bool haveSource = false;                 // set once the script (file, -e, or stdin) is determined
    std::vector<std::string> scriptArgs;     // everything after the script becomes Directive.Arguments

    for (int i = 1; i < argc; ++i) {
        std::string arg = argv[i];
        if (!haveSource && (arg == "-e" || arg == "--eval")) {
            if (i + 1 < argc) { source = argv[++i]; inline_mode = true; haveSource = true; }
            else { ui::error("directive: -e requires an argument"); return 2; }
        } else if (!haveSource && (arg == "-h" || arg == "--help")) {
            ui::echo("Directive (Open Directive) - a VBScript-like scripting language\n"
                     "Usage:\n"
                     "  directive <script.directive> [args...]   run a script file\n"
                     "  directive -e \"code\" [args...]            run an inline snippet\n"
                     "  directive - [args...]                    read script from stdin\n"
                     "\n"
                     "Any arguments after the script are readable from the script via\n"
                     "Directive.Arguments (0-based): Directive.Arguments.Count,\n"
                     "Directive.Arguments(0), and 'For Each a In Directive.Arguments'.");
            return 0;
        } else if (!haveSource && arg == "-") {
            source = readAll(std::cin); inline_mode = true; haveSource = true;
        } else if (!haveSource) {
            path = arg; haveSource = true;
        } else {
            scriptArgs.push_back(arg);       // positional (or flag-looking) arg belonging to the script
        }
    }

    if (!inline_mode) {
        if (path.empty()) { ui::error("directive: no input. Try 'directive --help'."); return 2; }
        std::ifstream f(path, std::ios::binary);
        if (!f) { ui::error("directive: cannot open file: " + path); return 2; }
        source = readAll(f);
    }

    // strip UTF-8 BOM if present
    if (source.size() >= 3 && (unsigned char)source[0] == 0xEF &&
        (unsigned char)source[1] == 0xBB && (unsigned char)source[2] == 0xBF)
        source = source.substr(3);

    int rc = 0;
    Interpreter interp;

    // Resolve the script directory (used for Include path resolution and Directive.ScriptPath).
    std::string baseDir, fullPath;
    auto included = std::make_shared<std::set<std::string>>();
    {
        std::error_code ec;
        if (!inline_mode && !path.empty()) {
            fs::path canon = fs::weakly_canonical(path, ec);
            if (ec || canon.empty()) canon = fs::absolute(path, ec);
            fullPath = canon.string();
            included->insert(fullPath);                 // guard: the main file can't be re-included
            baseDir = canon.parent_path().string();
        }
        if (baseDir.empty()) baseDir = fs::current_path(ec).string();
    }
    interp.setScriptInfo(fullPath, baseDir);
    interp.setArguments(scriptArgs);

    try {
        Lexer lexer(source);
        Parser parser(lexer.tokenize(), baseDir, included);
        std::vector<StmtP> program = parser.parseProgram();
        interp.run(program);
    } catch (QuitSignal& q) {
        rc = q.code;
    } catch (ExitSignal&) {
        rc = 0;   // stray Exit at top level: ignore
    } catch (VbError& e) {
        // Report in the same shape as the Windows Script Host dialog:
        // Script / Line / Char / Error / Code / Source.
        long ln  = e.line ? e.line : interp.errorLine();
        long col = e.col  ? e.col  : interp.errorCol();
        unsigned hres = (e.number & 0xFFFF0000L)
                        ? (unsigned)e.number                              // already a full HRESULT (e.g. Err.Raise)
                        : (0x800A0000u | (unsigned)(e.number & 0xFFFF));  // standard code -> 800Axxxx
        char code[16]; std::snprintf(code, sizeof(code), "%08X", hres);
        std::ostringstream m;
        m << "Script:\t" << (fullPath.empty() ? "(inline script)" : fullPath) << "\n"
          << "Line:\t"   << (ln  > 0 ? std::to_string(ln)  : std::string("1")) << "\n"
          << "Char:\t"   << (col > 0 ? std::to_string(col) : std::string("1")) << "\n"
          << "Error:\t"  << e.description << "\n"
          << "Code:\t"   << code << "\n"
          << "Source:\t" << (e.source.empty() ? "Directive runtime error" : e.source);
        ui::error(m.str());
        rc = 1;
    } catch (const std::exception& e) {
        ui::error(std::string("Directive internal error: ") + e.what());
        rc = 70;
    }

#ifdef _WIN32
    CoUninitialize();
#endif
    return rc;
}

// Console entry point (default build, and all non-Windows builds).
int main(int argc, char** argv) {
    return runMain(argc, argv);
}

#if defined(_WIN32) && defined(DIRECTIVE_GUI)
// GUI entry point: no console window. Build with -DDIRECTIVE_GUI and a windowed
// subsystem (MSVC: /link /SUBSYSTEM:WINDOWS ; MinGW: -mwindows). Using the ANSI
// WinMain means no -municode is required; we still read the Unicode command line.
int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int) {
    // If we were launched from a console (cmd), attach to it so Directive.StdOut /
    // .StdIn work there. If double-clicked, there's no parent console (that's fine).
    if (AttachConsole(ATTACH_PARENT_PROCESS)) {
#ifdef _MSC_VER
        FILE* f;
        freopen_s(&f, "CONOUT$", "w", stdout);
        freopen_s(&f, "CONOUT$", "w", stderr);
        freopen_s(&f, "CONIN$",  "r", stdin);
#else
        (void)!freopen("CONOUT$", "w", stdout);
        (void)!freopen("CONOUT$", "w", stderr);
        (void)!freopen("CONIN$",  "r", stdin);
#endif
    }
    int argc = 0;
    LPWSTR* wargv = CommandLineToArgvW(GetCommandLineW(), &argc);
    std::vector<std::string> args;
    for (int i = 0; i < argc; ++i) args.push_back(wideToUtf8(wargv[i], lstrlenW(wargv[i])));
    if (wargv) LocalFree(wargv);
    std::vector<char*> argp;
    for (auto& a : args) argp.push_back(&a[0]);
    return runMain((int)argp.size(), argp.empty() ? nullptr : argp.data());
}
#endif
