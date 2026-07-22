# Directive — Complete API Reference

A checklist of everything the interpreter exposes: every built-in function with
its parameters, all statements, operators, constants, and objects. Function and
member names were extracted directly from the source, so this matches what the
engine actually implements.

**Legend**
- `[arg]` = optional argument
- ✔ tested = exercised by the regression/edge suites on this platform
- 🪟 Windows-only = compiles and links, runs only on real Windows (raises a clean
  error on non-Windows). Items marked 🪟 have **not** been runtime-tested.

---

## 1. Running scripts

```
directive script.directive          # run a file (console build)
directive script.directive a b c    # run a file, passing arguments (Directive.Arguments)
directive -e "directive.echo 1+1"   # run an inline snippet
directivew script.directive         # GUI build: Echo/Input via message boxes
```

---

## 2. Built-in functions (101)

### Strings
| Function | Parameters | Notes |
|---|---|---|
| `Len` | `(string \| var)` | length; `Len(array)` not supported — use `UBound` |
| `Left` | `(string, count)` | |
| `Right` | `(string, count)` | |
| `Mid` | `(string, start, [length])` | 1-based; also a **statement** (see §4) |
| `InStr` | `([start,] haystack, needle, [compare])` | 1-based; 0 if not found |
| `InStrRev` | `(haystack, needle, [start, [compare]])` | |
| `Replace` | `(string, find, repl, [start, [count, [compare]]])` | |
| `Split` | `(string, [delim, [count, [compare]]])` | ✔ `count`/`compare` supported |
| `Join` | `(array, [delim])` | |
| `UCase` / `LCase` | `(string)` | |
| `Trim` / `LTrim` / `RTrim` | `(string)` | |
| `Space` | `(count)` | |
| `String` | `(count, char)` | char may be code or 1-char string |
| `StrReverse` | `(string)` | |
| `StrComp` | `(a, b, [compare])` | returns -1/0/1 |
| `Filter` | `(array, match, [include, [compare]])` | include=False excludes matches |
| `Asc` / `AscW` | `(string)` | |
| `Chr` / `ChrW` | `(code)` | |

### Conversion
| Function | Parameters | Notes |
|---|---|---|
| `CInt` | `(expr)` | banker's rounding |
| `CLng` | `(expr)` | banker's rounding |
| `CByte` | `(expr)` | 0–255 |
| `CDbl` / `CSng` | `(expr)` | |
| `CCur` | `(expr)` | currency (stored as double) |
| `CStr` | `(expr)` | |
| `CBool` | `(expr)` | |
| `CDate` | `(expr)` | string/number → date serial |
| `Hex` | `(number)` | 32-bit two's complement for negatives |
| `Oct` | `(number)` | |
| `Int` | `(number)` | floor (toward −∞) |
| `Fix` | `(number)` | truncate (toward 0) |
| `Round` | `(number, [digits])` | banker's (round-half-to-even) |

### Math
| Function | Parameters |
|---|---|
| `Abs` | `(number)` |
| `Sgn` | `(number)` |
| `Sqr` | `(number)` |
| `Atn` / `Cos` / `Sin` / `Tan` | `(number)` — radians |
| `Exp` / `Log` | `(number)` — natural |
| `Rnd` | `([seed])` |
| `Randomize` | `([seed])` — statement-style |
| `RGB` | `(r, g, b)` → 0xBBGGRR long |

### Dates
| Function | Parameters | Notes |
|---|---|---|
| `Now` / `Date` / `Time` | `()` | |
| `Year` / `Month` / `Day` | `(date)` | |
| `Hour` / `Minute` / `Second` | `(date)` | |
| `Weekday` | `(date, [firstDayOfWeek])` | |
| `MonthName` | `(n, [abbrev])` | |
| `WeekdayName` | `(n, [abbrev, [firstDayOfWeek]])` | |
| `DateSerial` | `(year, month, day)` | |
| `TimeSerial` | `(hour, minute, second)` | |
| `DateValue` | `(expr)` | ✔ date part only |
| `TimeValue` | `(expr)` | ✔ time-of-day only |
| `DateAdd` | `(interval, number, date)` | intervals `yyyy,m,d,ww,h,n,s` |
| `DateDiff` | `(interval, date1, date2, [fdow, [fwoy]])` | |
| `DatePart` | `(interval, date, [fdow, [fwoy]])` | |

### Formatting
| Function | Parameters |
|---|---|
| `FormatNumber` | `(number, [decimals, [leadingDigit, [parensNeg, [groupDigits]]]])` |
| `FormatCurrency` | `(number, [decimals, …])` |
| `FormatPercent` | `(number, [decimals, …])` |
| `FormatDateTime` | `(date, [namedFormat])` — 0 General,1 LongDate,2 ShortDate,3 LongTime,4 ShortTime |

### Arrays
| Function | Parameters | Notes |
|---|---|---|
| `Array` | `(v1, v2, …)` | build a variant array |
| `UBound` | `(array, [dimension])` | |
| `LBound` | `(array, [dimension])` | always 0 |
| `Erase` | `(array)` | **statement** (see §4) |

### Type inspection
| Function | Parameters |
|---|---|
| `IsArray` / `IsDate` / `IsEmpty` / `IsNull` / `IsNumeric` / `IsObject` | `(expr)` |
| `TypeName` | `(expr)` → `"Long"`,`"Double"`,`"String"`,`"Boolean"`,`"Empty"`,`"Null"`,`"Object"`,class name, etc. |
| `VarType` | `(expr)` → numeric VbVarType code |

### Evaluation & references
| Function | Parameters | Notes |
|---|---|---|
| `Eval` | `(string)` | evaluate an expression |
| `Execute` | `(string)` | run statements in local scope |
| `ExecuteGlobal` | `(string)` | run statements in global scope |
| `GetRef` | `(procName)` | returns a callable reference |

### Engine / locale
| Function | Parameters |
|---|---|
| `ScriptEngine` | `()` → `"Directive"` |
| `ScriptEngineMajorVersion` / `ScriptEngineMinorVersion` / `ScriptEngineBuildVersion` | `()` |
| `GetLocale` / `SetLocale` | `([locale])` |

### Dialogs & clipboard
| Function | Parameters | Notes |
|---|---|---|
| `MsgBox` | `(prompt, [buttons, [title]])` | console prints; GUI shows a box |
| `InputBox` | `(prompt, [title, [default]])` | |
| `ClipPut` | `(text)` | 🪟 Windows-only |
| `ClipGet` | `()` | 🪟 Windows-only |

### COM (Windows)
| Function | Parameters | Notes |
|---|---|---|
| `CreateObject` | `(progID)` | 🪟 real COM/ActiveX; err 429 elsewhere |
| `GetObject` | `([path, [progID]])` | 🪟 |

---

## 3. Operators
Arithmetic `+  -  *  /  \ (int div)  Mod  ^ (power)` ·
Concatenation `&` (always string; `+` concatenates only when **both** sides are
strings, else adds) · Comparison `=  <>  <  >  <=  >=` ·
Object identity `Is` · Logical/bitwise `And  Or  Not  Xor  Eqv  Imp`.

---

## 4. Statements & language constructs
- `Option Explicit`
- `Dim` / `Public` / `Private` / `ReDim [Preserve]` / `Const`
- `Set x = …` (object assignment), plain `x = …` (value assignment)
- `If … Then … [ElseIf …] [Else] … End If` (and single-line `If … Then …`)
- `For … To … [Step …] … Next`
- `For Each … In … … Next`
- `Do [While|Until] … Loop` and `Do … Loop [While|Until]`
- `While … Wend`
- `Select Case … Case … [Case Else] End Select` (value lists, `Is` comparisons, `To` ranges)
- `With obj … End With` (nestable)
- `Sub` / `Function` with `ByRef` / `ByVal`; `Exit Sub`/`Exit Function`
- `Call proc(args)` and paren-less `proc args`
- `Class … End Class` with `Public`/`Private` members, `Property Get/Let/Set`,
  `Public Default` member, `Me`, `Class_Initialize`, `Class_Terminate`; `Exit Property`
- `Mid(target, start[, length]) = value` (in-place string replacement statement)
- `Erase array`
- `On Error Resume Next` / `On Error Goto 0`
- `Randomize [seed]`
- `#…#` date literals, e.g. `#2020-03-15#`, `#1/15/2020 3:45 PM#`, `#13:45:00#`
- Native `Include "path"` (top-level; splices another file at parse time)

---

## 5. Constants
**Whitespace/strings:** `vbCr` `vbLf` `vbCrLf` `vbNewLine` `vbTab` `vbNullString`
`vbNullChar` `vbBack` `vbFormFeed` `vbVerticalTab`
**Booleans/compare:** `vbTrue` `vbFalse` `vbBinaryCompare` `vbTextCompare`
**VarType:** `vbEmpty` `vbNull` `vbInteger` `vbLong` `vbSingle` `vbDouble`
`vbCurrency` `vbDate` `vbString` `vbObject` `vbError` `vbBoolean` `vbVariant`
`vbByte` `vbArray`
**MsgBox buttons/icons/returns:** `vbOKOnly` `vbOKCancel` `vbAbortRetryIgnore`
`vbYesNoCancel` `vbYesNo` `vbRetryCancel` `vbCritical` `vbQuestion` `vbExclamation`
`vbInformation` `vbOK` `vbCancel` `vbAbort` `vbRetry` `vbIgnore` `vbYes` `vbNo`
**Weekdays:** `vbSunday`…`vbSaturday` `vbUseSystemDayOfWeek` `vbFirstJan1`
**Date formats:** `vbGeneralDate` `vbLongDate` `vbShortDate` `vbLongTime` `vbShortTime`
**Errors:** `vbObjectError`

---

## 6. Objects

### 6.1 `Directive` — host object (always available)
| Member | Signature | Notes |
|---|---|---|
| `Echo` | `Directive.Echo a, b, …` | space-joined; objects use their default member |
| `Arguments` (alias `Args`) | `Directive.Arguments` | command-line args collection (see below) |
| `StdOut` / `StdErr` / `StdIn` | stream objects with `.Write`/`.WriteLine`/`.ReadLine`/`.ReadAll` | |
| `Sleep` | `Directive.Sleep ms` | |
| `Quit` | `Directive.Quit [code]` | |
| `ScriptName` / `ScriptFullName` / `ScriptPath` (alias `ScriptDir`) | path info | |

**`Directive.Arguments`** — a 0-based collection of the arguments passed after the
script on the command line (the WSH `WScript.Arguments` equivalent; the name is
always `Directive`, there is no `WScript` alias):

| Member | Signature |
|---|---|
| `Count` (alias `Length`) | `.Count` |
| `Item` (default) | `.Item(i)` / `Directive.Arguments(i)` — 0-based |
| For Each | iterates the argument strings |

```
directive myscript.directive one two "three with spaces" file.txt
```
```vbscript
For Each a In Directive.Arguments
    directive.echo a
Next
```
On Windows, associate the `.directive` extension with `directive.exe "%1" %*`
(or `directivew.exe "%1" %*`) so that dropping files onto a script passes their
paths through as arguments.

### 6.2 `Dictionary` — `New Dictionary`  ✔ tested
Keys are **typed**: numeric `1` ≠ string `"1"`; `1` and `1.0` coincide; objects
key by identity.

| Member | Signature |
|---|---|
| `Add` | `.Add key, item` |
| `Item` (default) | `.Item(key)` / `d(key)` — get or set |
| `Exists` | `.Exists(key)` |
| `Remove` | `.Remove key` |
| `RemoveAll` | `.RemoveAll` |
| `Keys` / `Items` | `.Keys` / `.Items` → array |
| `Count` | `.Count` |
| `CompareMode` | read/write (0 binary) |
| For Each | iterates keys |

### 6.3 `List` — `New List` (alias `LinkedList`)  ✔ tested
A dynamic list beyond stock VBScript.

| Member | Signature |
|---|---|
| `Add` / `Push` / `Append` | `.Add value` |
| `Insert` | `.Insert index, value` |
| `Remove` | `.Remove value` (first match) |
| `RemoveAt` | `.RemoveAt index` |
| `Item` (default) | `.Item(index)` / `l(index)` — get or set |
| `Count` | `.Count` |
| `Clear` | `.Clear` |
| `Contains` | `.Contains(value)` |
| `IndexOf` | `.IndexOf(value)` |
| `Reverse` | `.Reverse` |
| `Sort` | `.Sort` |
| `Swap` | `.Swap i, j` |
| `First` / `Last` | `.First` / `.Last` |
| `ToArray` | `.ToArray` → array |
| For Each | iterates items |

### 6.4 `RegExp` — `New RegExp`  ✔ tested
| Member | Signature |
|---|---|
| `Pattern` | read/write |
| `Global` / `IgnoreCase` / `MultiLine` | read/write booleans |
| `Test` | `.Test(string)` → bool |
| `Execute` | `.Execute(string)` → Matches collection |
| `Replace` | `.Replace(string, replacement)` |

**Match** (from `Execute`): `.Value` (default), `.FirstIndex`, `.Length`,
`.SubMatches([i])`, `.Count`. Iterable with For Each.

### 6.5 `FileSystem` — `New FileSystem` (aliases `FileSystemObject`, `FSO`)  ✔ tested
| Member | Signature |
|---|---|
| `FileExists` / `FolderExists` | `(path)` |
| `GetFile` / `GetFolder` | `(path)` → File/Folder |
| `CreateFolder` | `(path)` |
| `CreateTextFile` | `(path, [overwrite, [unicode]])` → TextStream |
| `DeleteFile` / `DeleteFolder` | `(path)` |
| `CopyFile` | `(src, dest)` |
| `GetSpecialFolder` | `(n)` |

**File / Folder** objects: `.Name`, `.Path`, `.ShortName`, `.Size`, `.Type`,
`.Attributes`, `.DateCreated`, `.DateLastModified`, `.DateLastAccessed`,
`.Drive`, `.ParentFolder`, `.IsRootFolder`, `.Copy`, `.Move`, `.Delete`,
`.OpenAsTextStream` (File). Folder also: `.Files`, `.SubFolders`,
`.CreateTextFile`, `.CreateFolder`, `.FileExists`, `.FolderExists`, `.GetFolder`.

**TextStream:** `.Write`, `.WriteLine`, `.ReadLine`, `.ReadAll`,
`.AtEndOfStream`, `.Close`.

### 6.6 `WavWriter` — `New WavWriter` (alias `WavFile`)  ✔ tested (cross-platform)
Concatenates PCM WAV files (same format only).
| Member | Signature |
|---|---|
| `Create` | `.Create(path)` |
| `Append` (alias `Add`) | `.Append(wavPath)` |
| `Duration` | `.Duration` → seconds written |
| `Close` | `.Close` |

### 6.7 `Mouse` — `New Mouse`  🪟 Windows-only
| Member | Signature |
|---|---|
| `Move` | `.Move x, y` |
| `Click` | `.Click([button, [x, [y, [count]]]])` — button `"left"`/`"right"`/`"middle"` |
| `Down` / `Up` | `.Down([button])` / `.Up([button])` |
| `Wheel` | `.Wheel("up"\|"down", [clicks])` |
| `GetPos` | `.GetPos` → position |
| `X` / `Y` | current coordinates |

### 6.8 `Screen` — `New Screen`  🪟 Windows-only
| Member | Signature |
|---|---|
| `Width` / `Height` | screen dimensions |
| `PixelColor` (alias `GetPixel`) | `.PixelColor(x, y)` → 0xRRGGBB |

### 6.9 `Sound` — `New Sound`  🪟 Windows-only
| Member | Signature |
|---|---|
| `Play` | `.Play(file, [wait])` — wait=True blocks until done |
| `Stop` | `.Stop` |
| `Beep` | `.Beep([frequency, [duration]])` |

---

## 7. Known limitations (see README for the full list)
- Not a security sandbox — as capable as VBScript (file access, COM, OS
  automation, Include). Don't run untrusted scripts.
- No distinct Date subtype: dates are OLE serial doubles, so `TypeName(#…#)` is
  `"Double"` and `IsDate(#…#)` is `False` (the literal is already a number).
- `Private` class members are **not enforced** for external access (they behave
  as accessible). They work correctly inside the class.
- A leading-dot argument to a paren-less call (`Directive.Echo .Member` inside a
  `With`) parses as member chaining. Workaround: parenthesize —
  `Directive.Echo(.Member)`.
- Reference cycles leak and skip `Class_Terminate` (same as COM; no GC).
- `Rnd` is not bit-identical to VBScript's generator.
- `Mouse`/`Screen`/`Sound`/`Clipboard`/`CreateObject` are Windows-only and, in
  this build environment, are compile/link-verified but not runtime-tested.
