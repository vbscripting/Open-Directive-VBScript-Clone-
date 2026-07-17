# Open Directive (VBScript-Clone)

VBScript will be retired and eliminated from future versions of Windows...
I really wanna create / need a better alternative now that Microsoft wants to abandon it!

# Directive Script

Directive is a small scripting language with **VBScript-like syntax** that branches
off into its own thing. It's a single self-contained C++17 file (`directive.cpp`,
no external dependencies) that compiles to a native executable.

It keeps the parts of VBScript that are pleasant to write — `Dim`, `If/Select`,
`For/Do`, `Sub/Function`, `Class`, `On Error`/`Err` — and changes the object model
to be cleaner and more explicit:

- **`New` creates native engine objects** (`Dictionary`, `List`, `RegExp`,
  `FileSystem`). They work on every platform.
- **`CreateObject` is reserved for real COM/ActiveX** automation on Windows.
- The host object is **`Directive`** (not `WScript`): `Directive.Echo`,
  `Directive.StdOut.Write`, `Directive.StdIn.ReadLine`, `Directive.Quit`.

## What changed from VBScript

| VBScript | Directive |
|----------|-----------|
| `CreateObject("Scripting.Dictionary")` | `New Dictionary` |
| `CreateObject("Scripting.FileSystemObject")` | `New FileSystem` (or `New FSO`) |
| `CreateObject("VBScript.RegExp")` / `New RegExp` | `New RegExp` |
| *(no built-in list)* | `New List` |
| *(AutoIt-style automation)* | `New Mouse`, `New Screen`, `New Sound`, `New WavWriter` |
| `CreateObject("Excel.Application")` | `CreateObject("Excel.Application")` — real COM, unchanged |
| `WScript.Echo` | `Directive.Echo` |
| `WScript.StdOut` / `.StdIn` | `Directive.StdOut` / `.StdIn` |
| `WScript.Quit` | `Directive.Quit` |

So `New` = in-process, cross-platform, engine-provided; `CreateObject` = the
operating system's COM, Windows-only. The split keeps the two worlds from blurring.

## Build

One file, no build system.

**Console build** (cscript-style: `Directive.Echo` and the `StdOut`/`StdIn`
streams use the terminal). This is the portable default and the only mode on
Linux/macOS.

```sh
# Linux / macOS
g++ -std=c++17 -O2 directive.cpp -o directive

# Windows, MSVC (x64 Native Tools prompt)
cl /std:c++17 /EHsc /O2 directive.cpp /Fe:directive.exe /link /STACK:16777216

# Windows, MinGW-w64 — link statically so the .exe is self-contained (no DLLs to ship)
g++ -std=c++17 -O2 -static -static-libgcc -static-libstdc++ -Wl,--stack,16777216 directive.cpp -o directive.exe -lole32 -loleaut32 -luuid -lwinmm -lgdi32 -luser32
```

**GUI build** (wscript-style: `MsgBox`, `InputBox`, and `Directive.Echo` show real
windows). Add `-DDIRECTIVE_GUI` and a windowed subsystem:

```sh
# Windows, MSVC
cl /std:c++17 /EHsc /O2 /DDIRECTIVE_GUI directive.cpp /Fe:directivew.exe /link /SUBSYSTEM:WINDOWS /STACK:16777216

# Windows, MinGW-w64 (static)
g++ -std=c++17 -O2 -DDIRECTIVE_GUI -static -static-libgcc -static-libstdc++ -Wl,--stack,16777216 directive.cpp -o directivew.exe -mwindows -lole32 -loleaut32 -luuid -lwinmm -lgdi32 -luser32
```

> **MinGW users: use `-static`.** Without it the `.exe` depends on
> `libstdc++-6.dll` / `libgcc_s_seh-1.dll`, which won't exist on a clean machine
> and the program will fail to start. Static linking bundles everything into one
> file. (MSVC links its runtime differently and doesn't need this.)

The GUI build also calls `AttachConsole` at startup, so if you launch it from a
`cmd` window `Directive.StdOut` still writes there; double-clicked, it just shows
dialogs.

## Run

```sh
./directive script.directive          # run a file
./directive -e "Directive.Echo 6*7"   # inline
./directive -                         # read from stdin
```

The two I/O modes:

- **A — GUI:** associate `.directive` with `directivew.exe` and double-click a
  script; `MsgBox` / `InputBox` / `Directive.Echo` pop up windows.
- **B — console:** run `directive.exe script.directive` from a terminal and use
  `Directive.Echo`, `Directive.StdOut.Write`, `Directive.StdIn.ReadLine`, etc.

## Error messages

Errors are reported in the same shape as the Windows Script Host dialog — script
path, line, column, message, an HRESULT-style code, and a source:

```
Script: /path/to/script.directive
Line:   3
Char:   1
Error:  Division by zero
Code:   800A000B
Source: Directive runtime error
```

(The fields are tab-aligned into a column, like the Windows Script Host dialog.)

The `Code` is the classic VBScript form `0x800A0000 + errorNumber`, so it lines up
with the numbers you already know: type mismatch `800A000D` (13), division by zero
`800A000B` (11), subscript out of range `800A0009` (9), object required `800A01A8`
(424), undefined variable `800A01F4` (500). Compilation errors carry
`Source: Directive compilation error` and point at the offending token; runtime
errors point at the statement being executed. In the console build this prints to
stderr; in the GUI build it appears in a message box.

## The `List` object

`New List` is a genuine doubly-linked list:

```vbscript
Dim L : Set L = New List
L.Add "alpha"
L.Insert 1, "beta"          ' by index
L.Swap 0, 1
L.Remove "beta"             ' by value; returns True/False
L.RemoveAt 0                ' by index
x = L(0)                    ' Item(index), 0-based; also L(0) = "new"
L.Reverse
L.Sort
For Each item In L : Directive.Echo item : Next
```

Members: `Add`/`Append`/`Push`, `Insert(i, v)`, `Remove(v)`, `RemoveAt(i)`,
`Item(i)` (get/set, default), `Count`, `Clear`, `Contains(v)`, `IndexOf(v)`,
`Swap(i, j)`, `Reverse`, `Sort`, `First`, `Last`, `ToArray`, and `For Each`.

## The `Directive` host object

- `Directive.Echo a, b, …` — writes a line (console: stdout; GUI: a message box)
- `Directive.StdOut` / `Directive.StdErr` — `.Write`, `.WriteLine`, `.WriteBlankLines(n)`
- `Directive.StdIn` — `.ReadLine`, `.Read(n)`, `.ReadAll`, `.AtEndOfStream`
- `Directive.Quit [code]`, `Directive.Sleep ms`
- `Directive.ScriptName` — the script's file name; `Directive.ScriptFullName` — its
  absolute path; `Directive.ScriptPath` (alias `ScriptDir`) — the folder containing
  it. Handy for locating files relative to the script (`-e`/stdin scripts report
  the current directory).

## Include

A native `Include` pulls another file's definitions into the current script:

```vbscript
Include "lib/strutil.directive"     ' path is relative to THIS file's folder
directive.echo TitleCase("hello world")
```

Unlike the old `ExecuteGlobal`-plus-read-the-file trick, `Include` happens at
**parse time**, so everything the included file defines (Const, Function, Sub,
Class) is visible everywhere in the including file — including *above* the
`Include` line. Details:

- **Paths resolve relative to the including file's directory** (nested includes
  resolve relative to their own location), so a library can `Include` its
  siblings without caring where the top-level script lives.
- **Each file is included at most once.** Diamond includes (A pulls in B and C,
  both of which pull in D) and circular includes (A ↔ B) are handled by an
  include guard — no duplicate-definition errors, no infinite loops.
- **Top level only** — `Include` inside a `Sub`/`Function`/`Class` is a parse
  error. Parse errors inside an included file name the file.

## File / Folder object model

`New FileSystem` gives you the full object model, not just loose functions:

```vbscript
Dim fso : Set fso = New FileSystem
Dim folder : Set folder = fso.GetFolder("demo")
Directive.Echo folder.Name & " has " & folder.Files.Count & " files"
Dim f
For Each f In folder.Files
    Directive.Echo f.Name & " — " & f.Size & " bytes, modified " & f.DateLastModified
Next
Dim doc : Set doc = fso.GetFile("demo\notes.txt")
doc.Copy "demo\backup.txt"
Set ts = doc.OpenAsTextStream(1) : Directive.Echo ts.ReadLine : ts.Close
```

- **FileSystem**: `FileExists`, `FolderExists`, `CreateFolder`, `DeleteFile`,
  `DeleteFolder`, `CopyFile`, `CopyFolder`, `MoveFile`, `MoveFolder`,
  `CreateTextFile`, `OpenTextFile`, `GetFile`, `GetFolder`, `GetSpecialFolder`,
  `GetBaseName`, `GetFileName`, `GetExtensionName`, `GetParentFolderName`,
  `BuildPath`, `GetAbsolutePathName`, `GetTempName`.
- **Folder**: `Path`, `Name`, `ParentFolder`, `Drive`, `Size`, `Type`,
  `Attributes`, `DateLastModified` (etc.), `IsRootFolder`, `Files`, `SubFolders`,
  `Delete`, `Copy`, `Move`, `CreateTextFile`.
- **File**: `Path`, `Name`, `ParentFolder`, `Drive`, `Size`, `Type`,
  `Attributes`, `DateLastModified` (etc.), `Delete`, `Copy`, `Move`,
  `OpenAsTextStream`.
- **Files / SubFolders**: `Count`, `Item(name)`, and `For Each`.

On non-Windows, `\` in paths is treated as a separator so ordinary scripts work.

## Automation & multimedia

AutoIt-style desktop automation, plus a cross-platform WAV combiner.

```vbscript
' --- Mouse (Windows) ---
Dim m : Set m = New Mouse
m.Move 400, 300              ' instant; add a 3rd arg for smooth motion, e.g. m.Move 400,300,5
m.Click "left"               ' Click([button],[x],[y],[count]); button = left/right/middle
m.Click "left", 400, 300, 2  ' move there, then double-click
m.Down "left" : m.Up "left"  ' press/release separately (drag)
m.Wheel "down", 3            ' scroll
directive.echo "cursor at " & m.X & "," & m.Y   ' also m.GetPos -> [x, y]

' --- Screen (Windows) ---
Dim s : Set s = New Screen
directive.echo Hex(s.PixelColor(100, 100))   ' 0xRRGGBB, like AutoIt PixelGetColor
directive.echo s.Width & "x" & s.Height

' --- Sound playback (Windows) ---
Dim snd : Set snd = New Sound
snd.Play "chime.wav"          ' async by default (returns immediately)
snd.Play "chime.wav", True    ' True = synchronous (waits until finished)
snd.Stop                      ' stop async playback
snd.Beep 1000, 250            ' frequency Hz, duration ms

' --- Clipboard (Windows; AutoIt-style flat functions) ---
ClipPut "hello from Directive"        ' returns True on success
directive.echo ClipGet()              ' reads clipboard text back

' --- WavWriter: concatenate WAV files into one (all platforms) ---
Dim w : Set w = New WavWriter
w.Create "out.wav"
w.Append "a.wav"              ' first file sets the format
w.Append "b.wav"              ' must match channels / sample-rate / bit-depth
directive.echo "combined length: " & w.Duration & " sec"
w.Close                        ' writes the RIFF/WAVE header and closes
```

`WavWriter` streams each input's PCM straight into the output and patches the
header on `Close`, a bit like a SAPI file stream but by direct concatenation
(same-format uncompressed PCM only — mismatched files raise an error rather than
silently corrupting the output).

> **Platform note.** `Mouse`, `Screen`, `Sound`, and the `ClipPut`/`ClipGet`
> clipboard functions are **Windows-only** (Win32 input, GDI, `winmm`, and
> clipboard APIs); on Linux/macOS they raise a clear "requires the Windows build"
> error. `WavWriter` is pure file I/O and works everywhere. See the verification
> note under *Known limitations*.

## Language coverage

Everything VBScript-ish you'd expect: `Option Explicit`, `Dim`/`Public`/`Private`,
`Const`, all operators (`Mod`, `\`, `^`, `&`, `Is`, `And/Or/Not/Xor/Eqv/Imp`),
`If/ElseIf/Else` (block and single-line), `For/Next` (+`Step`), `For Each`,
`Do/Loop` (all forms), `While/Wend`, `Select Case` (lists, `Is`, `To` ranges),
`With`, `Exit`, `Sub`/`Function` with `ByRef`/`ByVal`, `Class` with
`Property Get/Let/Set`, `Class_Initialize`/`Class_Terminate`, `On Error Resume Next`/`Err`, and
~85 built-in functions (strings, conversion, math, dates, arrays, type checks,
plus `Eval`, `GetRef`, `Erase`, `MonthName`/`WeekdayName`, `FormatNumber`/
`FormatCurrency`/`FormatPercent`, `DatePart`, `ScriptEngine`, `ClipPut`/`ClipGet`).
The `vb*` constants (`vbCrLf`, `vbTab`, `vbTextCompare`, …) are kept.

## Examples

See `examples/`. Highlights: `05_list.directive`, `06_dictionary.directive`,
`07_files.directive` (File/Folder objects), `09_console_io.directive`,
`11_automation.directive` (Mouse/Screen/Sound + WavWriter),
`12_include.directive` (native `Include` pulling in `lib/strutil.directive`).

```sh
./directive examples/05_list.directive
```

## Extending

- **New built-in function** → add a lambda in `Interpreter::registerBuiltins()`.
- **New `New`-able object** → implement the `IObject` interface and add a case in
  `Interpreter::evalNew()`.

Pipeline is `Lexer` → recursive-descent `Parser` → tree-walking `Interpreter`, all
in one file.

## Known limitations

- **Objects — `Set` vs plain assignment.** `Set x = obj` works as in VBScript.
  Directive is *more lenient* than VBScript about the missing-`Set` case: plain
  `x = obj` also stores the object reference, whereas strict VBScript would try to
  read the object's default property. So correct `Set`-using code behaves
  identically; only code that deliberately relied on VBScript's stricter rule
  differs. (Not a bug — a superset.)
- **Arrays copy on assignment**, exactly like VBScript — `b = a` gives `b` its own
  copy, so changing `b(0)` doesn't touch `a(0)`. The copy is shallow: if elements
  are objects, the references are copied (the objects aren't cloned), which is also
  how VBScript behaves.
- **`Class_Terminate` runs deterministically** — when you `Set x = Nothing`,
  reassign the last reference, a local goes out of scope at end of a `Sub`/
  `Function`, or the script ends. The one case where it can't run is an unbroken
  **reference cycle** (objects holding `Set` references to each other): the cycle
  is never collected, so it neither terminates nor frees — the same behavior COM/
  VBScript had. Break cycles (`Set a.other = Nothing`) if you need the terminators.
- Dates are stored as OLE serial doubles, so `TypeName` reports a date-valued
  number as a number. `Rnd` isn't bit-identical to VBScript's.
- File/Folder `DateCreated`/`DateLastAccessed` report the last-modified time
  (portable `std::filesystem` exposes only that); `Attributes` are approximated on
  non-Windows.
- **Verification status of the automation objects.** The `WavWriter` combiner and
  the `ClipPut`/`ClipGet` clipboard functions are tested (both natively where
  applicable and through the Windows `.exe`). `Mouse`, `Screen`, and `Sound` are
  **compile- and link-verified for Windows but not runtime-tested** — they were
  built in an environment with no display, mouse, or audio device, so give them a
  real-Windows smoke test before relying on them.

## Safety & robustness (read before trusting it with real work)

Directive is a capable hobby interpreter that was created using Claude.ai,
tested with the bundled examples, a set of targeted stress tests, and an AddressSanitizer/UBSan/LeakSanitizer pass.
It has **not** been fuzzed at scale or security-audited to production-runtime standards.
Concretely:

- **It is not a security sandbox.** Like VBScript, it deliberately exposes powerful
  capabilities — `FileSystem` (read/write/delete files), `CreateObject` (arbitrary
  COM on Windows, including shells), OS automation (`Mouse`/`Screen`/clipboard),
  and `Include` (reads arbitrary files). A `.directive` script can do anything the
  user running it can do, so **treat an untrusted script exactly like an untrusted
  `.exe` — don't run it.** (Directive is not wired into browsers, Office, or email
  the way VBScript was, which is what made VBScript such a widespread malware
  vector; but the language itself is just as capable.)
- **Reference cycles leak.** Objects are reference-counted; if two objects hold
  `Set` references to each other (`a.other = b : b.other = a`) and you drop all
  outside references, that cluster won't be freed until the process exits. This is
  the same limitation COM/VBScript had. Harmless for short-lived scripts (the OS
  reclaims everything on exit); avoid unbroken cycles in long-running processes.
- **Recursion is bounded** at ~2000 nested calls; deeper raises a catchable
  *Out of stack space* (code `800A001C`) instead of crashing. The shipped Windows
  builds also link with a 16 MB stack. Normal recursion is unaffected.
- **Use-after-close is safe** (writing to a closed `TextStream` is a no-op, not a
  crash), and deeply nested `Select`/`For`/`Do`/`If`/`With` and class bodies behave
  correctly.
- Bottom line: fine for personal automation, build scripts, and trusted internal
  tooling — the niches where you'd have reached for `.vbs`. Don't put it in front of
  untrusted input or use it as a security boundary.

## License

MIT.
