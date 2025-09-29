
# Automation Runner – Sample Script & Usage

This repo contains a **sample `script.csv`** demonstrating all supported actions of your C# + FlaUI automation runner, and a quick guide to run and extend it.

---

## Files

- `script.csv` – a ready-to-run script showing **FocusWindow, WaitFor, SetText, Click, Menu, Sleep, SendKeys, IfExists, Log**.
- `README.md` – this guide.

> Note: The runner executes **row-by-row**. Each row is one action.

---

## CSV Schema

Columns (header is required):

| Column       | Meaning                                                                                         |
|--------------|--------------------------------------------------------------------------------------------------|
| `Action`     | One of: `FocusWindow`, `WaitFor`, `Click`, `SetText`, `SendKeys`, `Menu`, `Sleep`, `IfExists`, `Log` |
| `Window`     | Window title (or part). Used to scope searches to the correct top-level window.                 |
| `ControlType`| Optional hint: `Edit, Button, ComboBox, DataGrid, MenuItem, Window, Pane, ...`                   |
| `By`         | Selector type: `AutomationId`, `Name`, `ClassName`, `Title`, `Path`, `Keys`                      |
| `Selector`   | The selector value (e.g., `txtOrderId`, `Search`, `Orders Grid`, `File>Export CSV`, literal keys)|
| `Value`      | Payload for the action (e.g., text for `SetText`).                                               |
| `TimeoutMs`  | Optional per-row timeout/wait override. For `Sleep`, it's the sleep duration.                    |

### Action semantics

- **FocusWindow**: Brings a window to the foreground. Uses `Selector` (preferred) or `Window` to match by title substring.
- **WaitFor**: Repeatedly searches for the element until `TimeoutMs` elapses or it is found.
- **Click**: Clicks the target element. If it's a `Button`, uses `Invoke()`, else falls back to `Click()`.
- **SetText**: Sets text into an edit control. If direct value setting fails, it focuses the element, Select-All, and types.
- **SendKeys**: Sends literal text to the focused control/window. *(Special keys like `{ENTER}` are **not** parsed by default; extend `ConvertKeySequence` if you need token support.)*
- **Menu**: Invokes a menu path like `File>Export CSV`. Each segment is looked up by `Name` under the active window.
- **Sleep**: Pauses the script for `TimeoutMs` milliseconds.
- **IfExists**: Checks for the element but does **not** fail if not found (no-op). Useful for optional dialogs.
- **Log**: Prints a message to the console logger (`Selector` or `Value`).

---

## Sample `script.csv`

> You can open this file to see all actions in context. Adjust `Legacy App - Orders`, selectors, and values to your app.

Columns: `Action,Window,ControlType,By,Selector,Value,TimeoutMs`

```
FocusWindow,,Window,Title,Legacy App - Orders,,5000
WaitFor,Legacy App - Orders,Edit,AutomationId,txtOrderId,,8000
SetText,Legacy App - Orders,Edit,AutomationId,txtOrderId,12345,
Click,Legacy App - Orders,Button,Name,Search,,0
WaitFor,Legacy App - Orders,DataGrid,Name,Orders Grid,,10000
IfExists,Legacy App - Orders,Button,Name,OK,,2000
Click,Legacy App - Orders,Button,Name,OK,,
Menu,Legacy App - Orders,MenuItem,Path,File>Export CSV,,
Sleep,,,,,,,1500
SendKeys,Legacy App - Orders,,Keys,Export_Complete,,
Log,,,,,Finished script successfully,
```

---

## Running the Runner

### Launch the target EXE
```powershell
AutomationRunner.exe "C:\Path\To\YourLegacyApp.exe" "C:\scripts\script.csv"
```

### Or attach to a running window (match by title substring)
```powershell
AutomationRunner.exe "Attach:Legacy App - Orders" "C:\scripts\script.csv"
```

> If the app runs **elevated (UAC)**, start the runner **as Administrator** too.

---

## Finding reliable selectors

Install one of these inspectors:
- **Microsoft Inspect** (Windows SDK)
- **FlaUInspect** (from FlaUI repo)

Prefer `AutomationId` where available. If not, combine `ControlType + Name`. Avoid index-based paths.

---

## Troubleshooting

- **`Application is ambiguous`**: Either remove `<UseWindowsForms>true</UseWindowsForms>` from the `.csproj`, or alias: `using FlaUIApplication = FlaUI.Core.Application;`.
- **`VirtualKeyShort does not exist`**: Add `using FlaUI.Core.WindowsAPI;` in `Program.cs`.
- **`?? cannot be applied to void`**: Replace `el.AsButton()?.Invoke() ?? el.Click();` with an `if` block.
- **`ConditionFactory.True / .And` not found**: Return `null` for “no condition” and use `new AndCondition(...)` to combine filters.
- **Element not found**: Add a preceding `WaitFor` with a longer `TimeoutMs`. Confirm selector via inspector.
- **SendKeys and special keys**: The sample sends text literally. To support tokens like `{ENTER}`, extend `ConvertKeySequence` in code to translate tokens into `Keyboard.Press/Release(...)` calls.

---

## Extending the DSL

Ideas you can add quickly:
- `AssertText` for verifying labels or cell values.
- `Screenshot` to save a PNG when a step fails.
- `MouseClick(x,y)` for coordinate-based fallback (for truly inaccessible UIs).
- `ComboSelect` to choose items by text.
- `Retry` or `ContinueOnError` columns for resilient flows.

---

## Safety & Stability Tips

- Match the **bitness** when using inspector tools against very old 32‑bit apps.
- Use **explicit waits** (`WaitFor`) between navigation/dialog steps.
- Run the runner under the **same elevation** as the target app.
- Disable screensavers/lock during unattended runs.

---

Happy automating! If you share a screenshot from Inspect or the specific window title and a couple of control properties, I can tailor the CSV for your app.
