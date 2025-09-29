using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using Microsoft.Extensions.Logging;
using FlaUIApplication = FlaUI.Core.Application;

class Program
{
    static ILogger logger = LoggerFactory.Create(b => b.AddConsole()).CreateLogger("Runner");

    static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: AutomationRunner.exe <PathToExe OR 'Attach:WindowTitlePart'> <PathToCsv>");
            return 2;
        }

        var target = args[0];
        var csvPath = args[1];

        using var app = LaunchOrAttach(target);
        using var automation = new UIA3Automation();

        var mainWindows = app.GetAllTopLevelWindows(automation);
        if (mainWindows.Length == 0)
        {
            logger.LogError("No top-level windows found for target.");
            return 3;
        }

        var steps = LoadSteps(csvPath);
        foreach (var s in steps)
        {
            try
            {
                ExecuteStep(app, automation, s);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Step failed: {Action} [{By}={Selector}] Value='{Value}'", s.Action, s.By, s.Selector, s.Value);
                return 4; // fail fast (you can change to continue-on-error)
            }
        }

        return 0;
    }

    static FlaUIApplication LaunchOrAttach(string target)
    {
        if (target.StartsWith("Attach:", StringComparison.OrdinalIgnoreCase))
        {
            var titlePart = target.Substring("Attach:".Length);
            // Try attach by window title:
            var procs = System.Diagnostics.Process.GetProcesses()
                .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle) && p.MainWindowTitle.Contains(titlePart, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (procs.Length == 0)
                throw new InvalidOperationException($"No process with main window containing '{titlePart}'");

            return FlaUIApplication.Attach(procs[0]);
        }
        else
        {
            return FlaUIApplication.Launch(target);
        }
    }

    static List<Step> LoadSteps(string csvPath)
    {
        var cfg = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true,
            IgnoreBlankLines = true,
            TrimOptions = TrimOptions.Trim
        };
        using var sr = new StreamReader(csvPath);
        using var csv = new CsvReader(sr, cfg);
        return csv.GetRecords<Step>().ToList();
    }

    static void ExecuteStep(FlaUIApplication app, UIA3Automation automation, Step s)
    {
        var timeout = TimeSpan.FromMilliseconds(s.TimeoutMs.GetValueOrDefault(4000));

        switch (s.Action?.Trim().ToLowerInvariant())
        {
            case "focuswindow":
                {
                    var win = FindWindow(app, automation, s.Selector ?? s.Window, timeout);
                    win?.Focus();
                    break;
                }
            case "waitfor":
                {
                    FindElement(app, automation, s.Window, s.ControlType, s.By, s.Selector, timeout, required: true);
                    break;
                }
            case "click":
                {
                    var el = FindElement(app, automation, s.Window, s.ControlType, s.By, s.Selector, timeout, required: true);
                    var button = el.AsButton();
                    if (button != null)
                    {
                        button.Invoke();
                    }
                    else
                    {
                        el.Click();
                    }
                    break;
                }
            case "settext":
                {
                    var el = FindElement(app, automation, s.Window, s.ControlType, s.By, s.Selector, timeout, required: true);
                    var edit = el.AsTextBox();
                    if (edit != null)
                    {
                        edit.Focus();
                        edit.Text = s.Value ?? string.Empty;
                    }
                    else
                    {
                        // generic fallback: set via keyboard
                        el.Focus();
                        SelectAllAndType(s.Value ?? "");
                    }
                    break;
                }
            case "sendkeys":
                {
                    var win = FindWindow(app, automation, s.Window, timeout);
                    win?.Focus();
                    Keyboard.Type(ConvertKeySequence(s.Selector ?? s.Value ?? ""));
                    break;
                }
            case "menu":
                {
                    var win = FindWindow(app, automation, s.Window, timeout);
                    if (win == null) throw new Exception("Window not found for menu action");
                    InvokeMenuPath(win, s.Selector ?? s.Value ?? "");
                    break;
                }
            case "sleep":
                {
                    Thread.Sleep(timeout);
                    break;
                }
            case "ifexists":
                {
                    var el = FindElement(app, automation, s.Window, s.ControlType, s.By, s.Selector, timeout, required: false);
                    // No-op; you can add conditional branching if you extend the DSL
                    break;
                }
            case "log":
                {
                    logger.LogInformation("LOG: {Value}", s.Value ?? s.Selector);
                    break;
                }
            default:
                throw new NotSupportedException($"Unknown action '{s.Action}'");
        }
    }

    static Window? FindWindow(FlaUIApplication app, UIA3Automation automation, string? titleOrPart, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            var wins = app.GetAllTopLevelWindows(automation);
            Window? match = null;
            if (string.IsNullOrWhiteSpace(titleOrPart))
                match = wins.FirstOrDefault();
            else
                match = wins.FirstOrDefault(w => (w.Title ?? "").Contains(titleOrPart, StringComparison.OrdinalIgnoreCase));
            if (match != null) return match;
            Thread.Sleep(100);
        }
        return null;
    }

    static AutomationElement FindElement(FlaUIApplication app, UIA3Automation automation,
        string? windowTitle, string? controlType, string? by, string? selector, TimeSpan timeout, bool required)
    {
        var win = FindWindow(app, automation, windowTitle, timeout) ?? throw new Exception($"Window '{windowTitle}' not found");
        var condition = BuildCondition(automation, controlType, by, selector);
        var deadline = DateTime.UtcNow + timeout;

        AutomationElement? el;

        while (DateTime.UtcNow < deadline)
        {
            el = condition == null
                ? win.FindFirstDescendant()
                : win.FindFirstDescendant(condition);

            if (el != null) return el;
            Thread.Sleep(100);
        }

        if (required) throw new Exception($"Element not found: ControlType={controlType}, By={by}, Selector={selector}");
        return null!;
    }
    static ConditionBase? BuildCondition(UIA3Automation automation, string? controlType, string? by, string? selector)
    {
        var ct = controlType?.Trim();
        var byKey = (by ?? "").Trim().ToLowerInvariant();
        var sel = selector ?? "";

        var pieces = new List<ConditionBase>();

        if (!string.IsNullOrEmpty(ct))
        {
            if (Enum.TryParse(ct, out ControlType type))
                pieces.Add(automation.ConditionFactory.ByControlType(type));
        }

        var cf = automation.ConditionFactory;
        switch (byKey)
        {
            case "automationid": pieces.Add(cf.ByAutomationId(sel)); break;
            case "name": pieces.Add(cf.ByName(sel)); break;
            case "classname": pieces.Add(cf.ByClassName(sel)); break;
            case "title": pieces.Add(cf.ByName(sel)); break;
            case "path":
                // handled elsewhere
                break;
            case "keys":
                break;
        }

        if (pieces.Count == 0)
        {
            return null; // no condition means search everything
        }
        else if (pieces.Count == 1)
        {
            return pieces[0];
        }
        else
        {
            return new AndCondition(pieces.ToArray());
        }
    }

    static void SelectAllAndType(string text)
    {
        Keyboard.Press(VirtualKeyShort.CONTROL);
        Keyboard.Type("a");
        Keyboard.Release(VirtualKeyShort.CONTROL);
        Keyboard.Type(text);
    }

    static void InvokeMenuPath(Window win, string path)
    {
        // Example path: "File>Export CSV"
        var parts = path.Split('>').Select(p => p.Trim()).ToArray();
        var current = (AutomationElement)win;
        foreach (var part in parts)
        {
            var item = current.FindFirstDescendant(cf => cf.ByControlType(ControlType.MenuItem).And(cf.ByName(part)))
                      ?? current.FindFirstDescendant(cf => cf.ByName(part));
            if (item == null) throw new Exception($"Menu item '{part}' not found in '{path}'");
            var mi = item.AsMenuItem();
            if (mi != null) { mi.Focus(); mi.Invoke(); }
            else { item.Click(); }
            current = win; // many menus collapse; re-scope to window
            Thread.Sleep(150);
        }
    }

    // Simple literal pass-through; you can expand to parse things like {TAB}, {ENTER}, etc.
    static string ConvertKeySequence(string seq) => seq;
}
