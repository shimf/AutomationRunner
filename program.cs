// Program.cs
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

// Alias to avoid ambiguity with System.Windows.Forms.Application
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

        try
        {
            using var app = LaunchOrAttach(target);
            using var automation = new UIA3Automation();

            // Wait up to 10s for any top-level window to surface
            var firstWin = WaitForAnyTopWindow(app, automation, TimeSpan.FromSeconds(10), logger);
            if (firstWin == null)
            {
                logger.LogError("No top-level windows found for target after waiting. " +
                                "Check elevation/localized titles. If using Calculator, try launching calc.exe or using the localized title (e.g., 'מחשבון').");
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
                    logger.LogError(ex, "Step failed: {Action} [{By}={Selector}] Value='{Value}'",
                        s.Action, s.By, s.Selector, s.Value);
                    return 4;
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Fatal error");
            return 1;
        }
    }

    static FlaUIApplication LaunchOrAttach(string target)
    {
        if (target.StartsWith("Attach:", StringComparison.OrdinalIgnoreCase))
        {
            var titlePart = target.Substring("Attach:".Length);
            var procs = System.Diagnostics.Process.GetProcesses()
                .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle) &&
                            p.MainWindowTitle.IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToArray();

            if (procs.Length == 0)
                throw new InvalidOperationException($"No process with main window containing '{titlePart}'");

            var app = FlaUIApplication.Attach(procs[0]);
            // In case main handle is not yet ready
            app.WaitWhileMainHandleIsMissing(TimeSpan.FromSeconds(10));
            return app;
        }
        else
        {
            var app = FlaUIApplication.Launch(target);
            app.WaitWhileMainHandleIsMissing(TimeSpan.FromSeconds(10));
            return app;
        }
    }

    static IReadOnlyList<Window> GetTopWindows(FlaUIApplication app, UIA3Automation automation)
    {
        try { return app.GetAllTopLevelWindows(automation); }
        catch { return Array.Empty<Window>(); }
    }

    static Window? WaitForAnyTopWindow(FlaUIApplication app, UIA3Automation automation, TimeSpan timeout, ILogger log)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            var wins = GetTopWindows(app, automation);
            if (wins.Count > 0) return wins[0];
            Thread.Sleep(150);
        }
        log.LogWarning("No top-level windows appeared within {Seconds}s.", timeout.TotalSeconds);
        return null;
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
                if (button != null) button.Invoke();
                else el.Click();
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
                    // Fallback: focus + select all + type
                    el.Focus();
                    SelectAllAndType(s.Value ?? "");
                }
                break;
            }
            case "sendkeys":
            {
                var win = FindWindow(app, automation, s.Window, timeout);
                win?.Focus();
                SendKeySequence(s.Selector ?? s.Value ?? "");
                break;
            }
            case "menu":
            {
                var win = FindWindow(app, automation, s.Window, timeout)
                          ?? throw new Exception("Window not found for menu action");
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
                _ = FindElement(app, automation, s.Window, s.ControlType, s.By, s.Selector, timeout, required: false);
                break;
            }
            case "log":
            {
                logger.LogInformation("LOG: {Message}", s.Value ?? s.Selector);
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
                match = wins.FirstOrDefault(w => (w.Title ?? "").IndexOf(titleOrPart, StringComparison.OrdinalIgnoreCase) >= 0);
            if (match != null) return match;
            Thread.Sleep(100);
        }
        return null;
    }

    static AutomationElement FindElement(
        FlaUIApplication app, UIA3Automation automation,
        string? windowTitle, string? controlType, string? by, string? selector,
        TimeSpan timeout, bool required)
    {
        var win = FindWindow(app, automation, windowTitle, timeout) ?? throw new Exception($"Window '{windowTitle}' not found");
        var condition = BuildCondition(automation, controlType, by, selector);
        var deadline = DateTime.UtcNow + timeout;

        AutomationElement? el = null;
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
        var cf = automation.ConditionFactory;

        if (!string.IsNullOrEmpty(ct) && Enum.TryParse(ct, out ControlType type))
            pieces.Add(cf.ByControlType(type));

        switch (byKey)
        {
            case "automationid": pieces.Add(cf.ByAutomationId(sel)); break;
            case "name":         pieces.Add(cf.ByName(sel)); break;
            case "classname":    pieces.Add(cf.ByClassName(sel)); break;
            case "title":        pieces.Add(cf.ByName(sel)); break;
            // "path" and "keys" are handled in their specific actions
        }

        if (pieces.Count == 0) return null;
        if (pieces.Count == 1) return pieces[0];
        return new AndCondition(pieces.ToArray());
    }

    static void SelectAllAndType(string text)
    {
        Keyboard.Press(VirtualKeyShort.CONTROL);
        Keyboard.Type("a");
        Keyboard.Release(VirtualKeyShort.CONTROL);
        Keyboard.Type(text);
    }

    static void SendKeySequence(string seq)
    {
        // Minimal token support: Alt+F4 as "%{F4}"
        if (string.Equals(seq, "%{F4}", StringComparison.OrdinalIgnoreCase))
        {
            Keyboard.Press(VirtualKeyShort.ALT); // Alt
            Keyboard.Type(VirtualKeyShort.F4);
            Keyboard.Release(VirtualKeyShort.ALT);
            return;
        }

        // Default: literal typing
        Keyboard.Type(seq);
    }

    static void InvokeMenuPath(Window win, string path)
    {
        // Example path: "File>Export CSV"
        var parts = (path ?? "").Split(new[] { '>' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(p => p.Trim()).ToArray();
        if (parts.Length == 0) throw new ArgumentException("Menu path is empty");

        var current = (AutomationElement)win;
        foreach (var part in parts)
        {
            var item =
                current.FindFirstDescendant(cf => cf.ByControlType(ControlType.MenuItem).And(cf.ByName(part))) ??
                current.FindFirstDescendant(cf => cf.ByName(part));

            if (item == null)
                throw new Exception($"Menu item '{part}' not found in path '{path}'");

            var mi = item.AsMenuItem();
            if (mi != null) { mi.Focus(); mi.Invoke(); }
            else { item.Click(); }

            // Re-scope to window between clicks (menus can re-render)
            current = win;
            Thread.Sleep(150);
        }
    }
}
