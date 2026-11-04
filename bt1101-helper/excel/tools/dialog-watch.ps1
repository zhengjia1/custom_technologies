# Watches one Excel process for modal dialogs (VBA compile errors, run-time errors,
# message boxes), appends each dialog's text to $LogFile and dismisses it so automation
# never hangs. The Excel window handle is read from $HwndFile once excel-task.wsf writes it.
# Stops when $StopFile exists.
param([string]$HwndFile, [string]$LogFile, [string]$StopFile)

Add-Type -TypeDefinition @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class DialogWatch {
  delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc, IntPtr lParam);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr lParam);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder text, int max);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);
  static string Cls(IntPtr h) { var b = new StringBuilder(256); GetClassName(h, b, 256); return b.ToString(); }
  static string Txt(IntPtr h) { var b = new StringBuilder(4096); GetWindowText(h, b, 4096); return b.ToString(); }
  public static uint ProcessOf(IntPtr hwnd) { uint pid; GetWindowThreadProcessId(hwnd, out pid); return pid; }
  // Returns the text of every visible dialog of the process and presses its dismiss button.
  public static List<string> Sweep(uint pid) {
    var found = new List<string>();
    EnumWindows((h, l) => {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p != pid || !IsWindowVisible(h) || Cls(h) != "#32770") return true;
      var text = new StringBuilder(Txt(h) + ": ");
      IntPtr chosen = IntPtr.Zero; int rank = 99;
      EnumChildWindows(h, (c, l2) => {
        string cls = Cls(c), t = Txt(c);
        if (cls == "Static" && t.Length > 0) text.Append(t.Replace("\r", " ").Replace("\n", " ")).Append(" ");
        if (cls == "Button") {
          string k = t.Replace("&", "");
          int r = k == "End" ? 0 : k == "OK" ? 1 : k == "No" ? 2 : k == "Cancel" ? 3 : 9;
          if (r < rank) { rank = r; chosen = c; }
        }
        return true;
      }, IntPtr.Zero);
      found.Add(text.ToString().Trim());
      if (chosen != IntPtr.Zero) PostMessage(chosen, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
      else PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@

$seen = @{}
$pid0 = 0
while (-not (Test-Path $StopFile)) {
  if ($pid0 -eq 0 -and (Test-Path $HwndFile)) {
    $raw = (Get-Content -Raw $HwndFile).Trim()
    if ($raw) { $pid0 = [DialogWatch]::ProcessOf([IntPtr][Int64]$raw) }
  }
  if ($pid0 -ne 0) {
    foreach ($text in [DialogWatch]::Sweep($pid0)) {
      if (-not $seen.ContainsKey($text)) { Add-Content -Path $LogFile -Value $text -Encoding UTF8; $seen[$text] = 1 }
    }
  }
  Start-Sleep -Milliseconds 150
}
