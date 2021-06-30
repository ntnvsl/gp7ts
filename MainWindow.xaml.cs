using AutoHotkey.Interop;
using FlaUI.UIA3;
using NonInvasiveKeyboardHookLibrary;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using Application = FlaUI.Core.Application;

namespace gp7ts
{
  public partial class MainWindow : Window
  {
    #region user32.dll
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    #endregion

    private int pid = 0, npid;
    private FlaUI.Core.AutomationElements.Window window;
    private FlaUI.Core.AutomationElements.AutomationElement[] buttons;
    private double dur, last;
    private string pos, cpos, duration;
    private int i, j = 0, b = 0, max;
    private bool btn = false;
    private string[] nums;

    private static readonly double[] bd = { 0.125, 0.25, 0.375, 0.5, 0.625, 0.75, 0.875, 1.0, 1.125, 1.25, 1.375, 1.5, 1.625, 1.75, 1.875, 2.0, 2.125, 2.25, 2.375, 2.5, 2.625, 2.75, 2.875, 3.0, 3.125, 3.25, 3.375, 3.5, 3.625, 3.75, 3.875, 4.0, 4.25, 4.5, 4.75, 5.0, 5.25, 5.5, 5.75, 6.0, 6.25, 6.5, 6.75, 7.0, 7.25, 7.5, 7.75, 8.0, 8.5, 9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 12.5, 13.0, 13.5, 14.0, 14.5, 15.0, 15.5, 16.0, 17.0, 18.0, 19.0, 20.0, 21.0, 22.0, 23.0, 24.0, 25.0, 26.0, 27.0, 28.0, 29.0, 30.0, 31.0, 32.0, 34.0, 36.0, 38.0, 40.0, 42.0, 44.0, 46.0, 48.0, 50.0, 52.0, 54.0, 56.0, 58.0, 60.0, 62.0, 64.0, 68.0, 72.0, 76.0, 80.0, 84.0, 88.0, 92.0, 96.0, 100.0, 104.0, 108.0, 112.0, 116.0, 120.0, 124.0, 128.0 };
    private static readonly string[][] times = { new[] { "1.32" }, new[] { "1.16", "2.32" }, new[] { "3.32" }, new[] { "1.8", "2.16", "4.32" }, new[] { "5.32" }, new[] { "3.16", "6.32" }, new[] { "7.32" }, new[] { "1.4", "2.8", "4.16", "8.32" }, new[] { "9.32" }, new[] { "5.16", "10.32" }, new[] { "11.32" }, new[] { "3.8", "6.16", "12.32" }, new[] { "13.32" }, new[] { "7.16", "14.32" }, new[] { "15.32" }, new[] { "2.4", "4.8", "8.16", "1.2", "16.32" }, new[] { "17.32" }, new[] { "9.16", "18.32" }, new[] { "19.32" }, new[] { "5.8", "10.16", "20.32" }, new[] { "21.32" }, new[] { "11.16", "22.32" }, new[] { "23.32" }, new[] { "3.4", "6.8", "12.16", "24.32" }, new[] { "25.32" }, new[] { "13.16", "26.32" }, new[] { "27.32" }, new[] { "7.8", "14.16", "28.32" }, new[] { "29.32" }, new[] { "15.16", "30.32" }, new[] { "31.32" }, new[] { "4.4", "8.8", "16.16", "2.2", "1.1", "32.32" }, new[] { "17.16" }, new[] { "9.8", "18.16" }, new[] { "19.16" }, new[] { "5.4", "10.8", "20.16" }, new[] { "21.16" }, new[] { "11.8", "22.16" }, new[] { "23.16" }, new[] { "6.4", "12.8", "24.16", "3.2" }, new[] { "25.16" }, new[] { "13.8", "26.16" }, new[] { "27.16" }, new[] { "7.4", "14.8", "28.16" }, new[] { "29.16" }, new[] { "15.8", "30.16" }, new[] { "31.16" }, new[] { "8.4", "16.8", "4.2", "32.16", "2.1" }, new[] { "17.8" }, new[] { "9.4", "18.8" }, new[] { "19.8" }, new[] { "10.4", "20.8", "5.2" }, new[] { "21.8" }, new[] { "11.4", "22.8" }, new[] { "23.8" }, new[] { "12.4", "24.8", "6.2", "3.1" }, new[] { "25.8" }, new[] { "13.4", "26.8" }, new[] { "27.8" }, new[] { "14.4", "28.8", "7.2" }, new[] { "29.8" }, new[] { "15.4", "30.8" }, new[] { "31.8" }, new[] { "16.4", "32.8", "8.2", "4.1" }, new[] { "17.4" }, new[] { "18.4", "9.2" }, new[] { "19.4" }, new[] { "20.4", "10.2", "5.1" }, new[] { "21.4" }, new[] { "22.4", "11.2" }, new[] { "23.4" }, new[] { "24.4", "12.2", "6.1" }, new[] { "25.4" }, new[] { "26.4", "13.2" }, new[] { "27.4" }, new[] { "28.4", "14.2", "7.1" }, new[] { "29.4" }, new[] { "30.4", "15.2" }, new[] { "31.4" }, new[] { "32.4", "16.2", "8.1" }, new[] { "17.2" }, new[] { "18.2", "9.1" }, new[] { "19.2" }, new[] { "20.2", "10.1" }, new[] { "21.2" }, new[] { "22.2", "11.1" }, new[] { "23.2" }, new[] { "24.2", "12.1" }, new[] { "25.2" }, new[] { "26.2", "13.1" }, new[] { "27.2" }, new[] { "28.2", "14.1" }, new[] { "29.2" }, new[] { "30.2", "15.1" }, new[] { "31.2" }, new[] { "32.2", "16.1" }, new[] { "17.1" }, new[] { "18.1" }, new[] { "19.1" }, new[] { "20.1" }, new[] { "21.1" }, new[] { "22.1" }, new[] { "23.1" }, new[] { "24.1" }, new[] { "25.1" }, new[] { "26.1" }, new[] { "27.1" }, new[] { "28.1" }, new[] { "29.1" }, new[] { "30.1" }, new[] { "31.1" }, new[] { "32.1" } };
    private readonly NotifyIcon ni = new() { Icon = Properties.Resources.gp7icon, Visible = true };
    private readonly AutoHotkeyEngine ahk = new();
    private readonly KeyboardHookManager khm = new();

    public MainWindow()
    {
      if (Process.GetProcessesByName("gp7ts").Length > 1)
        Shutdown();

      khm.RegisterHotkey(0xC0, () => Changer());
      khm.Start();

      InitializeComponent();

      ni.Click += delegate (object sender, EventArgs args) { Shutdown(true); };
    }

    private void Shutdown(bool kill = false)
    {
      if (kill)
      {
        khm.UnregisterAll();
        khm.Stop();
      }
      Environment.Exit(0);
    }

    private void Changer()
    {
      lock (this)
      {
        GetWindowThreadProcessId(GetForegroundWindow(), out npid);
        if (Process.GetProcessById(npid).ProcessName != "GuitarPro7")
          return;

        if (npid != pid)
          window = Application.Attach(pid = npid).GetMainWindow(new UIA3Automation());
        buttons = window.FindAllDescendants(cf => cf.ByLocalizedControlType("button"));
        max = Convert.ToInt32(buttons.Length);
        for (b = 8; b < max; b++)
          if (Regex.IsMatch(duration = buttons[b].Patterns.LegacyIAccessible.Pattern.Name.Value, @"^\d\d?\d?\."))
          {
            btn = true;
            break;
          }

        if (btn)
          btn = false;
        else
          return;

        cpos = buttons[b - 1].Patterns.LegacyIAccessible.Pattern.Name.Value;
        dur = Convert.ToDouble(duration.Split(':')[0]);
        if (dur is < 0.125 or > 128.0)
          return;

        if ((i = Array.IndexOf(bd, dur)) == -1)
          while (bd[++i] < dur)
            continue;

        nums = times[i][last == bd[i] && times[i].Length > ++j && pos == cpos ? j : j = 0].Split('.');
        last = bd[i];
        pos = cpos;

        ahk.ExecRaw($@"
          send ^t{{Tab 8}}{{Home}}
          if {nums[0].Substring(0, 1)} <= 2
            send {{End}}
          send {nums[0]}{{Tab}}{{End}}
          if {nums[1].Substring(0, 1)} = 3
            send {{Home}}
          send {nums[1]}{{Enter}}
        ");
      }
    }
  }
}