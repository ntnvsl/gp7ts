using System.Diagnostics;
using System.Windows;

namespace gp7ts
{
  /// <summary>
  /// Interaction logic for App.xaml
  /// </summary>
  public partial class App : Application
  {
    protected override void OnStartup(StartupEventArgs e)
    {
      base.OnStartup(e);
      if (Process.GetProcessesByName("gp7ts").Length > 1)
        Shutdown();
    }
  }
}
