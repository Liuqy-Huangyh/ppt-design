using System;
using System.Diagnostics;
using System.IO;

class PPTStudioLauncher
{
    [STAThread]
    static void Main(string[] args)
    {
        string root = @"C:\Users\91125\Desktop\PPT";
        string script = Path.Combine(root, "scripts", "launch-ppt-studio.ps1");
        var psi = new ProcessStartInfo("powershell.exe",
            "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\"");
        psi.UseShellExecute = false;
        psi.WorkingDirectory = root;
        psi.CreateNoWindow = false;
        Process.Start(psi);
    }
}
