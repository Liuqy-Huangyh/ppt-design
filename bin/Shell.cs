using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

class Shell : Form
{
    const string Root = @"C:\Users\91125\Desktop\PPT";
    const string Url = "http://127.0.0.1:8765/";
    WebView2 wv = new WebView2();
    Process server;

    static bool Healthy()
    {
        try
        {
            using (var wc = new WebClient())
            {
                wc.Proxy = null;
                var body = wc.DownloadString("http://127.0.0.1:8765/api/health");
                return body.Contains("\"ok\"");
            }
        }
        catch { return false; }
    }

    public Shell()
    {
        Text = "PPT Studio · 案例驱动的演示工作台";
        ClientSize = new Size(1440, 920);
        MinimumSize = new Size(1024, 720);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(247, 243, 234);
        try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
        wv.Dock = DockStyle.Fill;
        Controls.Add(wv);
        Load += OnLoad;
        FormClosing += OnClosing;
    }

    async void OnLoad(object sender, EventArgs e)
    {
        if (!Healthy())
        {
            var psi = new ProcessStartInfo(Path.Combine(Root, "..\\..\\..\\..\\..\\..\\..\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe"));
            psi = new ProcessStartInfo(@"C:\Users\91125\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe",
                "-m server.app --host 127.0.0.1 --port 8765");
            psi.WorkingDirectory = Root;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.EnvironmentVariables["PYTHONPATH"] = Root;
            server = Process.Start(psi);
            for (int i = 0; i < 60 && !Healthy(); i++) System.Threading.Thread.Sleep(250);
        }
        await wv.EnsureCoreWebView2Async(null);
        wv.CoreWebView2.Navigate(Url);
    }

    void OnClosing(object sender, FormClosingEventArgs e)
    {
        try { if (server != null && !server.HasExited) server.Kill(); } catch { }
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.Run(new Shell());
    }
}
