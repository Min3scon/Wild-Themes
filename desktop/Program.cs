// Wild Themes for desktop Word. Word's ribbon and the grey canvas around the page can't be themed
// from a file or an add-in, so this small app lays the animated scene over the canvas only: a
// click-through window owned by Word, cut to the canvas left and right of the page (found by
// looking at what Word draws). Word itself is never changed; quit from the tray icon.
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new Scene());
    }
}

class Scene : Form
{
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
    [DllImport("user32.dll")] static extern IntPtr SetWindowLongPtr(IntPtr h, int i, IntPtr v);
    [DllImport("user32.dll")] static extern bool SetLayeredWindowAttributes(IntPtr h, uint key, byte alpha, uint flags);
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
    const int GWLP_HWNDPARENT = -8;

    readonly PictureBox pic = new() { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.StretchImage };
    readonly Timer tick = new() { Interval = 250 };
    readonly NotifyIcon tray = new() { Visible = true, Text = "Wild Themes for Word" };
    readonly string dir = Path.Combine(AppContext.BaseDirectory, "scenes");
    IntPtr word = IntPtr.Zero; Rectangle last;

    public Scene()
    {
        FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual;
        BackColor = Color.Black; Controls.Add(pic);
        tray.Icon = SystemIcons.Application;
        var menu = new ContextMenuStrip();
        foreach (var f in Directory.GetFiles(dir, "*.gif").OrderBy(x => x)) { var id = Path.GetFileNameWithoutExtension(f); menu.Items.Add(id, null, (_, _) => Use(id)); }
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Quit", null, (_, _) => Close());
        tray.ContextMenuStrip = menu;
        Use(Settings.Get("scene", "synthwave"));
        tick.Tick += (_, _) => Follow();
        tick.Start();
    }

    // click-through, never takes focus, no taskbar button
    protected override CreateParams CreateParams { get { var p = base.CreateParams; p.ExStyle |= 0x80000 | 0x20 | 0x80 | 0x08000000; return p; } }
    protected override bool ShowWithoutActivation => true;
    protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); SetLayeredWindowAttributes(Handle, 0, 255, 2); }

    void Use(string id)
    {
        var f = Path.Combine(dir, id + ".gif");
        if (!File.Exists(f)) return;
        pic.Image = Image.FromFile(f);
        Settings.Set("scene", id);
    }

    void Follow()
    {
        var fg = GetForegroundWindow(); var cls = new StringBuilder(32); GetClassName(fg, cls, 32);
        if (cls.ToString() == "OpusApp" && fg != word) Attach(fg);                       // whichever Word window is in front
        if (word == IntPtr.Zero || !IsWindow(word)) { word = FindWindow("OpusApp", null); if (word != IntPtr.Zero) Attach(word); }
        if (word == IntPtr.Zero || IsIconic(word) || !IsWindowVisible(word) || !GetWindowRect(word, out var r)) { if (Visible) Hide(); return; }
        var b = new Rectangle(r.L, r.T, r.R - r.L, r.B - r.T);
        var region = Canvas(b.Width, b.Height);
        if (region == null) { if (Visible) Hide(); return; }
        if (b != last) { Bounds = b; last = b; }
        Region = region;
        if (!Visible) Show();
    }

    void Attach(IntPtr w)
    {
        word = w;
        SetWindowLongPtr(Handle, GWLP_HWNDPARENT, w);   // owned by Word: stays just above it, moves with it in the stack
    }

    // Find Word's canvas (the flat colour around the page) in what Word draws, and return the
    // strips left and right of the page. Null when it can't tell (then nothing is drawn).
    Region Canvas(int w, int h)
    {
        if (w < 400 || h < 300) return null;
        using var bmp = new Bitmap(w, h);
        using (var g = Graphics.FromImage(bmp)) { var dc = g.GetHdc(); PrintWindow(word, dc, 2); g.ReleaseHdc(dc); }
        int sx = w / 16, y = h * 2 / 3;
        var c = bmp.GetPixel(sx, y);
        bool Same(int x, int yy) { var p = bmp.GetPixel(x, yy); return Math.Abs(p.R - c.R) + Math.Abs(p.G - c.G) + Math.Abs(p.B - c.B) < 12; }
        // only the canvas joined to the sample point: stops at the ribbon, status bar and scroll bar
        int top = y; while (top > 0 && Same(sx, top - 1)) top--;
        int bottom = y; while (bottom < h - 1 && Same(sx, bottom + 1)) bottom++;
        int left = sx; while (left > 0 && Same(left - 1, y)) left--;
        int pageL = sx; while (pageL < w - 1 && Same(pageL, y)) pageL++;
        int pageR = pageL; while (pageR < w - 1 && !Same(pageR, y)) pageR++;
        int right = pageR; while (right < w - 1 && Same(right + 1, y)) right++;
        pageR--;
        if (bottom - top < h / 3 || pageR - pageL < w / 8) return null;   // not a page view
        var reg = new Region(new Rectangle(left, top, pageL - left, bottom - top + 1));
        reg.Union(new Rectangle(pageR + 1, top, right - pageR, bottom - top + 1));
        return reg;
    }

    protected override void OnFormClosed(FormClosedEventArgs e) { tray.Visible = false; base.OnFormClosed(e); }
}

static class Settings
{
    static readonly string file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WildThemes", "desktop.txt");
    public static string Get(string k, string d) { try { foreach (var l in File.ReadAllLines(file)) if (l.StartsWith(k + "=")) return l[(k.Length + 1)..]; } catch { } return d; }
    public static void Set(string k, string v) { try { Directory.CreateDirectory(Path.GetDirectoryName(file)); File.WriteAllText(file, k + "=" + v); } catch { } }
}
