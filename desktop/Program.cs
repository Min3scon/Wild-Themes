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

    // own GIF player: browsers play frames of 10 ms or less at 100 ms, so the scene runs at the web speed
    Image gif; int frame, frames; int[] delays = { 100 };
    readonly Timer anim = new();
    bool dark = Settings.Get("dark", "1") == "1";
    protected override void OnPaint(PaintEventArgs e)
    {
        if (gif == null) return;
        e.Graphics.DrawImage(gif, ClientRectangle);
        if (dark) using (var b = new SolidBrush(Color.FromArgb(110, 0, 0, 0))) e.Graphics.FillRectangle(b, ClientRectangle);
    }
    void Next()
    {
        if (gif == null || frames < 2) return;
        frame = (frame + 1) % frames;
        gif.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Time, frame);
        anim.Interval = delays[frame];
        Invalidate();
    }
    readonly Timer tick = new() { Interval = 250 };
    readonly NotifyIcon tray = new() { Visible = true, Text = "Wild Themes for Word" };
    readonly string dir = Path.Combine(AppContext.BaseDirectory, "scenes");
    IntPtr word = IntPtr.Zero; Rectangle last;

    public Scene()
    {
        FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual;
        BackColor = Color.Black; DoubleBuffered = true;
        anim.Tick += (_, _) => Next(); anim.Start();
        tray.Icon = SystemIcons.Application;
        var menu = new ContextMenuStrip();
        foreach (var f in Directory.GetFiles(dir, "*.gif").OrderBy(x => x)) { var id = Path.GetFileNameWithoutExtension(f); menu.Items.Add(id, null, (_, _) => Use(id)); }
        menu.Items.Add(new ToolStripSeparator());
        var dm = new ToolStripMenuItem("Dark mode") { Checked = dark, CheckOnClick = true };
        dm.CheckedChanged += (_, _) => { dark = dm.Checked; Settings.Set("dark", dark ? "1" : "0"); Invalidate(); };
        menu.Items.Add(dm);
        menu.Items.Add("Quit", null, (_, _) => Close());
        tray.ContextMenuStrip = menu;
        Use(Settings.Get("scene", "synthwave"));
        Listen();
        tick.Tick += (_, _) => Follow();
        tick.Start();
    }

    // click-through, never takes focus, no taskbar button
    protected override CreateParams CreateParams { get { var p = base.CreateParams; p.ExStyle |= 0x80000 | 0x20 | 0x80 | 0x08000000; return p; } }
    protected override bool ShowWithoutActivation => true;
    protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); SetLayeredWindowAttributes(Handle, 0, 255, 2); }

    void Use(string id, bool remember = true)
    {
        var f = Path.Combine(dir, id + ".gif");
        if (!File.Exists(f)) return;
        var old = gif; gif = Image.FromFile(f); old?.Dispose();
        frames = gif.GetFrameCount(System.Drawing.Imaging.FrameDimension.Time); frame = 0;
        try
        {
            var raw = gif.GetPropertyItem(0x5100).Value;   // frame delays, 1/100 s each
            delays = Enumerable.Range(0, frames).Select(i => { var d = BitConverter.ToInt32(raw, i * 4) * 10; return d <= 10 ? 100 : d; }).ToArray();
        }
        catch { delays = Enumerable.Repeat(100, Math.Max(1, frames)).ToArray(); }
        anim.Interval = delays[0];
        Invalidate();
        if (remember) Settings.Set("scene", id);
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)] static extern int CLSIDFromProgID(string p, out Guid g);
    [DllImport("oleaut32.dll")] static extern int GetActiveObject(ref Guid g, IntPtr r, [MarshalAs(UnmanagedType.IUnknown)] out object o);
    int n; string fileScene;
    // The open document's own theme (custom property WildTheme), so the sides match the page.
    void FromDocument()
    {
        try
        {
            CLSIDFromProgID("Word.Application", out var g);
            if (GetActiveObject(ref g, IntPtr.Zero, out var o) != 0) return;
            dynamic app = o;
            string id = app.ActiveDocument.CustomDocumentProperties.Item("WildTheme").Value;
            if (id != fileScene && File.Exists(Path.Combine(dir, id + ".gif"))) { fileScene = id; Use(id, false); }
            return;
        }
        catch { /* no document or no theme in it */ }
        if (fileScene != null) { fileScene = null; Use(Settings.Get("scene", "synthwave"), false); }   // back to the browser's theme
    }

    // The browser extension sends the theme chosen on the website or in the extension (this PC only).
    void Listen()
    {
        var l = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 47615);
        try { l.Start(); } catch { return; }   // another copy is already listening
        new System.Threading.Thread(() =>
        {
            while (true)
            {
                try
                {
                    using var c = l.AcceptTcpClient(); using var s = c.GetStream();
                    var buf = new byte[2048]; var line = Encoding.ASCII.GetString(buf, 0, s.Read(buf, 0, buf.Length)).Split('\n')[0];
                    var m = System.Text.RegularExpressions.Regex.Match(line, @"GET /theme\?id=([a-z0-9-]*)");
                    if (m.Success && File.Exists(Path.Combine(dir, m.Groups[1].Value + ".gif")))
                        BeginInvoke(() => { Settings.Set("scene", m.Groups[1].Value); if (fileScene == null) Use(m.Groups[1].Value); });
                    var ok = Encoding.ASCII.GetBytes("HTTP/1.1 204 No Content\r\nAccess-Control-Allow-Origin: *\r\nConnection: close\r\n\r\n");
                    s.Write(ok, 0, ok.Length);
                }
                catch { }
            }
        }) { IsBackground = true }.Start();
    }

    void Follow()
    {
        if (n++ % 8 == 0) FromDocument();
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
    public static void Set(string k, string v)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(file));
            var keep = File.Exists(file) ? File.ReadAllLines(file).Where(l => !l.StartsWith(k + "=")) : Enumerable.Empty<string>();
            File.WriteAllLines(file, keep.Append(k + "=" + v));
        }
        catch { }
    }
}
