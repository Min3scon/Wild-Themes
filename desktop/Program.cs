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
        // one copy at a time; after an update the new copy waits for the old one to quit
        using var one = new System.Threading.Mutex(false, "WildThemesDesktop");
        try { if (!one.WaitOne(TimeSpan.FromSeconds(15))) return; } catch (System.Threading.AbandonedMutexException) { }
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
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
    const int GWLP_HWNDPARENT = -8;

    // The extension's own CSS scene in WebView2 (Edge), so the grey space matches Word on the web 1:1.
    readonly Microsoft.Web.WebView2.WinForms.WebView2 web = new() { Dock = DockStyle.Fill, DefaultBackgroundColor = Color.Black };
    bool dark = Settings.Get("dark", "1") == "1", ready; string scene;
    bool animated = Settings.Get("animated", "1") == "1", sparks = Settings.Get("sparks", "1") == "1";
    readonly Sparks spark = new();
    readonly Timer tick = new() { Interval = 250 };
    readonly NotifyIcon tray = new() { Visible = true, Text = "Wild Themes for Word" };
    readonly string dir = Path.Combine(AppContext.BaseDirectory, "scenes");
    IntPtr word = IntPtr.Zero; Rectangle last; bool isWord;

    public Scene()
    {
        FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual;
        BackColor = Color.Black; Controls.Add(web);
        tray.Icon = SystemIcons.Application;
        var menu = new ContextMenuStrip();
        var themes = new ToolStripMenuItem("Theme");
        menu.Items.Add(themes);
        void List() { themes.DropDownItems.Clear(); foreach (var (id, t) in Themes.All.OrderBy(x => x.Value.name)) themes.DropDownItems.Add(t.name, null, (_, _) => Use(id)); }
        try { Themes.Load(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "themes.js"))); } catch { }
        List();
        Updates.Themes(() => BeginInvoke(() => { List(); spark.Theme(scene); }));
        Updates.App();
        menu.Items.Add(new ToolStripSeparator());
        var dm = new ToolStripMenuItem("Dark mode") { Checked = dark, CheckOnClick = true };
        dm.CheckedChanged += (_, _) => { dark = dm.Checked; Settings.Set("dark", dark ? "1" : "0"); Play(scene); };
        menu.Items.Add(dm);
        var an = new ToolStripMenuItem("Animated") { Checked = animated, CheckOnClick = true };
        an.CheckedChanged += (_, _) => { animated = an.Checked; Settings.Set("animated", animated ? "1" : "0"); Play(scene); };
        menu.Items.Add(an);
        var sp = new ToolStripMenuItem("Sparks when typing") { Checked = sparks, CheckOnClick = true };
        sp.CheckedChanged += (_, _) => { sparks = sp.Checked; Settings.Set("sparks", sparks ? "1" : "0"); };
        menu.Items.Add(sp);
        // runs from sign-in by default, so the scene is there whenever Word opens
        var run = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
        if (Settings.Get("autostart", "1") == "1") run.SetValue("WildThemes", $"\"{Environment.ProcessPath}\"");
        var st = new ToolStripMenuItem("Start with Windows") { Checked = Settings.Get("autostart", "1") == "1", CheckOnClick = true };
        st.CheckedChanged += (_, _) => { Settings.Set("autostart", st.Checked ? "1" : "0"); if (st.Checked) run.SetValue("WildThemes", $"\"{Environment.ProcessPath}\""); else run.DeleteValue("WildThemes", false); };
        menu.Items.Add(st);
        menu.Items.Add("Quit", null, (_, _) => Close());
        tray.ContextMenuStrip = menu;
        Init();
        Listen();
        tick.Tick += (_, _) => Follow();
        tick.Start();
    }

    async void Init()
    {
        var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null,
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WildThemes", "webview"));
        await web.EnsureCoreWebView2Async(env);
        web.CoreWebView2.SetVirtualHostNameToFolderMapping("wildthemes.local", AppContext.BaseDirectory, Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);
        web.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        // the mouse wheel over the scene scrolls Word (its scroll bar is hidden so the scene fills that strip)
        web.CoreWebView2.WebMessageReceived += (_, e) => { if (!isWord) return; try { int d = int.Parse(e.TryGetWebMessageAsString()); dynamic a = WordApp(); if (d > 0) a.ActiveWindow.SmallScroll(d); else a.ActiveWindow.SmallScroll(0, -d); } catch { } };
        ready = true;
        Use(Settings.Get("scene", "synthwave"));
    }

    // never takes focus, no taskbar button
    protected override CreateParams CreateParams { get { var p = base.CreateParams; p.ExStyle |= 0x80 | 0x08000000; return p; } }
    protected override bool ShowWithoutActivation => true;

    bool Known(string id) => id != null && (Themes.All.ContainsKey(id) || File.Exists(Path.Combine(dir, id + ".css")));

    void Use(string id, bool remember = true)
    {
        if (!Known(id)) return;
        if (remember) Settings.Set("scene", id);
        Play(id);
    }
    void Play(string id)
    {
        scene = id; spark.Theme(id);
        if (ready && id != null) web.CoreWebView2.Navigate($"https://wildthemes.local/scene.html?id={id}&dark={(dark ? 1 : 0)}&anim={(animated ? 1 : 0)}");
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)] static extern int CLSIDFromProgID(string p, out Guid g);
    [DllImport("oleaut32.dll")] static extern int GetActiveObject(ref Guid g, IntPtr r, [MarshalAs(UnmanagedType.IUnknown)] out object o);
    int n; string fileScene;
    static dynamic WordApp() { CLSIDFromProgID("Word.Application", out var g); return GetActiveObject(ref g, IntPtr.Zero, out var o) == 0 ? o : throw new InvalidOperationException(); }
    // The open document's own theme (custom property WildTheme), so the sides match the page.
    void FromDocument()
    {
        try
        {
            CLSIDFromProgID("Word.Application", out var g);
            if (GetActiveObject(ref g, IntPtr.Zero, out var o) != 0) return;
            dynamic app = o;
            string id = app.ActiveDocument.CustomDocumentProperties.Item("WildTheme").Value;
            if (id != fileScene && Known(id)) { fileScene = id; Use(id, false); }
            return;
        }
        catch { /* no document or no theme in it */ }
        if (fileScene != null) { fileScene = null; Use(Settings.Get("scene", "synthwave"), false); }   // back to the browser's theme
    }

    // The browser extension sends the theme chosen on the website or in the extension (this PC only).
    void Listen()
    {
        var l = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 47615);
        for (int i = 0; ; i++) { try { l.Start(); break; } catch { if (i > 20) return; System.Threading.Thread.Sleep(500); } }   // port frees when an old copy quits
        new System.Threading.Thread(() =>
        {
            while (true)
            {
                try
                {
                    using var c = l.AcceptTcpClient(); using var s = c.GetStream();
                    var buf = new byte[2048]; var line = Encoding.ASCII.GetString(buf, 0, s.Read(buf, 0, buf.Length)).Split('\n')[0];
                    var m = System.Text.RegularExpressions.Regex.Match(line, @"GET /theme\?id=([a-z0-9-]*)");
                    if (m.Success && Known(m.Groups[1].Value))
                        BeginInvoke(() => { Settings.Set("scene", m.Groups[1].Value); if (fileScene == null) Use(m.Groups[1].Value); });
                    var ok = Encoding.ASCII.GetBytes("HTTP/1.1 204 No Content\r\nAccess-Control-Allow-Origin: *\r\nAccess-Control-Allow-Private-Network: true\r\nAccess-Control-Allow-Methods: GET\r\nConnection: close\r\n\r\n");
                    s.Write(ok, 0, ok.Length);
                }
                catch { }
            }
        }) { IsBackground = true }.Start();
    }

    // Typing in Word moves its text cursor right on the same line: sparks there. Only the cursor's
    // position is read (GetGUIThreadInfo), never the keys.
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr pid);
    [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint tid, ref GUITHREADINFO g);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h, ref Point p);
    [StructLayout(LayoutKind.Sequential)] struct GUITHREADINFO { public int cbSize, flags; public IntPtr active, focus, capture, menuOwner, moveSize, hwndCaret; public RECT rcCaret; }
    [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] static extern uint ProcessOf(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h, uint flags);
    // the Teams app (new Teams: ms-teams.exe), top-level window only
    static bool IsTeams(IntPtr h)
    {
        if (h == IntPtr.Zero || GetAncestor(h, 2) != h) return false;
        ProcessOf(h, out var pid);
        try { return System.Diagnostics.Process.GetProcessById((int)pid).ProcessName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase); } catch { return false; }
    }
    Point lastCaret = Point.Empty;
    readonly Timer caret = new() { Interval = 40 };
    void Caret()
    {
        if (!sparks || word == IntPtr.Zero || GetForegroundWindow() != word) { lastCaret = Point.Empty; return; }
        var g = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
        if (!GetGUIThreadInfo(GetWindowThreadProcessId(word, IntPtr.Zero), ref g) || g.hwndCaret == IntPtr.Zero) return;
        var p = new Point(g.rcCaret.L, (g.rcCaret.T + g.rcCaret.B) / 2); ClientToScreen(g.hwndCaret, ref p);
        int dx = p.X - lastCaret.X, dy = Math.Abs(p.Y - lastCaret.Y);
        if (lastCaret != Point.Empty && dx > 0 && dx < 60 && dy < 4) spark.Burst(p, word);
        lastCaret = p;
    }

    void Follow()
    {
        if (n == 0) { caret.Tick += (_, _) => Caret(); caret.Start(); }
        if (n++ % 8 == 0) FromDocument();
        var fg = GetForegroundWindow(); var cls = new StringBuilder(32); GetClassName(fg, cls, 32);
        if ((cls.ToString() == "OpusApp" || IsTeams(fg)) && fg != word) Attach(fg);     // whichever Word or Teams window is in front
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
        word = w; var cn = new StringBuilder(32); GetClassName(w, cn, 32); isWord = cn.ToString() == "OpusApp";
        if (isWord) try { foreach (dynamic win in WordApp().Windows) win.DisplayVerticalScrollBar = false; } catch { }
        SetWindowLongPtr(Handle, GWLP_HWNDPARENT, w);   // owned by Word: stays just above it, moves with it in the stack
    }

    // Find the app's flat background (Word: the grey around the pages; Teams: the backdrop behind its
    // panes) in what the window draws, in 2 px blocks: a block counts only when every pixel matches,
    // so pages, panes, text and icons always stay on top. Null when there's too little of it.
    Region Canvas(int w, int h)
    {
        if (w < 400 || h < 300) return null;
        using var bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
        using (var g = Graphics.FromImage(bmp)) { var dc = g.GetHdc(); PrintWindow(word, dc, 2); g.ReleaseHdc(dc); }
        var d = bmp.LockBits(new Rectangle(0, 0, w, h), System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
        int st = d.Stride / 4; var px = new int[st * h]; Marshal.Copy(d.Scan0, px, 0, px.Length); bmp.UnlockBits(d);
        int c = px[(h / 2) * st + w / 40], cr = (c >> 16) & 255, cg = (c >> 8) & 255, cb = c & 255;
        // wide enough for Word's near-black edges, not its dark page (#262626)
        bool Same(int i) { int p = px[i]; return Math.Abs(((p >> 16) & 255) - cr) + Math.Abs(((p >> 8) & 255) - cg) + Math.Abs((p & 255) - cb) < 40; }
        // Word: only between the ribbon and the status bar (its ribbon is the same dark colour)
        int y0 = 0, y1 = h - 1;
        if (isWord) { int sx = w / 40, my = h / 2; if (!Same(my * st + sx)) return null; y0 = my; while (y0 > 0 && Same((y0 - 1) * st + sx)) y0--; y1 = my; while (y1 < h - 1 && Same((y1 + 1) * st + sx)) y1++; }
        const int B = 2; var path = new System.Drawing.Drawing2D.GraphicsPath(System.Drawing.Drawing2D.FillMode.Winding); long cells = 0;
        int grow = isWord ? 2 : 0;   // Word: cover its thin page border too
        for (int by = y0; by + B <= y1 + 1; by += B)
        {
            int run = -1;
            for (int bx = 0; bx <= w; bx += B)
            {
                bool ok = bx + B <= w;
                for (int y = by; ok && y < by + B; y++) for (int x = bx; ok && x < bx + B; x++) ok = Same(y * st + x);
                if (ok) { if (run < 0) run = bx; cells++; }
                else if (run >= 0) { path.AddRectangle(new Rectangle(run - grow, by - grow, bx - run + grow * 2, B + grow * 2)); run = -1; }
            }
        }
        return cells * B * B < (long)w * h / 25 ? null : new Region(path);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        try { foreach (dynamic win in WordApp().Windows) win.DisplayVerticalScrollBar = true; } catch { }   // Word back as it was
        tray.Visible = false; base.OnFormClosed(e);
    }
}

// Theme list and colours: the extension's themes.js, read live from the website (new themes arrive
// by themselves) with the copy inside the .exe as fallback.
static class Themes
{
    public static readonly System.Collections.Generic.Dictionary<string, (Color a, Color b, string name)> All = new();
    public const string Site = "https://min3scon.github.io/Wild-Themes/";
    public static void Load(string js)
    {
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(js,
            "\"(\\w+)\":\\s*\\{\\s*\"name\":\\s*\"([^\"]+)\"[^{}]*?\"accent\":\\s*\"(#[0-9a-fA-F]{6})\",\\s*\"accent2\":\\s*\"(#[0-9a-fA-F]{6})\""))
            All[m.Groups[1].Value] = (ColorTranslator.FromHtml(m.Groups[3].Value), ColorTranslator.FromHtml(m.Groups[4].Value), m.Groups[2].Value);
    }
}

// A small see-through, click-through window over the page at the text cursor: a burst in the
// theme's style and colours (snow drifts, petals fall, bubbles rise, embers float, code rains).
class Sparks : Form
{
    [DllImport("user32.dll")] static extern IntPtr SetWindowLongPtr(IntPtr h, int i, IntPtr v);
    readonly Timer t = new() { Interval = 15 };
    readonly Random rnd = new();
    readonly System.Collections.Generic.List<(PointF p, PointF v, float life, Color c, char g)> ps = new();
    const int S = 160;
    IntPtr owner; string style = ""; Color[] cols = { Color.Gold, Color.White };

    public Sparks()
    {
        FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual;
        BackColor = Color.Magenta; TransparencyKey = Color.Magenta; DoubleBuffered = true; Size = new Size(S, S);
        t.Tick += (_, _) => Step();
    }
    protected override CreateParams CreateParams { get { var p = base.CreateParams; p.ExStyle |= 0x80000 | 0x20 | 0x80 | 0x08000000; return p; } }
    protected override bool ShowWithoutActivation => true;

    public void Theme(string id)
    {
        style = id ?? "";
        cols = Themes.All.TryGetValue(style, out var t) ? new[] { t.a, t.b, Color.White } : new[] { Color.Gold, Color.White };
        if (style == "winter") cols = new[] { Color.White, Color.FromArgb(210, 235, 255) };
        if (style == "sakura") cols = new[] { Color.FromArgb(255, 183, 213), Color.FromArgb(255, 220, 235) };
        if (style == "matrix") cols = new[] { Color.FromArgb(80, 255, 120), Color.FromArgb(180, 255, 190) };
    }

    public void Burst(Point at, IntPtr win)
    {
        if (owner != win) { owner = win; SetWindowLongPtr(Handle, -8, win); }
        if (ps.Count == 0 || !Visible) Location = new Point(at.X - S / 2, at.Y - S / 2);
        float ox = at.X - Location.X, oy = at.Y - Location.Y;
        int n = style == "matrix" ? 3 : 6;
        for (int i = 0; i < n; i++)
        {
            double a = rnd.NextDouble() * Math.PI * 2, sp = 0.8 + rnd.NextDouble() * 2.0;
            var v = style switch
            {
                "winter" or "sakura" => new PointF((float)(rnd.NextDouble() - .5) * 1.4f, 0.4f + (float)rnd.NextDouble()),
                "ocean" or "embers" or "lava" or "fireflies" => new PointF((float)(rnd.NextDouble() - .5) * 1.2f, -0.8f - (float)rnd.NextDouble() * 1.4f),
                "matrix" => new PointF(0, 1.5f + (float)rnd.NextDouble() * 1.5f),
                _ => new PointF((float)(Math.Cos(a) * sp), (float)(Math.Sin(a) * sp - 1.0)),
            };
            ps.Add((new PointF(ox + (float)(rnd.NextDouble() - .5) * 6, oy), v, 1f, cols[rnd.Next(cols.Length)], (char)rnd.Next(0x30A0, 0x30FF)));
        }
        if (!Visible) Show();
        t.Start();
    }
    void Step()
    {
        float grav = style switch { "winter" or "sakura" => 0.01f, "ocean" or "embers" or "lava" or "fireflies" => -0.01f, "matrix" => 0f, _ => 0.1f };
        for (int i = ps.Count - 1; i >= 0; i--)
        {
            var (p, v, life, c, g) = ps[i];
            life -= grav == 0.1f ? 0.04f : 0.022f;
            var np = new PointF(p.X + v.X, p.Y + v.Y);
            if (life <= 0 || np.X < 0 || np.Y < 0 || np.X > S || np.Y > S) { ps.RemoveAt(i); continue; }
            ps[i] = (np, new PointF(v.X * 0.97f + (style == "sakura" ? (float)Math.Sin(life * 12) * 0.08f : 0), v.Y + grav), life, c, g);
        }
        if (ps.Count == 0) { t.Stop(); Hide(); }
        Invalidate();
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        var gr = e.Graphics;
        using var f = new Font("Consolas", 9, FontStyle.Bold);
        foreach (var (p, _, life, c, g) in ps)
        {
            float r = 1f + 2.4f * life;
            using var b = new SolidBrush(c);
            switch (style)
            {
                case "ocean": using (var pen = new Pen(c, 1)) gr.DrawEllipse(pen, p.X - r - 1, p.Y - r - 1, (r + 1) * 2, (r + 1) * 2); break;
                case "sakura": gr.FillEllipse(b, p.X - r * 1.4f, p.Y - r * 0.7f, r * 2.8f, r * 1.4f); break;
                case "matrix": gr.DrawString(g.ToString(), f, b, p.X, p.Y); break;
                case "party": gr.FillRectangle(b, p.X - r, p.Y - r / 2, r * 2, r); break;
                default: gr.FillEllipse(b, p.X - r, p.Y - r, r * 2, r * 2); break;
            }
        }
    }
}

// Updates without downloading anything by hand: the theme list comes from the website on every
// start (scenes themselves load live in scene.html), and a newer .exe from the GitHub release is
// swapped in (a running .exe can be renamed, not overwritten) and started.
static class Updates
{
    static readonly System.Net.Http.HttpClient http = new() { Timeout = TimeSpan.FromMinutes(5) };
    static Updates() { http.DefaultRequestHeaders.UserAgent.ParseAdd("WildThemes-desktop"); }
    public static async void Themes(Action done)
    {
        try { global::Themes.Load(await http.GetStringAsync(global::Themes.Site + "extension/lib/themes.js")); done(); } catch { /* offline: the built-in list */ }
    }
    public static async void App()
    {
        var exe = Environment.ProcessPath;
        try { File.Delete(exe + ".old"); } catch { }
        try
        {
            var json = await http.GetStringAsync("https://api.github.com/repos/Min3scon/Wild-Themes/releases/latest");
            var tag = System.Text.RegularExpressions.Regex.Match(json, "\"tag_name\":\\s*\"desktop-v([0-9.]+)\"").Groups[1].Value;
            var url = System.Text.RegularExpressions.Regex.Match(json, "\"browser_download_url\":\\s*\"([^\"]+WildThemes\\.exe)\"").Groups[1].Value;
            if (tag == "" || url == "" || !Version.TryParse(tag, out var latest) || latest <= System.Reflection.Assembly.GetEntryAssembly().GetName().Version) return;
            var tmp = exe + ".new";
            await File.WriteAllBytesAsync(tmp, await http.GetByteArrayAsync(url));
            File.Move(exe, exe + ".old", true);
            File.Move(tmp, exe);
            System.Diagnostics.Process.Start(exe);
            Application.Exit();
        }
        catch { /* offline or no newer version: carry on */ }
    }
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
