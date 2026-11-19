// Wild Themes for desktop Word and Teams. Their ribbons, rails and the grey space around Word's
// page can't be themed from a file or an add-in, so this small app lays the theme's animated scene
// (the extension's own CSS, rendered by WebView2) over each window's plain background only: one
// layer per Word or Teams window, owned by that window, cut to the background it draws. Word and
// Teams themselves are never changed; quit from the tray icon.
using System;
using System.Collections.Generic;
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
        // never die on a stray error (a window closing mid-frame, Word busy)
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, _) => { };
        Application.Run(new Host());
    }
}

static class Win
{
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
    [DllImport("user32.dll")] public static extern IntPtr SetWindowLongPtr(IntPtr h, int i, IntPtr v);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int w, int hh, uint f);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint tid, ref GUITHREADINFO g);
    [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref Point p);
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)] static extern int CLSIDFromProgID(string p, out Guid g);
    [DllImport("oleaut32.dll")] static extern int GetActiveObject(ref Guid g, IntPtr r, [MarshalAs(UnmanagedType.IUnknown)] out object o);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] public struct GUITHREADINFO { public int cbSize, flags; public IntPtr active, focus, capture, menuOwner, moveSize, hwndCaret; public RECT rcCaret; }

    public static string Class(IntPtr h) { var s = new StringBuilder(32); GetClassName(h, s, 32); return s.ToString(); }
    public static bool IsWord(IntPtr h) => Class(h) == "OpusApp";
    // the Teams app (new Teams: ms-teams.exe), visible top-level windows with a title
    public static bool IsTeams(IntPtr h)
    {
        if (h == IntPtr.Zero || GetAncestor(h, 2) != h) return false;
        GetWindowThreadProcessId(h, out var pid);
        try { var p = System.Diagnostics.Process.GetProcessById((int)pid); return p.ProcessName.Equals("ms-teams", StringComparison.OrdinalIgnoreCase) && p.MainWindowHandle == h; } catch { return false; }
    }
    public static dynamic WordApp() { CLSIDFromProgID("Word.Application", out var g); return GetActiveObject(ref g, IntPtr.Zero, out var o) == 0 ? o : throw new InvalidOperationException(); }
    public static void WordScrollBars(bool on) { try { foreach (dynamic w in WordApp().Windows) w.DisplayVerticalScrollBar = on; } catch { } }
}

// What every layer shows.
static class State
{
    public static string Scene;
    // Teams has little open background (title bar and rail only), so its layer is opt-in
    public static bool Teams = Settings.Get("teams", "0") == "1", Word = Settings.Get("word", "1") == "1";
    public static bool Dark = Settings.Get("dark", "1") == "1", Animated = Settings.Get("animated", "1") == "1", Sparks = Settings.Get("sparks", "1") == "1";
    public static readonly Lazy<System.Threading.Tasks.Task<Microsoft.Web.WebView2.Core.CoreWebView2Environment>> Env = new(() =>
        Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WildThemes", "webview")));
}

// The tray icon, the theme (from the open file, the website or the extension), sparks, updates,
// and one Layer per Word or Teams window.
class Host : ApplicationContext
{
    readonly NotifyIcon tray = new() { Visible = true, Text = "Wild Themes" };
    readonly Dictionary<IntPtr, Layer> layers = new();
    readonly Timer find = new() { Interval = 500 }, caret = new() { Interval = 40 };
    readonly Sparks spark = new();
    readonly string dir = Path.Combine(AppContext.BaseDirectory, "scenes");
    readonly Control ui = new();
    string fileScene; int n;

    public Host()
    {
        ui.CreateControl();
        tray.Icon = SystemIcons.Application;
        var menu = new ContextMenuStrip();
        var themes = new ToolStripMenuItem("Theme");
        menu.Items.Add(themes);
        void List() { themes.DropDownItems.Clear(); foreach (var (id, t) in Themes.All.OrderBy(x => x.Value.name)) themes.DropDownItems.Add(t.name, null, (_, _) => Use(id)); }
        try { Themes.Load(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "themes.js"))); } catch { }
        List();
        Updates.Themes(() => ui.BeginInvoke(() => { List(); spark.Theme(State.Scene); }));
        Updates.App();
        menu.Items.Add(new ToolStripSeparator());
        Toggle(menu, "Dark mode", State.Dark, v => { State.Dark = v; Settings.Set("dark", v ? "1" : "0"); Refresh(); });
        Toggle(menu, "Animated", State.Animated, v => { State.Animated = v; Settings.Set("animated", v ? "1" : "0"); Refresh(); });
        Toggle(menu, "Theme in Word", State.Word, v => { State.Word = v; Settings.Set("word", v ? "1" : "0"); if (!v) Win.WordScrollBars(true); Find(); });
        Toggle(menu, "Theme in Teams (rough)", State.Teams, v => { State.Teams = v; Settings.Set("teams", v ? "1" : "0"); Find(); });
        Toggle(menu, "Sparks when typing", State.Sparks, v => { State.Sparks = v; Settings.Set("sparks", v ? "1" : "0"); });
        // runs from sign-in by default, so the scene is there whenever Word or Teams opens
        var run = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
        if (Settings.Get("autostart", "1") == "1") run.SetValue("WildThemes", $"\"{Environment.ProcessPath}\"");
        Toggle(menu, "Start with Windows", Settings.Get("autostart", "1") == "1", v => { Settings.Set("autostart", v ? "1" : "0"); if (v) run.SetValue("WildThemes", $"\"{Environment.ProcessPath}\""); else run.DeleteValue("WildThemes", false); });
        menu.Items.Add("Quit", null, (_, _) => Quit());
        tray.ContextMenuStrip = menu;
        Use(Settings.Get("scene", "synthwave"));
        Listen();
        find.Tick += (_, _) => Find(); find.Start();
        caret.Tick += (_, _) => Caret(); caret.Start();
        Find();
    }
    static void Toggle(ContextMenuStrip m, string label, bool on, Action<bool> set)
    {
        var i = new ToolStripMenuItem(label) { Checked = on, CheckOnClick = true };
        i.CheckedChanged += (_, _) => set(i.Checked);
        m.Items.Add(i);
    }

    bool Known(string id) => id != null && (Themes.All.ContainsKey(id) || File.Exists(Path.Combine(dir, id + ".css")));
    void Use(string id, bool remember = true)
    {
        if (!Known(id)) return;
        if (remember) Settings.Set("scene", id);
        State.Scene = id; spark.Theme(id); Refresh();
    }
    void Refresh() { foreach (var l in layers.Values) l.Play(); }

    // every Word and Teams window gets its own layer; closed ones lose theirs
    void Find()
    {
        if (n++ % 4 == 0) FromDocument();
        var seen = new HashSet<IntPtr>();
        Win.EnumWindows((h, _) => { if (Win.IsWindowVisible(h) && ((State.Word && Win.IsWord(h)) || (State.Teams && Win.IsTeams(h)))) seen.Add(h); return true; }, IntPtr.Zero);
        foreach (var h in seen) if (!layers.ContainsKey(h)) layers[h] = new Layer(h);
        foreach (var h in layers.Keys.Where(k => !seen.Contains(k) || !Win.IsWindow(k)).ToList()) { layers[h].Close(); layers.Remove(h); }
    }

    // The open document's own theme (custom property WildTheme), so the space around it matches the page.
    void FromDocument()
    {
        try
        {
            string id = Win.WordApp().ActiveDocument.CustomDocumentProperties.Item("WildTheme").Value;
            if (id != fileScene && Known(id)) { fileScene = id; Use(id, false); }
            return;
        }
        catch { /* no document or no theme in it */ }
        if (fileScene != null) { fileScene = null; Use(Settings.Get("scene", "synthwave"), false); }   // back to the chosen theme
    }

    // The website and the extension send the chosen theme (this PC only).
    void Listen()
    {
        var l = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 47615);
        new System.Threading.Thread(() =>
        {
            for (int i = 0; ; i++) { try { l.Start(); break; } catch { if (i > 20) return; System.Threading.Thread.Sleep(500); } }   // port frees when an old copy quits
            while (true)
            {
                try
                {
                    using var c = l.AcceptTcpClient(); using var s = c.GetStream();
                    var buf = new byte[2048]; var line = Encoding.ASCII.GetString(buf, 0, s.Read(buf, 0, buf.Length)).Split('\n')[0];
                    var m = System.Text.RegularExpressions.Regex.Match(line, @"GET /theme\?id=([a-z0-9-]*)");
                    if (m.Success) ui.BeginInvoke(() => { if (!Known(m.Groups[1].Value)) return; Settings.Set("scene", m.Groups[1].Value); if (fileScene == null) Use(m.Groups[1].Value); });
                    var ok = Encoding.ASCII.GetBytes("HTTP/1.1 204 No Content\r\nAccess-Control-Allow-Origin: *\r\nAccess-Control-Allow-Private-Network: true\r\nAccess-Control-Allow-Methods: GET\r\nConnection: close\r\n\r\n");
                    s.Write(ok, 0, ok.Length);
                }
                catch { }
            }
        }) { IsBackground = true }.Start();
    }

    // Typing moves the text cursor right on the same line: sparks there. Only the cursor's position
    // is read (GetGUIThreadInfo), never the keys.
    Point lastCaret = Point.Empty;
    void Caret()
    {
        var fg = Win.GetForegroundWindow();
        if (!State.Sparks || !layers.ContainsKey(fg)) { lastCaret = Point.Empty; return; }
        var g = new Win.GUITHREADINFO { cbSize = Marshal.SizeOf<Win.GUITHREADINFO>() };
        if (!Win.GetGUIThreadInfo(Win.GetWindowThreadProcessId(fg, out _), ref g) || g.hwndCaret == IntPtr.Zero) return;
        var p = new Point(g.rcCaret.L, (g.rcCaret.T + g.rcCaret.B) / 2); Win.ClientToScreen(g.hwndCaret, ref p);
        int dx = p.X - lastCaret.X, dy = Math.Abs(p.Y - lastCaret.Y);
        if (lastCaret != Point.Empty && dx > 0 && dx < 60 && dy < 4) spark.Burst(p, fg);
        lastCaret = p;
    }

    void Quit()
    {
        foreach (var l in layers.Values) l.Close();
        Win.WordScrollBars(true);   // Word back as it was
        tray.Visible = false;
        ExitThread();
    }
}

// The scene over one Word or Teams window: follows it every frame, re-cut to its background a few
// times a second (not while it's being dragged).
class Layer : Form
{
    readonly IntPtr target; readonly bool word;
    readonly Microsoft.Web.WebView2.WinForms.WebView2 web = new() { Dock = DockStyle.Fill, DefaultBackgroundColor = Color.Black };
    readonly Timer move = new() { Interval = 15 }, cut = new() { Interval = 300 };
    bool ready; Rectangle last; DateTime moved;

    public Layer(IntPtr t)
    {
        target = t; word = Win.IsWord(t);
        FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual;
        BackColor = Color.Black; Controls.Add(web);
        CreateHandle();
        Win.SetWindowLongPtr(Handle, -8, t);   // owned by its window: stays just above it in the stack
        if (word) Win.WordScrollBars(false);   // the scene fills that strip; the wheel over it scrolls Word
        Init();
        move.Tick += (_, _) => Follow(); move.Start();
        cut.Tick += (_, _) => Recut(); cut.Start();
    }
    // never takes focus, no taskbar button
    protected override CreateParams CreateParams { get { var p = base.CreateParams; p.ExStyle |= 0x80 | 0x08000000; return p; } }
    protected override bool ShowWithoutActivation => true;
    protected override void OnFormClosed(FormClosedEventArgs e) { move.Stop(); cut.Stop(); base.OnFormClosed(e); }   // its window has gone

    async void Init()
    {
        await web.EnsureCoreWebView2Async(await State.Env.Value);
        web.CoreWebView2.SetVirtualHostNameToFolderMapping("wildthemes.local", AppContext.BaseDirectory, Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);
        web.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        web.CoreWebView2.WebMessageReceived += (_, e) => { if (!word) return; try { int d = int.Parse(e.TryGetWebMessageAsString()); dynamic a = Win.WordApp(); if (d > 0) a.ActiveWindow.SmallScroll(d); else a.ActiveWindow.SmallScroll(0, -d); } catch { } };
        ready = true; Play();
    }
    public void Play()
    {
        if (ready && State.Scene != null) web.CoreWebView2.Navigate($"https://wildthemes.local/scene.html?id={State.Scene}&dark={(State.Dark ? 1 : 0)}&anim={(State.Animated ? 1 : 0)}");
    }

    void Follow()
    {
        if (IsDisposed) return;
        if (!Win.IsWindow(target) || Win.IsIconic(target) || !Win.IsWindowVisible(target) || !Win.GetWindowRect(target, out var r)) { if (Visible) Hide(); return; }
        var b = new Rectangle(r.L, r.T, r.R - r.L, r.B - r.T);
        if (b == last) return;
        if (b.Size == last.Size) Win.SetWindowPos(Handle, IntPtr.Zero, b.X, b.Y, 0, 0, 0x1 | 0x4 | 0x10);   // moved only: keep the cut, no resize
        else { Bounds = b; moved = DateTime.Now; }
        last = b;
    }
    void Recut()
    {
        if (IsDisposed || last.IsEmpty || (DateTime.Now - moved).TotalMilliseconds < 250) return;
        var region = Background.Of(target, word, last.Width, last.Height);
        if (region == null) { if (Visible) Hide(); return; }
        Region = region;
        if (!Visible && !Win.IsIconic(target)) Show();
    }
}

// Find the plain background a window draws (Word: the grey around the pages; Teams: behind its
// panes), in 2 px blocks: a block counts only when every pixel matches, so pages, panes, text and
// icons stay on top. Teams: thin gaps (between icons, around text) are dropped, only open areas stay.
static class Background
{
    public static Region Of(IntPtr win, bool word, int w, int h)
    {
        if (w < 400 || h < 300) return null;
        using var bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
        using (var g = Graphics.FromImage(bmp)) { var dc = g.GetHdc(); Win.PrintWindow(win, dc, 2); g.ReleaseHdc(dc); }
        var d = bmp.LockBits(new Rectangle(0, 0, w, h), System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
        int st = d.Stride / 4; var px = new int[st * h]; Marshal.Copy(d.Scan0, px, 0, px.Length); bmp.UnlockBits(d);
        int c = px[(h / 2) * st + w / 40], cr = (c >> 16) & 255, cg = (c >> 8) & 255, cb = c & 255;
        // wide enough for Word's near-black edges, not its dark page (#262626)
        bool Same(int i) { int p = px[i]; return Math.Abs(((p >> 16) & 255) - cr) + Math.Abs(((p >> 8) & 255) - cg) + Math.Abs((p & 255) - cb) < (word ? 40 : 14); }
        // Word: only between the ribbon and the status bar (its ribbon is the same dark colour)
        int y0 = 0, y1 = h - 1;
        if (word) { int sx = w / 40, my = h / 2; if (!Same(my * st + sx)) return null; y0 = my; while (y0 > 0 && Same((y0 - 1) * st + sx)) y0--; y1 = my; while (y1 < h - 1 && Same((y1 + 1) * st + sx)) y1++; }
        const int B = 2; int mw = w / B, mh = (y1 - y0 + 1) / B;
        var m = new bool[mh, mw];
        for (int by = 0; by < mh; by++)
            for (int bx = 0; bx < mw; bx++)
            {
                bool ok = true;
                for (int y = y0 + by * B; ok && y < y0 + by * B + B; y++) for (int x = bx * B; ok && x < bx * B + B; x++) ok = Same(y * st + x);
                m[by, bx] = ok;
            }
        if (!word) { m = Grow(Shrink(m, 6), 6); }   // drop gaps under 24 px
        var path = new System.Drawing.Drawing2D.GraphicsPath(System.Drawing.Drawing2D.FillMode.Winding); long cells = 0;
        int grow = word ? 2 : 0;   // Word: cover its thin page border too
        for (int by = 0; by < mh; by++)
        {
            int run = -1;
            for (int bx = 0; bx <= mw; bx++)
            {
                bool ok = bx < mw && m[by, bx];
                if (ok) { if (run < 0) run = bx; cells++; }
                else if (run >= 0) { path.AddRectangle(new Rectangle(run * B - grow, y0 + by * B - grow, (bx - run) * B + grow * 2, B + grow * 2)); run = -1; }
            }
        }
        return cells * B * B < (long)w * h / (word ? 25 : 60) ? null : new Region(path);
    }
    // square min/max filters, k blocks each way, done one direction at a time
    static bool[,] Shrink(bool[,] m, int k) => Pass(Pass(m, k, true, true), k, false, true);
    static bool[,] Grow(bool[,] m, int k) => Pass(Pass(m, k, true, false), k, false, false);
    static bool[,] Pass(bool[,] m, int k, bool rows, bool all)
    {
        int H = m.GetLength(0), W = m.GetLength(1), A = rows ? H : W, L = rows ? W : H;
        var o = new bool[H, W]; var sum = new int[L + 1];
        for (int a = 0; a < A; a++)
        {
            for (int i = 0; i < L; i++) sum[i + 1] = sum[i] + ((rows ? m[a, i] : m[i, a]) ? 1 : 0);
            for (int i = 0; i < L; i++)
            {
                int lo = Math.Max(0, i - k), hi = Math.Min(L, i + k + 1), cnt = sum[hi] - sum[lo];
                bool v = all ? cnt == hi - lo && i - k >= 0 && i + k < L : cnt > 0;
                if (rows) o[a, i] = v; else o[i, a] = v;
            }
        }
        return o;
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
