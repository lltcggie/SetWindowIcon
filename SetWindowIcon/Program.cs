using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Timers;
using System.Windows.Forms;


namespace SetWindowIcon
{
    public class ProcessIconSetting
    {
        public string ProcessPath { get; set; } = string.Empty;
        public string IconPath { get; set; } = string.Empty;

        public ProcessIconSetting(string ProcessPath, string IconPath)
        {
            this.ProcessPath = ProcessPath.ToLower();
            this.IconPath = IconPath;
        }
    }

    public class MonitoringSetting
    {
        public int MonitoringIntervalMS { get; set; } = 0;
        public List<ProcessIconSetting> ProcessIconSettings { get; set; }

        public MonitoringSetting(int MonitoringIntervalMS, List<ProcessIconSetting> ProcessIconSettings)
        {
            this.MonitoringIntervalMS = MonitoringIntervalMS;
            this.ProcessIconSettings = ProcessIconSettings;
        }
    }

    internal static class Program
    {
        #region DllImport
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsDelegate lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        // ウィンドウを作成したプロセスIDを取得
        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hwnd, int message, int wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool QueryFullProcessImageNameW(IntPtr hProcess, int flags, StringBuilder text, ref int count);

        private const int WM_SETICON = 0x80;
        private const int ICON_SMALL = 0;
        private const int ICON_BIG = 1;

        private const int PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

        private delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lParam);
        #endregion

        private static MonitoringSetting mSetting;
        private static System.Timers.Timer mTimer;
        private static NotifyIcon mIcon;
        private static Dictionary<string, Tuple<Icon, Icon>> mProcessIconMap;

        // 何回もバッファ生成するのはなんなので保持しておく
        private static StringBuilder mStringBuffer = new StringBuilder(1024);

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            mSetting = ReadSetting();
            mProcessIconMap = CreateIcons();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            CreateNotifyIcon();
            SetTimer();

            OnInit();

            Application.Run();

            mIcon.Dispose();
            mTimer.Dispose();

            foreach (var p in mProcessIconMap)
            {
                p.Value.Item1.Dispose();
                p.Value.Item2.Dispose();
            }
        }

        static Dictionary<string, Tuple<Icon, Icon>> CreateIcons()
        {
            Dictionary<string, Tuple<Icon, Icon>> dic = new Dictionary<string, Tuple<Icon, Icon>>();

            foreach (var setting in mSetting.ProcessIconSettings)
            {
                Tuple<Icon, Icon> icons = new Tuple<Icon, Icon>(new Icon(setting.IconPath, 32, 32), new Icon(setting.IconPath, 16, 16));
                dic[setting.ProcessPath] = icons;
            }

            return dic;
        }

        static void OnInit()
        {
            new ToastContentBuilder()
                .AddText("プロセス監視開始", AdaptiveTextStyle.Default)
                .Show();
        }

        static MonitoringSetting ReadSetting()
        {
            MonitoringSetting ret;

            string settingPath = System.IO.Path.Combine(GetExeDir(), "Setting.json");
            using (FileStream fileStream = File.OpenRead(settingPath))
            {
                using (StreamReader reader = new StreamReader(fileStream, System.Text.Encoding.UTF8))
                {
                    ret = JsonSerializer.Deserialize<MonitoringSetting>(reader.ReadToEnd());
                }
            }

            return ret;
        }

        static string GetExeDir()
        {
            return System.IO.Path.GetDirectoryName(Application.ExecutablePath);
        }

        private static void SetTimer()
        {
            mTimer = new System.Timers.Timer(mSetting.MonitoringIntervalMS);
            mTimer.Elapsed += OnTimedEvent;
            mTimer.AutoReset = true;
            mTimer.Enabled = true;
        }

        private static void CreateNotifyIcon()
        {
            // 常駐アプリ（タスクトレイのアイコン）を作成
            mIcon = new NotifyIcon();
            mIcon.Icon = new Icon("Icon.ico");
            mIcon.ContextMenuStrip = ContextMenu();
            mIcon.Text = "プロセスアイコン設定";
            mIcon.Visible = true;
        }

        private static ContextMenuStrip ContextMenu()
        {
            // アイコンを右クリックしたときのメニューを返却
            var menu = new ContextMenuStrip();
            menu.Items.Add("終了", null, (s, e) =>
            {
                Application.Exit();
            });
            return menu;
        }

        private static Dictionary<Tuple<uint, IntPtr>, uint> mUpdatedHwndList = new Dictionary<Tuple<uint, IntPtr>, uint>();

        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            foreach (var key in new List<Tuple<uint, IntPtr>>(mUpdatedHwndList.Keys))
            {
                mUpdatedHwndList[key] = 0;
            }

            EnumWindows(EnumerateWindows, IntPtr.Zero);

            // 今回見つからなかった物はもう消えたということでキャッシュから消す
            foreach (var item in mUpdatedHwndList.Where(x => x.Value == 0).ToList())
            {
                mUpdatedHwndList.Remove(item.Key);
            }
        }

        // ウィンドウを列挙するコールバックメソッド
        static bool EnumerateWindows(IntPtr hWnd, IntPtr lParam)
        {
            if (hWnd == IntPtr.Zero)
            {
                return true;
            }

            if (!IsWindowVisible(hWnd))
            {
                return true;
            }

            // ウィンドウハンドルからプロセスIDを取得
            uint processId;
            if (GetWindowThreadProcessId(hWnd, out processId) == 0)
            {
                return true;
            }

            var hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, processId);
            if (hProc == IntPtr.Zero)
            {
                return true;
            }

            int count = mStringBuffer.Capacity;
            if (!QueryFullProcessImageNameW(hProc, 0, mStringBuffer, ref count))
            {
                CloseHandle(hProc);

                return true;
            }

            CloseHandle(hProc);

            var path = mStringBuffer.ToString().ToLower();

            try
            {
                Tuple<Icon, Icon> icons;
                if (mProcessIconMap.TryGetValue(path, out icons))
                {
                    UpdateIcon(processId, hWnd, icons);
                }
            }
            catch (Exception) { }

            return true;
        }

        static void UpdateIcon(uint processId, IntPtr hWnd, Tuple<Icon, Icon> icons)
        {
            Tuple<uint, IntPtr> key = new Tuple<uint, IntPtr>(processId, hWnd);

            uint count = 0;
            if (mUpdatedHwndList.TryGetValue(key, out count)) // すでにアイコン変更済み
            {
                mUpdatedHwndList[key] = count + 1;
            }
            else // 初めて見つけたhWnd
            {
                SendMessage(hWnd, WM_SETICON, ICON_BIG, icons.Item1.Handle);
                SendMessage(hWnd, WM_SETICON, ICON_SMALL, icons.Item2.Handle);
                mUpdatedHwndList[key] = 1;
            }
        }
    }
}
