using System.Runtime.InteropServices;

namespace Dm.NET.Helpers
{
    /// <summary>
    /// Win32Api，原生方法比Dm還快
    /// </summary>
    public static class WindowHelper
    {
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        public static extern int FindWindow(string lpClassName, string lpWindowName);

        [DllImport("User32.dll", EntryPoint = "FindWindowEx")]
        public static extern int FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpClassName, string lpWindowName);

        // 定義所需的Win32 API函數
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public const int SW_RESTORE = 9;
        public const int SW_MAXIMIZE = 3;
        public const int SW_MINIMIZE = 6;

        // 設定主控台視窗的大小和位置
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        public static void SetPositionAndSizeMode(IntPtr hWnd, int X, int Y, int nWidth, int nHeight)
        {
            // 設定視窗的位置和大小
            // 例如：位置 (100, 100)，大小 (800, 600)
            MoveWindow(hWnd, X, Y, nWidth, nHeight, true);
        }
    }
}