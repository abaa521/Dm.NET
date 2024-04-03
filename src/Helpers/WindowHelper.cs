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

        public const int SW_RESTORE = 9;  // 激活並顯示窗口。如果窗口最小化或最大化，系統會將其恢復到原始大小和位置。
    }
}