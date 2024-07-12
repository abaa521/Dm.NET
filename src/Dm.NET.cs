using Dm.NET.Helpers;
using System.Diagnostics;

namespace Dm.NET
{
    /// <summary>
    /// 大漠服務
    /// </summary>
    public class DmService
    {
        private int _width;
        private int _height;
        private int _sleepMilliseconds;
        public readonly dmsoft dm = new();
        private double _ratio = 1;

        public DmService()
        {
            if (!Debugger.IsAttached)

            {
                // 關閉錯誤訊息
                dm.SetShowErrorMsg(0);
            }
        }

        public void Init()
        {
            SetPath();
            SetDict();
            SetSize();
            SetSleep();
        }

        #region 其他變數

        public int re = -1;
        public object intX = -1;
        public object intY = -1;
        public int X => (int)intX;
        public int Y => (int)intY;
        private string resourcesPath = string.Empty;

        #endregion 其他變數

        #region 視窗方法

        public void MoveWindow(int hwnd0, int x, int y)
        {
            dm.MoveWindow(hwnd0, x, y);
        }

        public void SetWindowOnTop(int hwnd0)
        {
            dm.SetWindowState(hwnd0, 8);
        }

        private int hwnd;

        /// <summary>
        /// 綁定視窗
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public bool BindWindow(int hwnd)
        {
            this.hwnd = hwnd;
            return dm.BindWindow(hwnd, "gdi", "windows", "windows", 0) == 1;
        }

        /// <summary>
        /// 尋找視窗
        /// </summary>
        /// <param name="lpClassName"></param>
        /// <param name="lpWindowName"></param>
        /// <returns></returns>
        public static int FindWindow(string lpClassName, string lpWindowName)
        {
            return WindowHelper.FindWindow(lpClassName, lpWindowName);
        }

        /// <summary>
        /// 尋找視窗EX
        /// </summary>
        /// <param name="hwndParent"></param>
        /// <param name="hwndChildAfter"></param>
        /// <param name="lpClassName"></param>
        /// <param name="lpWindowName"></param>
        /// <returns></returns>
        public static int FindWindowEx(nint hwndParent, nint hwndChildAfter, string lpClassName, string lpWindowName)
        {
            return WindowHelper.FindWindowEx(hwndParent, hwndChildAfter, lpClassName, lpWindowName);
        }

        /// <summary>
        /// 設定資源路徑
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>

        #endregion 視窗方法

        #region 設定

        public bool SetPath(string? path = null)
        {
            if (string.IsNullOrEmpty(path))
            {
                var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                path = Path.Combine(baseDirectory, "Resources");
            }

            // 確保資料夾存在
            Directory.CreateDirectory(path);

            resourcesPath = path;
            return dm.SetPath(resourcesPath) == 1;
        }

        /// <summary>
        /// 設定字典，因為很常需要換字典，所以只傳字典名稱，不包含副檔名
        /// </summary>
        /// <param name="dict"></param>
        public bool SetDict(string? dictName = null)
        {
            if (string.IsNullOrEmpty(dictName))
            {
                dictName = "dm_soft";
            }
            var dictNameWithTxt = dictName + ".txt";
            var dictPath = Path.Combine(resourcesPath, dictNameWithTxt);
            if (!File.Exists(dictPath))
            {
                File.Create(dictPath);
            }
            return dm.SetDict(0, dictPath) == 1;
        }

        public void SetSize(int Width = 960, int Height = 540)
        {
            _height = Height;
            _width = Width;
        }

        public void SetSleep(int sleepMilliseconds = 2000)
        {
            _sleepMilliseconds = sleepMilliseconds;
        }

        public void SetRatio(double ratio)
        {
            _ratio = ratio;
        }

        #endregion 設定

        #region 其他方法

        public void GetClientSize(int hwnd1, out object width, out object height)
        {
            dm.GetClientSize(hwnd1, out width, out height);
        }

        public void SendString(string str)
        {
            dm.SendString(hwnd, str);
            Thread.Sleep(2000);
        }

        //public bool GetClientSize(int hwnd, out object width, out object height)
        //{
        //    return dm.GetClientSize(hwnd,) == 1;
        //}

        /// <summary>
        /// 是否卡死
        /// </summary>
        /// <param name="sec"></param>
        /// <returns></returns>
        public bool IsDisplayDead(int sec = 5)
        {
            return dm.IsDisplayDead(0, 0, _width, _height, sec) == 1;
        }

        /// <summary>
        /// 截圖，如果大於limit則不截圖，避免佔用太多空間
        /// </summary>
        /// <param name="bmp"></param>
        /// <param name="limit"></param>
        public string? Capture(string bmp, int limit = 100)
        {
            int index = 0;
            var currentFile = $"{bmp}.bmp";
            var filePath = Path.Combine(resourcesPath, currentFile);

            while (File.Exists(filePath))
            {
                index++;
                if (index > limit)
                {
                    Console.WriteLine($"圖片超過{limit}張");
                    return null;
                }
                currentFile = $"{bmp}{index}.bmp";
                filePath = Path.Combine(resourcesPath, currentFile);
            }

            dm.Capture(0, 0, _width, _height, filePath);
            return filePath;
        }

        /// <summary>
        /// 取色
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public string GetColor(int x, int y)
        {
            return dm.GetColor(x, y);
        }

        #endregion 其他方法

        #region 文字

        public bool FindStrB(string str, string colors, double sim = 0.7)
        {
            return dm.FindStr(0, 0, _width, _height, str, colors, sim, out intX, out intY) >= 0;
        }

        public bool FindStrB(int x1, int y1, int x2, int y2, string str, string colors, double sim = 0.7)
        {
            return dm.FindStr(x1, y1, x2, y2, str, colors, sim, out intX, out intY) >= 0;
        }

        #endregion 文字

        #region 圖片

        public bool FindPicB(string bmps, double sim = 0.7, bool traversal = false)
        {
            return FindPicBInternal(0, 0, _width, _height, bmps, sim, traversal);
        }

        public bool FindPicB(int x1, int y1, int x2, int y2, string bmps, double sim = 0.7, bool traversal = false)
        {
            return FindPicBInternal(x1, y1, x2, y2, bmps, sim, traversal);
        }

        private bool FindPicBInternal(int x1, int y1, int x2, int y2, string? bmps, double sim, bool traversal)
        {
            var bmp = ProcessBmpString(bmps, traversal);
            var re = dm.FindPic(x1, y1, x2, y2, bmp, "000000", sim, 0, out intX, out intY) >= 0;
            //if (re && click)
            //{
            //    MCS();
            //}
            return re;
            //return FindPicRInternal(x10, y1, x2, y2, bmps, click, time, sim, traversal);
        }

        public int FindPic(int x1, int y1, int x2, int y2, string bmps, double sim = 0.7, bool traversal = false)
        {
            var bmp = ProcessBmpString(bmps, traversal);
            return dm.FindPic(x1, y1, x2, y2, bmp, "000000", sim, 0, out intX, out intY);
        }

        public int FindPic(string bmps, double sim = 0.7, bool traversal = false)
        {
            var bmp = ProcessBmpString(bmps, traversal);

            return dm.FindPic(0, 0, _width, _height, bmp, "000000", sim, 0, out intX, out intY);
        }

        public bool FindPicR(string? bmps, int time = 10, double sim = 0.7, bool traversal = false)
        {
            return FindPicRInternal(0, 0, _width, _height, bmps, time, sim, traversal);
        }

        public bool FindPicR(int x1, int y1, int x2, int y2, string? bmps, int time = 10, double sim = 0.7, bool traversal = false)
        {
            return FindPicRInternal(x1, y1, x2, y2, bmps, time, sim, traversal);
        }

        private bool FindPicRInternal(int x1, int y1, int x2, int y2, string? bmps, int time, double sim, bool traversal)
        {
            var bmp = ProcessBmpString(bmps, traversal);

            var tmptime = 0;
            while (true)
            {
                if (dm.FindPic(x1, y1, x2, y2, bmp, "000000", sim, 0, out intX, out intY) >= 0)
                {
                    return true;
                }
                tmptime++;
                if (tmptime > time)
                {
                    return false;
                }

                Thread.Sleep(1000);
            }
        }

        private string ProcessBmpString(string? bmps, bool traversal)
        {
            if (string.IsNullOrEmpty(bmps))
            {
                throw new Exception("bmps IsNullOrEmpty");
            }

            if (traversal && bmps.Contains('|'))
            {
                throw new Exception(bmps + " 圖片判斷錯誤: 多張遍歷");
            }

            if (bmps.Contains(".bmp"))
            {
                bmps = bmps.Replace(".bmp", "");
                Console.WriteLine(bmps + "包含bmp");
            }

            string? bmp;
            if (traversal)
            {
                bmp = Traversal(bmps);
            }
            else if (bmps.Contains('|'))
            {
                var tmp = bmps.Split("|");
                for (int i = 0; i < tmp.Length; i++)
                {
                    tmp[i] += ".bmp";
                }

                bmp = string.Join('|', tmp);
            }
            else
            {
                bmp = bmps + ".bmp";
            }
            return bmp;
        }

        private string Traversal(string baseFilename)
        {
            List<string> filenames = [];
            //string basePath = AppDomain.CurrentDomain.BaseDirectory;

            int index = 0;
            string currentFile = baseFilename + ".bmp";
            while (File.Exists(Path.Combine(resourcesPath, currentFile)))
            {
                filenames.Add(currentFile);
                index++;
                currentFile = $"{baseFilename}{index}.bmp";
            }

            return string.Join("|", filenames);
        }

        #endregion 圖片

        #region 滑鼠

        /// <summary>
        /// 滑動滑鼠
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        public void Mdsmsrus(int x1, int y1, int x2, int y2)
        {
            int steps = 40;
            float stepX = (float)(x2 - x1) / steps;
            float stepY = (float)(y2 - y1) / steps;

            MoveToInternal(x1, y1);   // 初始位置
            dm.LeftDown();       // 按下左鍵
            Thread.Sleep(50);

            for (var i = 1; i <= steps; i++)
            {
                // 每次增加一定的步長
                int nextX = (int)(x1 + stepX * i);
                int nextY = (int)(y1 + stepY * i);
                MoveToInternal(nextX, nextY);
                Thread.Sleep(50);  // 暫停一下來模擬滑動
            }

            dm.LeftUp();         // 釋放左鍵
            Thread.Sleep(_sleepMilliseconds);  // 完成後等待
        }

        /// <summary>
        /// 長按滑鼠
        /// </summary>
        /// <param name="intX"></param>
        /// <param name="intY"></param>
        /// <param name="sec"></param>
        public void Mdsus(int intX, int intY, int sec = 2)
        {
            MoveToInternal(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftDown();
            Thread.Sleep(sec * 1000);
            dm.LeftUp();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void Mcs(int intX, int intY)
        {
            MoveToInternal(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void Mcs(int intX, int intY, int sec = 2)
        {
            MoveToInternal(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void Mcs(int intX, int intY, double sec = 2.0)
        {
            MoveToInternal(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep((int)(sec * 1000));
        }

        public void McsEx(int intXEx, int intYEx, int sec = 2)
        {
            MoveToInternal(X + intXEx + GetRandomNumberMove(), Y + intYEx + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void Mcs()
        {
            MoveToInternal(X + GetRandomNumberMove(), Y + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void Mcs(int sec)
        {
            MoveToInternal(X + GetRandomNumberMove(), Y + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void Mcs(double sec)
        {
            MoveToInternal(X + GetRandomNumberMove(), Y + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep((int)(sec * 1000));
        }

        public void MoveToInternal(int x, int y)
        {
            dm.MoveTo((int)(x * _ratio), (int)(y * _ratio));
        }

        private readonly Random random = new();

        private int GetRandomNumberMove()
        {
            return random.Next(0, 6); // 返回0到5的隨機整數
        }

        /// <summary>
        /// 準確移動到指定位置
        /// </summary>
        /// <param name="intX"></param>
        /// <param name="intY"></param>
        /// <param name="sec"></param>
        public void McsAccurate(int intX, int intY, int sec = 2)
        {
            MoveToInternal(intX, intY);
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        #endregion 滑鼠
    }
}