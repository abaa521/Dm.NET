using Dm.NET.Helpers;
using System.Diagnostics;

namespace Dm.NET
{
    /// <summary>
    /// 大漠服務
    /// </summary>
    public class DmService
    {
        private readonly int _width;
        private readonly int _height;
        private readonly int _sleepMilliseconds;
        public readonly dmsoft dm;

        public DmService(int Width = 960, int Height = 540, int sleepMilliseconds = 2000)
        {
            dm = new dmsoft();
            _height = Height;
            _width = Width;
            _sleepMilliseconds = sleepMilliseconds;

            SetPath();
            SetDict();
            if (!Debugger.IsAttached)
            {
                // 關閉錯誤訊息
                dm.SetShowErrorMsg(0);
            }
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

        /// <summary>
        /// 綁定視窗
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public bool BindWindow(int hwnd)
        {
            return dm.BindWindow(hwnd, "gdi", "windows", "windows", 0) == 1;
        }

        /// <summary>
        /// 尋找視窗
        /// </summary>
        /// <param name="lpClassName"></param>
        /// <param name="lpWindowName"></param>
        /// <returns></returns>
        public int FindWindow(string lpClassName, string lpWindowName)
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
        public int FindWindowEx(nint hwndParent, nint hwndChildAfter, string lpClassName, string lpWindowName)
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
            var dictPath = Path.Combine(resourcesPath, dictName + ".txt");
            if (!File.Exists(dictPath))
            {
                throw new Exception("字典不存在");
            }
            return dm.SetDict(0, Path.Combine(resourcesPath, dictPath)) == 1;
        }

        #endregion 設定

        #region 其他方法

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
        public void Capture(string bmp, int limit = 100)
        {
            int index = 0;
            var currentFile = bmp + ".bmp";
            while (File.Exists(Path.Combine(resourcesPath, currentFile)))
            {
                index++;
                currentFile = $"{bmp}{index}.bmp";
            }
            if (index > limit)
            {
                Console.WriteLine($"圖片超過{limit}張");
                return;
            }
            dm.Capture(0, 0, _width, _height, currentFile);
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

        public bool NotFindPicR(string? bmps, int time = 10, double sim = 0.7, bool traversal = false)
        {
            return FindPicRInternal(0, 0, _width, _height, bmps, time, sim, traversal);
        }

        public bool NotFindPicR(int x1, int y1, int x2, int y2, string? bmps, int time = 10, double sim = 0.7, bool traversal = false)
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
                    return false;
                }
                tmptime++;
                if (tmptime > time)
                {
                    return true;
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
        /// 長按滑鼠
        /// </summary>
        /// <param name="intX"></param>
        /// <param name="intY"></param>
        /// <param name="sec"></param>
        public void MDSU(int intX, int intY, int sec = 2)
        {
            dm.MoveTo(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftDown();
            Thread.Sleep(sec * 1000);
            dm.LeftUp();
        }

        public void MCS(int intX, int intY)
        {
            dm.MoveTo(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void MCS(int intX, int intY, int sec = 2)
        {
            dm.MoveTo(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void MCS(int intX, int intY, double sec = 2.0)
        {
            dm.MoveTo(intX + GetRandomNumberMove(), intY + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep((int)(sec * 1000));
        }

        public void MCSEx(int intXEx, int intYEx, int sec = 2)
        {
            dm.MoveTo(X + intXEx + GetRandomNumberMove(), Y + intYEx + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void MCS()
        {
            dm.MoveTo(X + GetRandomNumberMove(), Y + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void MCS(int sec)
        {
            dm.MoveTo(X + GetRandomNumberMove(), Y + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void MCS(double sec)
        {
            dm.MoveTo(X + GetRandomNumberMove(), Y + GetRandomNumberMove());
            dm.LeftClick();
            Thread.Sleep((int)(sec * 1000));
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
        public void MCSAccurate(int intX, int intY, int sec = 2)
        {
            dm.MoveTo(intX, intY);
            dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        #endregion 滑鼠
    }
}