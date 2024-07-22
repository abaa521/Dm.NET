using Dm.NET.Helpers;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Dm.NET
{
    /// <summary>
    /// 大漠服務
    /// </summary>
    public class DmService : IDisposable
    {
        private int _width;
        private int _height;
        private int _sleepMilliseconds;
        private string resourcesPath = string.Empty;
        private double _ratio;

        public dmsoft Dm { get; } = new();

        public DmService()
        {
            if (!Debugger.IsAttached)
            {
                // 關閉錯誤訊息
                Dm.SetShowErrorMsg(0);
            }
        }

        public void Init()
        {
            SetPath();
            //SetDict();
            SetSize();
            SetSleep();
            SetRatio();
        }

        #region 其他變數

        private object intX = -1;
        private object intY = -1;
        private int X => (int)intX;
        private int Y => (int)intY;

        private Dictionary<string, List<string>> bmpWithBmpListDict = [];

        #endregion 其他變數

        #region 視窗方法

        public void MoveWindow(int hwnd0, int x, int y)
        {
            Dm.MoveWindow(hwnd0, x, y);
        }

        public void SetWindowOnTop(int hwnd0)
        {
            Dm.SetWindowState(hwnd0, 8);
        }

        private int hwnd;

        /// <summary>
        /// 綁定視窗
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public bool BindWindow(int hwnd, bool sleep = true)
        {
            this.hwnd = hwnd;
            var result = Dm.BindWindow(hwnd, "gdi", "windows", "windows", 0) == 1;

            if (result && sleep)
                Thread.Sleep(1000);

            return result;
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
            return Dm.SetPath(resourcesPath) == 1;
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
                using (FileStream fs = File.Create(dictPath))
                {
                    // 這裡可以寫一些初始化的內容到檔案中
                    byte[] info = new UTF8Encoding(true).GetBytes("1F0F7E00C00801805FF0$0$0.0.33$12\r\n");
                    fs.Write(info, 0, info.Length);
                }
            }
            return Dm.SetDict(0, dictPath) == 1;
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

        public void SetRatio(double ratio = 1)
        {
            _ratio = ratio;
        }

        #endregion 設定

        #region 其他方法

        public void GetClientSize(int hwnd1, out object width, out object height)
        {
            Dm.GetClientSize(hwnd1, out width, out height);
        }

        public void SendString(string str)
        {
            Dm.SendString(hwnd, str);
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
            return Dm.IsDisplayDead(0, 0, _width, _height, sec) == 1;
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

            Dm.Capture(0, 0, _width, _height, filePath);
            return filePath;
        }

        /// <summary>
        /// 截圖，如果大於limit則不截圖，避免佔用太多空間
        /// </summary>
        /// <param name="bmp"></param>
        /// <param name="limit"></param>
        public string? Capture(int x1, int y1, int x2, int y2, string bmp)
        {
            var currentFile = $"{bmp}.bmp";
            var filePath = Path.Combine(resourcesPath, currentFile);

            Dm.Capture(x1, y1, x2, y2, filePath);
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
            return Dm.GetColor(x, y);
        }

        #endregion 其他方法

        #region 文字

        public bool FindStrB(string str, string colors, double sim = 0.7)
        {
            return Dm.FindStr(0, 0, _width, _height, str, colors, sim, out intX, out intY) >= 0;
        }

        public bool FindStrB(int x1, int y1, int x2, int y2, string str, string colors, double sim = 0.7)
        {
            return Dm.FindStr(x1, y1, x2, y2, str, colors, sim, out intX, out intY) >= 0;
        }

        #endregion 文字

        #region 圖片

        private int FindPicOrigin(int x1, int y1, int x2, int y2, string? bmpStr, double sim)
        {
            x1 = Math.Max(0, x1 - 1);
            y1 = Math.Max(0, y1 - 1);

            x2 = Math.Min(x2 + 1, _width);
            y2 = Math.Min(y2 + 1, _height);

            return Dm.FindPic(x1, y1, x2, y2, bmpStr, "000000", sim, 0, out intX, out intY);
        }

        private string ProcessBmpQuery(string? bmpQuery)
        {
            ArgumentException.ThrowIfNullOrEmpty(bmpQuery);

            if (bmpQuery.Contains(".bmp"))
            {
                bmpQuery = bmpQuery.Replace(".bmp", "");
                Console.WriteLine(bmpQuery + "包含bmp");
            }

            string bmpStr;
            var Querys = bmpQuery.Split("|");
            List<string> bmpResult = [];

            foreach (var query in Querys)
            {
                if (query.EndsWith('*'))
                {
                    var bmp = query.TrimEnd('*');
                    if (!bmpWithBmpListDict.TryGetValue(bmp, out List<string>? bmpList))
                    {
                        bmpList = Traversal(bmp);
                        bmpWithBmpListDict[bmp] = bmpList;
                    }
                    bmpResult.AddRange(bmpList);
                }
                else
                {
                    var bmp = query;
                    bmpResult.Add(bmp);
                }
            }

            for (int i = 0; i < bmpResult.Count; i++)
            {
                bmpResult[i] += ".bmp";
            }

            bmpStr = string.Join('|', bmpResult);

            return bmpStr;
        }

        private List<string> Traversal(string FilenameWithoutExtension)
        {
            List<string> fileNamesWithoutExtension = new List<string>();
            int index = 0;
            string currentFile;

            do
            {
                currentFile = index == 0 ? $"{FilenameWithoutExtension}.bmp" : $"{FilenameWithoutExtension}{index}.bmp";
                if (File.Exists(Path.Combine(resourcesPath, currentFile)))
                {
                    fileNamesWithoutExtension.Add(index == 0 ? FilenameWithoutExtension : $"{FilenameWithoutExtension}{index}");
                    index++;
                }
            } while (File.Exists(Path.Combine(resourcesPath, currentFile)));

            return fileNamesWithoutExtension;
        }

        #region picB

        public bool FindPicB(string? bmpQuery, double sim = 0.7)
        {
            return FindPicBInternal(0, 0, _width, _height, bmpQuery, sim);
        }

        public bool FindPicB(int x1, int y1, int x2, int y2, string? bmpQuery, double sim = 0.7)
        {
            return FindPicBInternal(x1, y1, x2, y2, bmpQuery, sim);
        }

        private bool FindPicBInternal(int x1, int y1, int x2, int y2, string? bmpQuery, double sim)
        {
            var bmpStr = ProcessBmpQuery(bmpQuery);
            return FindPicOrigin(x1, y1, x2, y2, bmpStr, sim) >= 0;
        }

        #endregion picB

        #region Pic

        public int FindPic(int x1, int y1, int x2, int y2, string? bmpQuery, double sim = 0.7)
        {
            return FindPicInternal(x1, y1, x2, y2, bmpQuery, sim);
        }

        public int FindPic(string? bmpQuery, double sim = 0.7)
        {
            return FindPicInternal(0, 0, _width, _height, bmpQuery, sim);
        }

        private int FindPicInternal(int x1, int y1, int x2, int y2, string? bmpQuery, double sim)
        {
            var bmpStr = ProcessBmpQuery(bmpQuery);
            return FindPicOrigin(x1, y1, x2, y2, bmpStr, sim);
        }

        #endregion Pic

        #region PicR

        public bool FindPicR(string? bmps, int times = 10, double sim = 0.7)
        {
            return FindPicRInternal(0, 0, _width, _height, bmps, times, sim);
        }

        public bool FindPicR(int x1, int y1, int x2, int y2, string? bmps, int times = 10, double sim = 0.7)
        {
            return FindPicRInternal(x1, y1, x2, y2, bmps, times, sim);
        }

        private bool FindPicRInternal(int x1, int y1, int x2, int y2, string? bmpQuery, int times, double sim)
        {
            var bmpStr = ProcessBmpQuery(bmpQuery);

            var currentTimes = 0;
            while (true)
            {
                currentTimes++;
                if (currentTimes > times)
                    return false;

                if (FindPicOrigin(x1, y1, x2, y2, bmpStr, sim) >= 0)
                    return true;

                Thread.Sleep(1000);
            }
        }

        #endregion PicR

        #endregion 圖片

        public bool FindR(bool result, int times = 10)
        {
            var currentTimes = 0;
            while (true)
            {
                currentTimes++;
                if (currentTimes > times)
                    return false;

                if (result)
                    return true;

                Thread.Sleep(1000);
            }
        }

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

            MoveToInternal(x1, y1, false);   // 初始位置
            Dm.LeftDown();       // 按下左鍵
            Thread.Sleep(50);

            for (var i = 1; i <= steps; i++)
            {
                // 每次增加一定的步長
                int nextX = (int)(x1 + stepX * i);
                int nextY = (int)(y1 + stepY * i);
                MoveToInternal(nextX, nextY, false);
                Thread.Sleep(50);  // 暫停一下來模擬滑動
            }

            Dm.LeftUp();         // 釋放左鍵
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
            MoveToInternal(intX, intY);
            Dm.LeftDown();
            Thread.Sleep(sec * 1000);
            Dm.LeftUp();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void Mcs(int intX, int intY)
        {
            MoveToInternal(intX, intY);
            Dm.LeftClick();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void Mcs(int intX, int intY, int sec = 2)
        {
            MoveToInternal(intX, intY);
            Dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void Mcs(int intX, int intY, double sec = 2.0)
        {
            MoveToInternal(intX, intY);
            Dm.LeftClick();
            Thread.Sleep((int)(sec * 1000));
        }

        public void McsEx(int intXEx, int intYEx, int sec = 2)
        {
            MoveToInternal(X + intXEx, Y + intYEx);
            Dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void Mcs()
        {
            MoveToInternal(X, Y);
            Dm.LeftClick();
            Thread.Sleep(_sleepMilliseconds);
        }

        public void Mcs(int sec)
        {
            MoveToInternal(X, Y);
            Dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        public void Mcs(double sec)
        {
            MoveToInternal(X, Y);
            Dm.LeftClick();
            Thread.Sleep((int)(sec * 1000));
        }

        public void MoveToInternal(int x, int y, bool random = true)
        {
            if (random)
            {
                Dm.MoveTo((int)(x * _ratio) + RandomHelper.GetRandomNumberMove(), (int)(y * _ratio + RandomHelper.GetRandomNumberMove()));
            }
            else
            {
                Dm.MoveTo((int)(x * _ratio), (int)(y * _ratio));
            }
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
            Dm.LeftClick();
            Thread.Sleep(sec * 1000);
        }

        private bool disposed = false; // To detect redundant calls

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // 釋放受控資源

                Dm.UnBindWindow();
            }

            // 釋放非受控資源（如果有）

            disposed = true;
        }

        ~DmService()
        {
            Dispose(false);
        }

        #endregion 滑鼠
    }
}