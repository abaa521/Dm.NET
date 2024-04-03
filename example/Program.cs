using Dm.NET;

namespace Example
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            DmService? dm = null;
            try
            {
                //初始化dm時，輸入預設視窗大小，之後找圖就不用設定範圍
                dm = new DmService(1920, 1080);
            }
            catch (Exception)
            {
                //通常是被防毒刪除或沒有註冊
                Console.WriteLine("dm初始化失敗，代表沒有註冊大漠");
                Console.ReadLine();
            }

            if (dm == null)
            {
                Console.WriteLine("未知原因");
                Console.ReadLine();
                return;
            }

            #region 後台控制

            //找窗口句炳
            var hwnd = dm.FindWindow("視窗類名", "視窗名稱");
            if (hwnd == 0)
            {
                Console.WriteLine("沒有找到視窗");
                Console.ReadLine();
                return;
            }
            //輸入句柄並綁定
            dm.BindWindow(hwnd);

            #endregion 後台控制

            //設定圖片資料夾路徑、字典名稱
            //dm.SetPath("圖片資料夾路徑"); // 不設定預設找Resource資料夾
            //dm.SetDict("字典名稱"); // 不設定預設找dm_soft.txt

            //注意，所有圖片副檔名都是bmp，故找圖不需要再寫副檔名

            //一般找圖
            if (!dm.FindPicB("圖片1"))
            {
                Console.WriteLine("沒找到 圖片1");
                Console.ReadLine();
                return;
            }
            //找到了
            Console.WriteLine("找到 圖片1");

            //滑鼠移動至圖片、點擊、休息2秒
            dm.MCS();

            //滑鼠移動至圖片、點擊、休息5秒
            //dm.MCS(5);

            // 滑鼠移動至100,100、點擊、休息2秒
            //dm.MCS(100, 100);

            // 滑鼠移動至100,100、點擊、休息5秒
            //dm.MCS(100, 100, 5);

            // 隔一秒找一次圖片，預設時間找10秒
            if (dm.NotFindPicR("圖片1"))
            {
                //時間內沒找到圖片
                Console.WriteLine($"沒找到 圖片1");
                Console.ReadLine();
                return;
            }
            //找到了往下執行
            Console.WriteLine($"找到圖片1，執行下一步");

            //滑鼠移動至圖片、點擊、休息2秒
            dm.MCS();
        }
    }
}