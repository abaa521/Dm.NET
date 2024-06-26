# Dm.NET簡介

大漠插件是一個用於Windows平台的鍵盤、滑鼠以及圖色控制的插件。

本類別庫的目的是為了便於.Net使用者使用這個插件，透過將其核心功能封裝成更加友好的接口。

## 功能特點

- **易於使用**：對大漠插件的主要功能進行了封裝，使其更易使用。

## 大漠安裝指南

1. 解壓縮專案中的 `dm.zip` 檔案。
2. 右鍵以系統管理員身分執行解壓縮後的 `.bat` 檔案。

# 初始化
```csharp
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
```
# 後臺控制
```csharp
//找窗口句炳
var hwnd = dm.FindWindow("視窗類名", "視窗名稱");
//輸入句柄並綁定
dm.BindWindow(hwnd);
```
# 找圖
- 注意，所有圖片副檔名都是bmp，故找圖不需要再寫副檔名
## 一般找圖
```csharp
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
```
## 持續找圖
```csharp
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
```
## 參數說明
### 找圖
- 找圖說明 FindPic可限定範圍 例如 FindPicB(100, 100, 200, 200, "圖片名稱")
- NotFindPicR 的 time為尋找時間，單位為秒，預設10秒
  
### 找多圖
- 找圖方法的參數 traversal 設為true時候，會一次尋找 "圖片名稱"、"圖片名稱1"、"圖片名稱2"等等，持續往上加，直到沒有此張圖片
- 找圖要找多張圖片 以|分開 例如 "圖片A|圖片名E|圖片X"
- 因此「traversal」與「bmp的|」使用方法為互斥，只能使用一個
### 其他
- sim 設為0.9時候，會尋找相似度90%以上的圖片，同大漠

## 使用方法

若要在您的.NET專案中使用`DmService`，須將專案設定成x86
右鍵選擇「專案」>「屬性」>「平台目標」>「X86」

然後依照下列步驟進行：

**方法一：加入參考**

1. **將`DmService`專案加入至您的解決方案**：
   - 在Visual Studio中，右鍵點擊解決方案名稱，選擇「新增」->「現有專案…」。
   - 瀏覽至`DmService`專案檔案（通常是`.csproj`或`.vbproj`檔案），然後點擊「開啟」。

2. **為您的專案新增對`DmService`的參考**：
   - 在解決方案資源管理器中，右鍵點擊您想使用`DmService`的專案名稱。
   - 選擇「新增」->「參考…」，在彈出的對話框中選擇「專案」分頁。
   - 於列表中勾選`DmService`專案，然後點擊「確定」。

3. **在您的專案中使用`Dm.NET`**：
   - 在您的程式碼檔案中，新增對`Dm.NET`命名空間的引用：
     ```csharp
     using Dm.NET; // 請替換為實際的命名空間。
     ```
   - 您可以透過繼承`DmService`來擴展功能：
     ```csharp
     public class MyDmService : DmService {
       // 新增或覆寫功能。
     }
     ```
   - 或者，您可以建立一個`DmService`的實例：
     ```csharp
     var dmService = new DmService();
     // 使用dmService實例。
     ```

請依照您實際的命名空間和類別名稱調整上述代碼範例。

## 如何解決找不到Dm插件的問題

如果在使用過程中你的程式無法找到Dm插件，請按照以下步驟操作：

1. 在你的解決方案中，右鍵選擇「專案」>「加入」>「COM參考」。
2. 在出現的對話框中搜索「Dm」。
3. 找到對應的項目後打勾，然後點擊「確定」。

這樣應該能夠解決找不到Dm插件的問題。
