using ExcelLibrary.SpreadSheet;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OzonParser;
using OpenQA.Selenium.Chrome;
using OzonParser.Models;
using OzonParser.UI;

var keys = File.ReadAllLines("Keys.txt");

var links = new List<SortedLinks>();

ChromeDriverService service = ChromeDriverService.CreateDefaultService(Directory.GetCurrentDirectory());
service.EnableVerboseLogging = false;
service.SuppressInitialDiagnosticInformation = true;
service.HideCommandPromptWindow = true;  
service.EnableAppendLog = false;

var options = new ChromeOptions();
options.PageLoadStrategy = PageLoadStrategy.Normal;
options.AddArgument("--disable-in-process-stack-traces");
options.AddArgument("--window-position=-32000,-32000");
options.AddArgument("ignore-certificate-errors");
options.AddArgument("--disable-crash-reporter");
options.AddArgument("--disable-dev-shm-usage");
options.AddArgument("--output=/dev/null");
options.AddArgument("--disable-logging");
options.AddArgument("--log-level=3");
options.AddArgument("--no-sandbox");

var doc = new HtmlDocument();

foreach (var key in keys)
{
    var browser = new ChromeDriver(service, options);
    browser.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
    
    browser.Navigate().GoToUrl($"https://www.ozon.ru/search/?text={key.Replace(" ", "+")}&from_global=true");
    
    //В зависимости есть ли фильтры на странице или нет, XPath меняется
    bool withFilters = false;
    
    //Проверка что все нужные элементы на странице загрузились
    //На странице всего 36 элементов
    for (int g = 0; g < 36; g++)
    {
        var index = g + 1;

        try
        {
            _ = browser.FindElement(By.XPath(
                    $"/html/body/div[1]/div/div[1]/div[2]/div[2]/div[2]/div[{(withFilters ? "3" : "5")}]/div[1]/div/div/div[{index}]/div[1]/a/span/span"),
                5);
        }
        catch
        {
            withFilters = true;
        }
    }

    doc.LoadHtml(browser.PageSource);

    links.Add(new SortedLinks()
    {
        Key = key,
        Links = doc.DocumentNode
            .SelectSingleNode(
                $"/html/body/div[1]/div/div[1]/div[2]/div[2]/div[2]/div[{(withFilters ? "3" : "5")}]/div[1]/div/div")
            .ChildNodes
            .Where(x => x.Name == "div")
            .Where(x => x.GetAttributeValue("style", string.Empty) == String.Empty)
            .Select(x => x.ChildNodes.FirstOrDefault(x => x.Name == "a")!.GetAttributeValue("href", string.Empty))
            .ToList()
    });

    browser.Close();
    browser.Quit();

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Выбрано {links.First(x => x.Key == key).Links.Count} товаров по ключу: {key}");
    Console.ForegroundColor = ConsoleColor.White;
}

var allProductsCount = links.SelectMany(x => x.Links).Count();

Console.WriteLine();
Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine($"Количество всех выбранных товаров составляет: {allProductsCount}");

var ozonProducts = keys.Select(x => new SortedOzonProducts()
{
    Key = x,
    OzonProudcts = new List<OzonProudct>()
}).ToList();

Console.WriteLine();
Console.Write("Парсинг товаров по отдельности... ");
Console.ForegroundColor = ConsoleColor.White;
int progressBarIndex = 0;
using (var progress = new ProgressBar())
{
    foreach (var sortedLinks in links)
    {
        foreach (var link in sortedLinks.Links)
        {
            var browser = new ChromeDriver(service, options);
            browser.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            browser.Navigate().GoToUrl("https://www.ozon.ru" + link);

            var title = browser.FindElement(By.ClassName("o6r"), 5)
                .Text;

            string brand = browser.TryFindWhileBy(new[]
            {
                By.ClassName("mr4")
            }, 5)?.Text ?? String.Empty;

            var id = browser.FindElement(By.ClassName("n4q"), 5).Text
                .Replace(":", String.Empty).Trim().Split(" ")[1];
            
            string feedbacks = browser.TryFindWhileBy(new[]
            {
                By.XPath("/html/body/div[1]/div/div[1]/div[3]/div[2]/div/div/div[2]/div/div[1]/div[1]/div/div/div[2]/a/div/div"),
                By.XPath("/html/body/div[1]/div/div[1]/div[3]/div[3]/div[3]/div/div[8]/div[1]/div/div/div[2]/a/div/div"),
                By.XPath("/html/body/div[1]/div/div[1]/div[3]/div[3]/div[3]/div/div[7]/div[1]/div/div/div[2]/a/div/div")
            }, 5)!.Text.Split(" ")[0];

            string price = browser.TryFindWhileBy(new[]
                {
                    By.ClassName("o3p"), By.ClassName("po2")
                }, 5)!
                .Text.Split(" ")[0];


            ozonProducts.FirstOrDefault(x => x.Key == sortedLinks.Key)!.OzonProudcts
                .Add(new OzonProudct()
                {
                    Title = title,
                    Brand = brand,
                    Feedbacks = feedbacks,
                    Id = id,
                    Price = price
                });

            browser.Close();
            browser.Quit();
            progressBarIndex += 1;
            progress.Report((double)progressBarIndex / allProductsCount);
        }
    }
}

Console.WriteLine();
Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine("Создание ozonproducts.xlsx");
Workbook workbook = new Workbook();

foreach (var sortedOzonProducts in ozonProducts)
{
    Worksheet worksheet = new Worksheet(sortedOzonProducts.Key);
    worksheet.Cells[0, 0] = new Cell("Title");
    worksheet.Cells[0, 1] = new Cell("Brand");
    worksheet.Cells[0, 2] = new Cell("Id");
    worksheet.Cells[0, 3] = new Cell("Feedbacks");
    worksheet.Cells[0, 4] = new Cell("Price");

    for (int i = 0; i < sortedOzonProducts.OzonProudcts.Count; i++)
    {
        var rowIndex = i + 1;
        worksheet.Cells[rowIndex, 0] = new Cell(sortedOzonProducts.OzonProudcts[i].Title);
        worksheet.Cells[rowIndex, 1] = new Cell(sortedOzonProducts.OzonProudcts[i].Brand);
        worksheet.Cells[rowIndex, 2] = new Cell(sortedOzonProducts.OzonProudcts[i].Id);
        worksheet.Cells[rowIndex, 3] = new Cell(sortedOzonProducts.OzonProudcts[i].Feedbacks);
        worksheet.Cells[rowIndex, 4] = new Cell(sortedOzonProducts.OzonProudcts[i].Price);
    }

    workbook.Worksheets.Add(worksheet);
}

workbook.Save("ozonproducts.xlsx");

Console.WriteLine();
Console.WriteLine("Работа OzonParser'a закончена");
Console.ForegroundColor = ConsoleColor.White;
Console.WriteLine("Нажмите на любую кнопку в консоли для ее закрытия.");
Console.ReadKey();