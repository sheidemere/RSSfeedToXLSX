using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeOpenXml;
using Serilog;
using ExcelLib2;
using RSSLib;

class Program
{
    static async Task Main(string[] args)
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.Console()
            .WriteTo.File("logs/logfile.txt", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        string url = "https://www.goha.ru/rss/news";

        string rssContent = null;

        var path = AppDomain.CurrentDomain.BaseDirectory;

        Log.Information("Запуск...");

        using (var httpClient = new HttpClient())
        {
            ILoadRSS rssLoader = new LoadRSS(httpClient);
            rssContent = await rssLoader.LoaderAsync(url);
        }

        IParseRSS parser = new ParseRSS();
        var rssValues = parser.Parse(rssContent);

        ICreateExcel excelCreator = new Create();
        string filePath = excelCreator.CreateExcel(path);

        IAddData excelAdder = new Add();
        excelAdder.AddData(filePath, rssValues);

        IFormatData formatCreator = new Format();
        formatCreator.FormatData(filePath);

        Log.CloseAndFlush();
    }
}
