using System;
using System.IO;
using System.Security.AccessControl;
using Serilog;
using OfficeOpenXml;
using System.ComponentModel;
using RSSLib;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ExcelLib2
{
    public interface ICreateExcel
    {
        string CreateExcel(string directoryPath);
    }

    public interface IAddData
    {
        void AddData(string filePath, Dictionary<int, RSSItem> rssValues);
    }

    public interface IFormatData
    {
        void FormatData(string filePath);
    }

    public class Create : ICreateExcel
    {
        public string CreateExcel(string directoryPath)
        {
            try
            {
                var fileName = "test.xlsx";
                var filePath = Path.Combine(directoryPath, fileName);

                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // обязательное требование пакета для использования некоторых функций

                using (var package = new ExcelPackage())
                {
                    package.Workbook.Worksheets.Add("RSS-лента");
                    Log.Information("Файл .xlsx успешно создан по пути {FilePath}", filePath);

                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo); // Сохраняем файл
                }

                return filePath;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при создании Excel файла");
                throw;
            }
        }
    }

    public class Add : IAddData
    {

        public void AddData(string filePath, Dictionary<int, RSSItem> rssValues)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                Log.Error("Путь к файлу не может быть пустым или null.");
                throw new ArgumentException("Путь к файлу не может быть пустым или null.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                Log.Error("Файл {FilePath} не найден.", filePath);
                throw new FileNotFoundException($"Файл {filePath} не найден.", filePath);
            }

            try
            {
                if (rssValues != null && rssValues.Any())
                {
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        worksheet.Cells[1, 1].Value = "Заголовок";
                        worksheet.Cells[1, 2].Value = "Ссылка";
                        worksheet.Cells[1, 3].Value = "Дата публикации";
                        worksheet.Cells[1, 4].Value = "Автор";

                        var row = 2;

                        foreach (var item in rssValues)
                        {
                            worksheet.Cells[row, 1].Value = item.Value.Title;
                            worksheet.Cells[row, 2].Value = item.Value.Link;
                            worksheet.Cells[row, 3].Value = item.Value.PublishDate.ToString("g");
                            worksheet.Cells[row, 4].Value = item.Value.FirstAuthor;
                            row++;
                        }

                        Log.Information("Файл .xlsx успешно заполнен данными по пути {FilePath}", filePath);
                        package.Save();
                    }
                }
                else
                {
                    Log.Warning("Нет данных для добавления в файл {FilePath}.", filePath);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при добавлении данных в файл {FilePath}", filePath);
                throw;
            }
        }
    }

    public class Format : IFormatData
    {
        public void FormatData(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                Log.Error("Путь к файлу не может быть пустым или null.");
                throw new ArgumentException("Путь к файлу не может быть пустым или null.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                Log.Error("Файл {FilePath} не найден.", filePath);
                throw new FileNotFoundException($"Файл {filePath} не найден.", filePath);
            }

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    using (var range = worksheet.Cells[1, 1, 1, 4])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    }

                    worksheet.Column(1).Width = 40;
                    worksheet.Column(2).Width = 60;
                    worksheet.Column(3).Width = 20;
                    worksheet.Column(4).Width = 10;

                    Log.Information("К файлу .xlsx применено форматирование по пути {FilePath}", filePath);

                    package.Save();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка при форматировании файла {FilePath}", filePath);
                throw;
            }
        }
    }
}
