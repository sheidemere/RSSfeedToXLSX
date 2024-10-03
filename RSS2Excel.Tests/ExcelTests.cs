using ExcelLib2;
using OfficeOpenXml;
using RSSLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RSS2Excel.Tests
{
    public class ExcelTests
    {
        private readonly Add _add;
        private readonly Create _excel;

        public ExcelTests()
        {
            _add = new Add();
            _excel = new Create();
        }

        [Fact]
        public void AddData_EmptyFilePath_ThrowsArgumentException()
        {
            // Arrange
            string filePath = string.Empty;
            var rssValues = new Dictionary<int, RSSItem>();

            // Act & Assert
            var exception = Assert.Throws<ArgumentException>(() => _add.AddData(filePath, rssValues));
            Assert.Equal("Путь к файлу не может быть пустым или null. (Parameter 'filePath')", exception.Message);
        }

        [Fact]
        public void AddData_FileNotFound_ThrowsFileNotFoundException()
        {
            // Arrange
            string filePath = "nonexistent.xlsx";
            var rssValues = new Dictionary<int, RSSItem>();

            // Act & Assert
            var exception = Assert.Throws<FileNotFoundException>(() => _add.AddData(filePath, rssValues));
            Assert.Equal($"Файл {filePath} не найден.", exception.Message);
        }

        [Fact]
        public void AddData_ValidFilePath_AddsDataToFile()
        {
            // Arrange
            string directoryPath = Path.GetTempPath();
            string filePath = _excel.CreateExcel(directoryPath);

            var rssValues = new Dictionary<int, RSSItem>
    {
        { 1, new RSSItem { Title = "Title 1", Link = "http://example.com/1", PublishDate = DateTime.Now, Authors = new List<string> { "Author 1" } } },
        { 2, new RSSItem { Title = "Title 2", Link = "http://example.com/2", PublishDate = DateTime.Now.AddHours(1), Authors = new List<string> { "Author 2" } } }
    };

            Assert.True(File.Exists(filePath), "Файл не был создан.");

            // Act
            _add.AddData(filePath, rssValues);

            // Assert
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                Assert.NotNull(worksheet);
                Assert.Equal("Title 1", worksheet.Cells[2, 1].Value);
                Assert.Equal("http://example.com/1", worksheet.Cells[2, 2].Value);

                // Получаем значение из ячейки и пытаемся его преобразовать в DateTime
                string dateValue = worksheet.Cells[2, 3].Value?.ToString();
                Assert.NotNull(dateValue);

                // Пробуем преобразовать строку в DateTime
                Assert.True(DateTime.TryParse(dateValue, out DateTime parsedDate), "Не удалось преобразовать строку в DateTime.");
                Assert.Equal(rssValues[1].PublishDate.ToString("g").Split(' ')[0], parsedDate.ToString("g").Split(' ')[0]);

                Assert.Equal("Author 1", worksheet.Cells[2, 4].Value);
                Assert.Equal("Title 2", worksheet.Cells[3, 1].Value);
            }

            File.Delete(filePath);
        }
    }
}
