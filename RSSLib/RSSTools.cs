using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Serilog;

namespace RSSLib
{
    public interface ILoadRSS
    {
        Task<string> LoaderAsync(string url);
    }

    public interface IParseRSS
    {
        Dictionary<int, RSSItem> Parse(string rssContent);
    }

    public class LoadRSS : ILoadRSS
    {
        private readonly HttpClient _client;

        public LoadRSS(HttpClient client)
        {
            _client = client ?? throw new ArgumentNullException(nameof(client));
        }

        public async Task<string> LoaderAsync(string url)
        {
            try
            {
                Log.Information("Загрузка RSS...");
                return await _client.GetStringAsync(url);
            }
            catch (Exception ex)
            {
                Log.Error($"Ошибка при загрузке RSS: {ex?.Message ?? "Неизвестная ошибка"}");
                return null;
            }
        }
    }

    public class ParseRSS : IParseRSS
    {
        public Dictionary<int, RSSItem> Parse(string rssContent)
        {

            var RSSValues = new Dictionary<int, RSSItem>();

            try
            {
                if (!string.IsNullOrWhiteSpace(rssContent))
                {
                    var rssFeed = XDocument.Parse(rssContent);
                    var items = rssFeed.Descendants("item");

                    int id = 0;

                    foreach (var item in items)
                    {
                        id++;
                        var rssItem = new RSSItem
                        {
                            Title = item.Element("title")?.Value,
                            PublishDate = DateTime.Parse(item.Element("pubDate")?.Value),
                            Link = item.Element("link")?.Value,
                            Summary = item.Element("description")?.Value,
                            Authors = item.Elements("{http://purl.org/dc/elements/1.1/}creator")
                                          .Select(author => author.Value)
                                          .ToList()
                        };

                        Log.Information($"Статья {id} успешно загружена");

                        RSSValues[id] = rssItem;
                    }
                    Log.Information("RSS успешно загружен");
                }
                else
                {
                    Log.Error("Контент пуст, проверьте URL или повторите попытку");
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Ошибка при анализе RSS {ex.Message}");
            }

            return RSSValues;
        }
    }
}
