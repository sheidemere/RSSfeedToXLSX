using RSSLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RSS2Excel.Tests
{
    public class ParseTests
    {
        private readonly ParseRSS _parser;

        public ParseTests()
        {
            _parser = new ParseRSS();
        }

        [Fact]
        public void Parse_ValidRssContent_ReturnsCorrectDictionary()
        {
            // Arrange
            string rssContent = @"
                <rss>
                    <channel>
                        <item>
                            <title>Title 1</title>
                            <pubDate>2024-10-01</pubDate>
                            <link>http://example.com/1</link>
                            <description>Summary 1</description>
                            <dc:creator xmlns:dc='http://purl.org/dc/elements/1.1/'>Author 1</dc:creator>
                        </item>
                        <item>
                            <title>Title 2</title>
                            <pubDate>2024-10-02</pubDate>
                            <link>http://example.com/2</link>
                            <description>Summary 2</description>
                            <dc:creator xmlns:dc='http://purl.org/dc/elements/1.1/'>Author 2</dc:creator>
                        </item>
                    </channel>
                </rss>";

            // Act
            var result = _parser.Parse(rssContent);

            // Assert
            Assert.Equal(2, result.Count);
            Assert.Contains(result, kvp => kvp.Value.Title == "Title 1");
            Assert.Contains(result, kvp => kvp.Value.Title == "Title 2");
        }

        [Fact]
        public void Parse_EmptyRssContent_ReturnsEmptyDictionary()
        {
            // Arrange
            string rssContent = string.Empty;

            // Act
            var result = _parser.Parse(rssContent);

            // Assert
            Assert.Empty(result);
        }
    }
}
