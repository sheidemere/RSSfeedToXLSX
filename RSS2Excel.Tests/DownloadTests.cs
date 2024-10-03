using System.Net;
using RSSLib;
using Moq;
using Moq.Protected;

namespace RSS2Excel.Tests
{
    public class DownloadTests
    {
        public class LoadRSSTests
        {
            [Fact]
            public async Task LoaderAsync_ReturnsRssContent_WhenUrlIsValid()
            {
                // Arrange
                var expectedContent = "<rss>Valid RSS content</rss>";
                var mockHttpMessageHandler = new Mock<HttpMessageHandler>();

                mockHttpMessageHandler.Protected()
                    .Setup<Task<HttpResponseMessage>>(
                        "SendAsync",
                        ItExpr.IsAny<HttpRequestMessage>(),
                        ItExpr.IsAny<CancellationToken>()
                    )
                    .ReturnsAsync(new HttpResponseMessage
                    {
                        StatusCode = HttpStatusCode.OK,
                        Content = new StringContent(expectedContent)
                    });

                var httpClient = new HttpClient(mockHttpMessageHandler.Object);
                var loadRss = new LoadRSS(httpClient);

                // Act
                var result = await loadRss.LoaderAsync("http://valid-url.com/rss");

                // Assert
                Assert.Equal(expectedContent, result);
            }

            [Fact]
            public async Task LoaderAsync_ReturnsNull_WhenExceptionOccurs()
            {
                // Arrange
                var mockHttpMessageHandler = new Mock<HttpMessageHandler>();

                mockHttpMessageHandler.Protected()
                    .Setup<Task<HttpResponseMessage>>(
                        "SendAsync",
                        ItExpr.IsAny<HttpRequestMessage>(),
                        ItExpr.IsAny<CancellationToken>()
                    )
                    .ThrowsAsync(new HttpRequestException("Network error"));

                var httpClient = new HttpClient(mockHttpMessageHandler.Object);
                var loadRss = new LoadRSS(httpClient);

                // Act
                var result = await loadRss.LoaderAsync("http://invalid-url.com/rss");

                // Assert
                Assert.Null(result);
            }
        }
    }
}