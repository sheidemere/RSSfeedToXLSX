using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RSSLib
{
    public class RSSItem
    {
        public string Title { get; set; }
        public DateTime PublishDate { get; set; }
        public string Link { get; set; }
        public string Summary { get; set; }
        public List<string> Authors { get; set; } = new List<string>();
        public string FirstAuthor => Authors.FirstOrDefault() ?? "Автор не указан";
    }
}
