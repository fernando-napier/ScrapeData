using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScrapeData
{
    public class Configuration
    {
        public string WebsiteUrl { get; set; }
        public string Term { get; set; }
        public string ScrapeValue {  get; set; }
        public string Directory { get; set; }
        public string EmailPassword { get; set; }
        public string EmailAddress { get; set; }
    }
}
