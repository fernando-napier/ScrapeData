using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScrapeData
{
    public class InsuranceInfo
    {
        public string CompanyName { get; set; }
        public string PolicyNumber { get; set; }
        public string City { get; set; }
        public string InsurerType { get; set; }
        public string InsurerName { get; set; }
        public string EffectiveDate { get; set; }
        public string ExpirationDate { get; set; }
    }
}
