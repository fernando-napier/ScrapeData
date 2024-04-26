using CsvHelper;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Globalization;

namespace ScrapeData
{
    public interface IScrapeWorker 
    {
        void Run();
    }

    public class ScrapeWorker : IScrapeWorker
    {
        private readonly Configuration _config;
        private Random _random = new Random();

        public ScrapeWorker(Configuration config) 
        {
            _config = config;
        }

        public void Run()
        {
            var path = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName;
            var driver = new ChromeDriver(path + @"\Drivers\");

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            driver.Navigate().GoToUrl(_config.WebsiteUrl);

            // MANUAL INTERVENTION, click "I'm not a robot"
            var list = new List<InsuranceInfo>();

            for (char c1 = 'A'; c1 <= 'Z'; c1++)
            {
                for (char c2 = 'A'; c2 <= 'Z'; c2++)
                {
                    var term = string.Format(_config.Term, _config.ScrapeValue, c1, c2);
                    list.AddRange(ScrapeData(driver, term));
                    Console.WriteLine($"Done with {term} rolling total {list.Count}");
                }
            }

            using (var writer = new StreamWriter($"{_config.Directory}{DateTime.Now.ToString("yyyy-MM-dd")}-{_config.ScrapeValue}-scrape.csv"))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteHeader<InsuranceInfo>();
                csv.NextRecord();
                csv.WriteRecords(list);
                writer.Flush();
            }

            // write email
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Scrape Worker", _config.EmailAddress));
            message.To.Add(new MailboxAddress("", _config.EmailAddress));
            message.Subject = $"Scrape Done for {_config.ScrapeValue}";
            message.Body = new TextPart("plain") { Text = "Check out the results dawg" };

            using (var client = new SmtpClient())
            {
                client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                client.Authenticate(_config.EmailAddress, _config.EmailPassword);
                client.Send(message);
                client.Disconnect(true);
            }

            Console.WriteLine("Process Complete");
        }

        private List<InsuranceInfo> ScrapeData(ChromeDriver driver, string term)
        {
            var list = new List<InsuranceInfo>();

            try
            {
                var name = driver.FindElement(By.Id("eQInsEmpName"));
                name.SendKeys(term);
                var searchButton = driver.FindElement(By.Id("btnSearch"));
                searchButton.Click();
                Thread.Sleep(_random.Next(1000, 2000));

                // grab table and go through table results
                ScrapeTable(driver, list, term);

                // reset query
                var resetButton = driver.FindElement(By.Id("btnReset"));
                resetButton.Click();
                Thread.Sleep(_random.Next(1000, 2000));
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            return list;
           
        }

        private void ScrapeTable(ChromeDriver driver, List<InsuranceInfo> list, string term, int pagesScraped = 0)
        {
            try
            {
                driver.FindElement(By.Id("dgrSearchResults"));
            }
            catch(NoSuchElementException e) 
            {
                Console.WriteLine("---------");
                Console.WriteLine($"No Search Results for {term}");
                return;
            }

            try
            {
                var table = driver.FindElement(By.Id("dgrSearchResults")).FindElement(By.TagName("tbody"));
                var rowCount = table.FindElements(By.TagName("tr")).Count;
                for (int i = 0; i < rowCount; i++)
                {

                    table = driver.FindElement(By.Id("dgrSearchResults")).FindElement(By.TagName("tbody"));
                    var row = table.FindElements(By.TagName("tr"))[i];
                    var tableDataList = row.FindElements(By.TagName("td"));

                    // this is the navigation element
                    if (i == rowCount - 1)
                    {
                        var navData = tableDataList.First().FindElements(By.TagName("a"));
                        if (navData.Any(x => x.Text.Contains("Next")) && pagesScraped < 3)
                        {
                            var nextElement = navData.First(x => x.Text.Contains("Next"));
                            nextElement.Click();
                            Thread.Sleep(_random.Next(1000, 2000));
                            Console.WriteLine("---------");
                            Console.WriteLine($"Another page for {term}, total in batch {list.Count}");
                            ScrapeTable(driver, list, term, pagesScraped++);
                        }
                        break;
                    }

                    var policyNumber = tableDataList[0].Text;
                    var companyName = tableDataList[1].Text;
                    var city = tableDataList[2].Text;
                    
                    // skip the first record due to it being the 
                    if (policyNumber == "Policy Number")
                    {
                        continue;
                    }

                    if (list.Any(x => x.PolicyNumber == policyNumber && x.CompanyName == x.CompanyName))
                    {
                        if (!list.Any(x => x.PolicyNumber == policyNumber && 
                        x.CompanyName == x.CompanyName && x.City == city))
                        {
                            var matchingPolicy = list.First(x => x.PolicyNumber == policyNumber && x.CompanyName == x.CompanyName);
                            list.Add(new InsuranceInfo
                            {
                                PolicyNumber = policyNumber,
                                CompanyName = companyName,
                                City = city,
                                InsurerType = "matched policy",
                                InsurerName = matchingPolicy.InsurerName,
                                EffectiveDate = matchingPolicy.EffectiveDate,
                                ExpirationDate = matchingPolicy.ExpirationDate
                            });
                        }
                    }

                    // click on the link 
                    var clickableLink = tableDataList[0].FindElement(By.TagName("a"));
                    clickableLink.Click();
                    Thread.Sleep(_random.Next(1000, 2000));

                    // grab insurance information
                    var effectiveDate = string.Empty;
                    var expirationDate = string.Empty;
                    var insurerType = string.Empty;
                    var insurerName = string.Empty;
                    var insureTable = driver.FindElement(By.XPath("//table[@class=\"style1\"]"));
                    var insureTbody = insureTable.FindElement(By.TagName("tbody"));
                    foreach (var insureRow in insureTbody.FindElements(By.TagName("tr")))
                    {
                        var insureDataList = insureRow.FindElements(By.TagName("td"));

                        if (insureDataList[0].Text.Contains("Effective Date"))
                        {
                            effectiveDate = insureDataList[1].Text;
                            continue;
                        }

                        if (insureDataList[0].Text.Contains("Expiration Date"))
                        {
                            expirationDate = insureDataList[1].Text;
                            continue;
                        }

                        if (insureDataList[0].Text.Contains("Insurer Type"))
                        {
                            insurerType = insureDataList[1].Text;
                            continue;
                        }

                        if (insureDataList[0].Text.Contains("Insurer Name"))
                        {
                            insurerName = insureDataList[1].Text;
                            continue;
                        }
                    }

                    list.Add(new InsuranceInfo
                    {
                        CompanyName = companyName,
                        PolicyNumber = policyNumber,
                        City = city,
                        InsurerName = insurerName,
                        InsurerType = insurerType,
                        EffectiveDate = effectiveDate,
                        ExpirationDate = expirationDate,
                    });

                    driver.Navigate().Back();
                    Thread.Sleep(_random.Next(1000, 2000));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
