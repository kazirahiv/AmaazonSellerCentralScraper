using AmazonScraper.Models;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace AmazonScraper
{
    class Program
    {

        #region Properties 
        private static string email;
        private static string password;
        private static string baseUrl = "https://sellercentral.amazon.com/";
        private static string otp;
        private static string captcha;
        private static string UnicorpURI;
        public static string Email
        {
            get { return email; }

            set
            {

                if (String.IsNullOrEmpty(value) || value.Length > 35 || !value.Contains("@"))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    throw new ArgumentException("Invalid Email length.");
                }
                Console.ResetColor();
                email = value;
            }
        }
        public static string Password
        {
            get { return password; }

            set
            {
                if (String.IsNullOrEmpty(value) || value.Length > 35)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    throw new ArgumentException("Invalid password length.");
                }
                Console.ResetColor();
                password = value;
            }
        }

        public static string OTP
        {
            get { return otp; }

            set
            {
                bool isDigit = Regex.IsMatch(value, @"^\d+$");

                if (String.IsNullOrEmpty(value) || value.Length != 6 || !isDigit)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    throw new ArgumentException("Invalid OTP Entered.");
                }
                Console.ResetColor();
                otp = value;
            }
        }
        public static bool otpDone;
        public static string Captcha
        {
            get { return captcha; }

            set
            {

                if (String.IsNullOrEmpty(value))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    throw new ArgumentException("Invalid Captcha Entered.");
                }
                Console.ResetColor();
                captcha = value;
            }
        }
        public static bool captchaRaised;
        public static bool captchaDone;


        #endregion

        static void Main(string[] args)
        {
            var options = new ChromeOptions();
            IConfiguration configuration = GetAppConfig();
            var section = configuration.GetSection("UserAuth");
            var emailFromConfig = section.GetValue<string>("Email");
            var passwordFromConfig = section.GetValue<string>("Password");

            IConfiguration scrapeInfo = GetScrapeInfo();

            string lastScrapedDateTimeString = string.Empty;
            DateTime lastScrapedDateTime = DateTime.MinValue;
            try
            {
                lastScrapedDateTime = scrapeInfo.GetValue<DateTime>("LastScrapedDatePickerTime");
                lastScrapedDateTimeString = lastScrapedDateTime.ToString("dd-MMMM-yyyy");
            }
            catch { }


            if (!string.IsNullOrEmpty(emailFromConfig) && !string.IsNullOrEmpty(passwordFromConfig))
            {
                Email = emailFromConfig;
                Password = passwordFromConfig;
            }
            UnicorpURI = configuration.GetValue<string>("UnicorpURI");

            using (var driver = new ChromeDriver(options))
            {
                driver.Navigate().GoToUrl(baseUrl);
                try
                {
                    var signInButton = driver.FindElementByCssSelector("#wp-content > div.as-body.desktop > div.border-color-squid-ink.flex-container.flex-align-items-stretch.flex-align-content-flex-start.flex-full-width.amsg-2018.fonts-loaded.border-color-squid-ink.design-Sell > div > div > div.background-color-aqua.border-color-mermaid.padding-left-xxlarge.padding-right-xxlarge.padding-top-xsmall.padding-bottom-xsmall.flex-container.flex-align-items-center.flex-align-content-flex-start.flex-full-width.amsg-2018.fonts-loaded.border-color-mermaid.design-Sell > div > div.border-color-squid-ink.flex-container.flex-align-items-center.flex-align-content-flex-start.amsg-2018.fonts-loaded.border-color-squid-ink.design-Sell > div:nth-child(1) > div.border-color-squid-ink.padding-right-xsmall.flex-container.flex-align-items-stretch.flex-align-content-flex-start.flex-full-width.amsg-2018.fonts-loaded.border-color-squid-ink.design-Sell > div > a > strong");
                    signInButton.Click();
                }
                catch (Exception e)
                {
                    driver.FindElementById("sign-in-button").Click();
                    Console.WriteLine(e.Message);
                }

                if (string.IsNullOrEmpty(Email) || string.IsNullOrEmpty(Password))
                {
                    #region Email Input and Password Input Validation Check 
                    bool emailEntered = false;
                    while (!emailEntered)
                    {
                        try
                        {
                            Console.WriteLine();
                            Console.WriteLine("Please enter your Email:");
                            Email = Console.ReadLine();
                            emailEntered = true;
                        }
                        catch (ArgumentException ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(ex.Message);
                            Console.ResetColor();
                        }
                    }
                    bool passwordEntered = false;
                    while (!passwordEntered)
                    {
                        try
                        {
                            Console.WriteLine();
                            Console.WriteLine("Please enter your Password:");

                            #region Mask the input password with STAR
                            string unmaskedPass = "";
                            do
                            {
                                ConsoleKeyInfo key = Console.ReadKey(true);
                                // Backspace Should Not Work
                                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                                {
                                    unmaskedPass += key.KeyChar;
                                    Console.Write("*");
                                }
                                else
                                {
                                    if (key.Key == ConsoleKey.Backspace && unmaskedPass.Length > 0)
                                    {
                                        unmaskedPass = unmaskedPass.Substring(0, (unmaskedPass.Length - 1));
                                        Console.Write("\b \b");
                                    }
                                    else if (key.Key == ConsoleKey.Enter)
                                    {
                                        break;
                                    }
                                }
                            } while (true);
                            #endregion

                            Password = unmaskedPass;
                            passwordEntered = true;
                        }
                        catch (ArgumentException ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(ex.Message);
                            Console.ResetColor();
                        }
                    }
                    #endregion
                }


                driver.FindElementByName("email").SendKeys(Email);
                driver.FindElementByName("password").SendKeys(Password);
                driver.FindElementById("signInSubmit").Click();
                #region Captcha
                try
                {
                    var captchaBox = driver.FindElement(By.Id("auth-captcha-guess"));
                    if (captchaBox != null)
                    {
                        captchaRaised = true;
                        //re-enter the password
                        driver.FindElementByName("password").SendKeys(Password);
                        #region Captcha Validation Check
                        bool captchaEntered = false;
                        while (!captchaEntered)
                        {
                            try
                            {
                                Console.WriteLine();
                                Console.Write("Enter Captcha -- ");
                                Captcha = Console.ReadLine();
                                captchaEntered = true;
                            }
                            catch (ArgumentException ex)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(ex.Message);
                                Console.ResetColor();
                            }
                        }
                        #endregion
                        captchaBox.SendKeys(Captcha);
                        driver.FindElementByCssSelector("#a-autoid-0").Click();
                        captchaDone = true;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                #endregion
                #region OTP
                try
                {
                    driver.FindElementById("auth-mfa-form");

                    #region OTP Validation Check
                    bool otpEntered = false;
                    while (!otpEntered)
                    {
                        try
                        {
                            Console.WriteLine();
                            Console.Write("Enter OTP -- ");
                            OTP = Console.ReadLine();
                            otpEntered = true;
                        }
                        catch (ArgumentException ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(ex.Message);
                            Console.ResetColor();
                        }
                    }
                    #endregion

                    driver.FindElementById("auth-mfa-otpcode").SendKeys(OTP);
                    driver.FindElementById("auth-signin-button").Click();
                    otpDone = true;


                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                #endregion

                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("#sc-navtab-reports")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.CssSelector("#sc-navtab-reports > ul:nth-child(2) > li:nth-child(2) > a:nth-child(1)")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.CssSelector("#report_DetailSalesTrafficByChildItem")).Click();

                try
                {
                    Thread.Sleep(2000);
                    driver.FindElement(By.CssSelector("#fromDate2")).Click();

                }
                catch
                {
                    try
                    {
                        driver.FindElement(By.CssSelector("#report_DetailSalesTrafficByChildItem")).Click();
                        Thread.Sleep(2000);
                        driver.FindElement(By.CssSelector("#fromDate2")).Click();
                    }
                    catch
                    {

                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Network error, Exiting ...");
                        Console.ResetColor();
                        Thread.Sleep(TimeSpan.FromSeconds(5));
                        Environment.Exit(0);
                    }

                }


                int yearCount = 0;
                var prevButton = driver.FindElementByXPath("/html/body/div[3]/div/a[1]");

                bool prevButtonIsNotClickable = prevButton.GetAttribute("class").Contains("ui-state-disabled");

                while (!prevButtonIsNotClickable)
                {
                    string month = driver.FindElementByClassName("ui-datepicker-month").Text;
                    string year = driver.FindElementByClassName("ui-datepicker-year").Text;
                    if (!string.IsNullOrEmpty(lastScrapedDateTimeString))
                    {
                        string[] date = lastScrapedDateTimeString.Split('-');
                        if (month == date[1] && year == date[2])
                        {
                            break;
                        }
                    }
                    prevButton.Click();
                    prevButton = driver.FindElementByXPath("/html/body/div[3]/div/a[1]");
                    prevButtonIsNotClickable = prevButton.GetAttribute("class").Contains("ui-state-disabled");
                    yearCount++;
                }


                Data data = new Data();


                for (int i = 0; i <= yearCount; i++)
                {
                    string month = driver.FindElementByClassName("ui-datepicker-month").Text;
                    string year = driver.FindElementByClassName("ui-datepicker-year").Text;
                    var datePickerTable = driver.FindElementByXPath("/html/body/div[3]/table");
                    var availableDates = driver.FindElements(By.XPath("//*[@class='ui-datepicker-calendar']/tbody/tr/td/a[contains(@class, 'ui-state-default')]")).ToList();
                    for (int j = 0; j < availableDates.Count; j++)
                    {
                        string day = String.Empty;
                        day = availableDates[j].Text;

                        if (day.Length == 1)
                        {
                            day = '0' + day;
                        }
                        var curDay = GetCurrentDatePickerDateTime(day, month, year, driver);
                        if (curDay > lastScrapedDateTime)
                        {
                            availableDates[j].Click();
                            try
                            {
                                Thread.Sleep(1000);
                                var element = driver.FindElement(By.XPath("//*[@id=\"fromDate2\"]"));
                                element.Click();
                            }
                            catch
                            {
                                try
                                {
                                    Thread.Sleep(10000);
                                    var element = driver.FindElement(By.XPath("//*[@id=\"fromDate2\"]"));
                                    element.Click();
                                }
                                catch
                                {
                                    try
                                    {
                                        Thread.Sleep(20000);
                                        var element = driver.FindElement(By.XPath("//*[@id=\"fromDate2\"]"));
                                        element.Click();
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            Thread.Sleep(TimeSpan.FromMinutes(1));
                                            var element = driver.FindElement(By.XPath("//*[@id=\"fromDate2\"]"));
                                            element.Click();
                                        }
                                        catch
                                        {
                                            try
                                            {
                                                Thread.Sleep(TimeSpan.FromMinutes(2));
                                                var element = driver.FindElement(By.XPath("//*[@id=\"fromDate2\"]"));
                                                element.Click();
                                            }
                                            catch
                                            {
                                                Console.ForegroundColor = ConsoleColor.Red;
                                                Console.WriteLine("Network issue occured. Writing all datas scraped till now.");
                                                Console.ResetColor();
                                                data.AllScrapedTillDate = false;
                                            }
                                        }
                                    }

                                }
                            }
                            Thread.Sleep(2000);
                            #region Scrape the table of selected date 
                            var table = driver.FindElementByCssSelector("#dataTable > tbody:nth-child(3)");
                            var tElements = table.FindElements(By.TagName("tr"));
                            foreach (var rows in tElements)
                            {
                                if (rows.FindElements(By.TagName("td")) != null && rows.FindElements(By.TagName("td")).Count > 0)
                                {
                                    Report report = new Report();
                                    foreach (var col in rows.FindElements(By.TagName("td")))
                                    {
                                        //Opting the check input col because we don't need that 
                                        var attributeValue = col.GetAttribute("class");
                                        if (attributeValue.Contains("cbCell"))
                                        {
                                            continue;
                                        }
                                        else if (attributeValue.Contains("_AR_SC_MA_ParentASIN_25990"))
                                        {
                                            report.ParentASIN = col.Text;
                                        }
                                        else if (attributeValue.Contains("_AR_SC_MA_ChildASIN_25991"))
                                        {
                                            report.ChildASIN = col.Text;
                                        }
                                        else if (attributeValue.Contains("_AR_SC_MA_Sessions_25920"))
                                        {
                                            report.Sessions = col.Text;
                                        }
                                        else if (attributeValue.Contains("_AR_SC_MA_UnitsOrdered_40590"))
                                        {
                                            report.UnitsOrdered = col.Text;
                                        }
                                        else if (attributeValue.Contains("_AR_SC_MA_OrderedProductSales_40591"))
                                        {
                                            report.ProductSales = col.Text;
                                        }
                                        else if (attributeValue.Contains("_AR_SC_MA_TotalOrderItems_1"))
                                        {
                                            report.TotalOrderItems = col.Text;
                                        }
                                        else
                                        {
                                            //None 
                                        }
                                    }
                                    report.Date = curDay;
                                    data.LastScraped = DateTime.Now;
                                    data.LastScrapedDatePickerTime = curDay;
                                    data.Reports.Add(report);
                                    WriteToJson(data);
                                }
                            }
                        }
                        #endregion
                        availableDates = driver.FindElements(By.XPath("//*[@class='ui-datepicker-calendar']/tbody/tr/td/a")).ToList();
                    }
                    var nextButton = driver.FindElementByXPath("/html/body/div[3]/div/a[2]");
                    bool nextButtonIsClickable = nextButton.GetAttribute("class").Contains("ui-state-disabled");

                    if (!nextButtonIsClickable)
                    {
                        nextButton.Click();
                    }
                }


                #region Checking if UnicorpLTD is operational & Send scraped data.
                try
                {

                    //Ping pingSender = new Ping();
                    //PingReply reply = pingSender.Send(UnicorpURI);
                    //if (reply.Status == IPStatus.Success)
                    //{
                    //    string result = AddProductToUnicorpAsync(data).Result;
                    //    Console.WriteLine(result);
                    //}
                    //else
                    //{
                    //    Console.ForegroundColor = ConsoleColor.Red;
                    //    Console.WriteLine("UnicorpLTD  is not live, Not sending the scraped informations.");
                    //    Console.ResetColor();
                    //    Thread.Sleep(TimeSpan.FromSeconds(5));
                    //    Environment.Exit(0);
                    //}

                    //Data data = new Data { Reports = reports, LastScraped = DateTime.Now };
                    //string result = AddProductToUnicorpAsync(data).Result;
                    //Console.ForegroundColor = ConsoleColor.Green;
                    //Console.WriteLine(result);
                    //Environment.Exit(0);
                    //Thread.Sleep(TimeSpan.FromSeconds(5));
                    //Environment.Exit(0);
                }
                catch (Exception e)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("UnicorpLTD  is not live, Not sending the scraped informations.");
                    Console.ResetColor();
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                    Environment.Exit(0);
                }

                #endregion
            }
        }

        public static DateTime GetCurrentDatePickerDateTime(string day, string month, string year, ChromeDriver driver)
        {

            string date = day + "-" + month + "-" + year;
            return DateTime.ParseExact(date, "dd-MMMM-yyyy", CultureInfo.InvariantCulture);
        }

        public static void WriteToJson(Data data)
        {
            string startupPath = Directory.GetCurrentDirectory();
            string collectionHistoryPath = Path.Combine(startupPath, "collections.json");

            if (!File.Exists(collectionHistoryPath))
            {
                string json = JsonConvert.SerializeObject(data, Formatting.Indented);
                File.WriteAllText(collectionHistoryPath, json);
            }
            else
            {
                string json = string.Empty;
                using (StreamReader r = new StreamReader(collectionHistoryPath))
                {
                    json = r.ReadToEnd();
                }
                if (string.IsNullOrEmpty(json))
                {
                    json = JsonConvert.SerializeObject(data, Formatting.Indented);
                    File.WriteAllText(collectionHistoryPath, json);
                }
                else
                {
                    Data jdata = JsonConvert.DeserializeObject<Data>(json);
                    jdata.Reports = new List<Report>();
                    foreach (var report in data.Reports)
                    {
                        jdata.Reports.Add(report);
                    }
                    jdata.AllScrapedTillDate = data.AllScrapedTillDate;
                    jdata.LastScraped = data.LastScraped;
                    jdata.LastScrapedDatePickerTime = data.LastScrapedDatePickerTime;
                    var convertedJson = JsonConvert.SerializeObject(jdata, Formatting.Indented);
                    File.WriteAllText(collectionHistoryPath, convertedJson);
                }
            }
        }
        static IConfiguration GetAppConfig()
        {
            string startupPath = Directory.GetCurrentDirectory();

            string configPath = Path.Combine(startupPath, "scraperSettings.json");


            if (!File.Exists(configPath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Config file not found !");
                Console.ResetColor();
                Thread.Sleep(TimeSpan.FromSeconds(5));
                Environment.Exit(0);
            }

            return new ConfigurationBuilder()
                    .AddJsonFile(configPath, optional: true, reloadOnChange: true).Build();
        }
        static IConfiguration GetScrapeInfo()
        {
            string startupPath = Directory.GetCurrentDirectory();
            string configPath = Path.Combine(startupPath, "collections.json");
            if (!File.Exists(configPath))
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("No Previous Scrape Informations !");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Previous Scrape Informations Found !");
                Console.ResetColor();
            }
            return new ConfigurationBuilder()
                    .AddJsonFile(configPath, optional: true, reloadOnChange: true).Build();
        }
        public static async Task<string> AddProductToUnicorpAsync(Data data)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(UnicorpURI);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));



                var content = new StringContent(JsonConvert.SerializeObject(data).ToString(), Encoding.UTF8, "application/json");
                var result = client.PostAsync("api/AddSalesCentralData", content).Result;
                if (result.IsSuccessStatusCode)
                {
                    return await result.Content.ReadAsStringAsync();
                }
            }
            return "";
        }
    }
}
