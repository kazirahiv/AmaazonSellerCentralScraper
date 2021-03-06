﻿using AmazonScraper.Models;
using AmazonScraper.Models.JSONModels;
using Data;
using Data.Models;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data;
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
            #region Reading configurations from Json Files
            var options = new ChromeOptions();
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.SuppressInitialDiagnosticInformation = true;
            IConfiguration configuration = GetAppConfig();
            var section = configuration.GetSection("UserAuth");
            var emailFromConfig = section.GetValue<string>("Email");
            var passwordFromConfig = section.GetValue<string>("Password");
            IConfiguration scrapeInfo = GetScrapeInfo();
            string lastScrapedDateTimeString = string.Empty;
            DateTime lastScrapedDateTime = DateTime.MinValue;
            try
            {
                if (!string.IsNullOrEmpty(scrapeInfo.GetValue<string>("LastScrapedDatePickerTime")))
                {
                    lastScrapedDateTime = scrapeInfo.GetValue<DateTime>("LastScrapedDatePickerTime");
                }
                else
                {
                    lastScrapedDateTime = GetDateTime("14", "August", "2018");
                }

            }
            catch
            {
                lastScrapedDateTime = GetDateTime("14", "August", "2018");
            }
            if (!string.IsNullOrEmpty(emailFromConfig) && !string.IsNullOrEmpty(passwordFromConfig))
            {
                Email = emailFromConfig;
                Password = passwordFromConfig;
            }
            UnicorpURI = configuration.GetValue<string>("UnicorpURI");
            #endregion

            using (var driver = new ChromeDriver(service, options))
            {
                driver.Navigate().GoToUrl(baseUrl);
                try
                {
                    var signInButton = driver.FindElementByCssSelector("#wp-content > div.as-body.desktop > div.border-color-squid-ink.flex-container.flex-align-items-stretch.flex-align-content-flex-start.flex-full-width.amsg-2018.fonts-loaded.border-color-squid-ink.design-Sell > div > div > div.background-color-aqua.border-color-mermaid.padding-left-xxlarge.padding-right-xxlarge.padding-top-xsmall.padding-bottom-xsmall.flex-container.flex-align-items-center.flex-align-content-flex-start.flex-full-width.amsg-2018.fonts-loaded.border-color-mermaid.design-Sell > div > div.border-color-squid-ink.flex-container.flex-align-items-center.flex-align-content-flex-start.amsg-2018.fonts-loaded.border-color-squid-ink.design-Sell > div:nth-child(1) > div.border-color-squid-ink.padding-right-xsmall.flex-container.flex-align-items-stretch.flex-align-content-flex-start.flex-full-width.amsg-2018.fonts-loaded.border-color-squid-ink.design-Sell > div > a > strong");
                    signInButton.Click();
                }
                catch (Exception e)
                {
                    try
                    {

                        driver.FindElementById("sign-in-button").Click();
                    }
                    catch
                    {
                        try
                        {
                            var signInButton = driver.FindElementByXPath("/html/body/div/div[1]/div/div/div[1]/div[1]/div/div/div[2]/div/div[2]/div[1]/div[1]/div/a/strong");
                            signInButton.Click();
                        }
                        catch
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
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
                    //Console.WriteLine(e.Message);
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


                try
                {
                    var wrongOTPBox = driver.FindElement(By.Id("auth-error-message-box"));
                    if (wrongOTPBox != null)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Wrong OTP entered. Please close and restart the process ..");
                        Environment.Exit(0);
                    }
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e.Message);
                }

                #endregion


                var scrapeToDate = DateTime.Now;


                while (lastScrapedDateTime < scrapeToDate)
                {
                    string date = lastScrapedDateTime.ToString("MM/dd/yyyy");

                    var url = "https://sellercentral.amazon.com/gp/site-metrics/load-report-JSON.html/ref=au_xx_cont_sitereport?sortColumn=12&filterFromDate=" + date + "&filterToDate=" + date + "&fromDate=" + date + "&toDate=" + date + "&cols=/c0/c1/c2/c3/c4/c5/c6/c7/c8/c9/c10/c11&reportID=102:DetailSalesTrafficByChildItem&sortIsAscending=0&currentPage=0&dateUnit=1&viewDateUnits=ALL&runDate=";

                    driver.Navigate().GoToUrl(url);


                    try
                    {
                        var json = driver.FindElementById("sc-content-container").Text;
                        var deserializedJSON = JsonConvert.DeserializeObject<ReportJSON>(json);
                        var rows = deserializedJSON.data.rows;
                        ScrapeDataTable(rows, lastScrapedDateTime);
                    }
                    catch { }



                    lastScrapedDateTime = lastScrapedDateTime.AddDays(1);
                }

                FlagScrapeStatusToJson(true);

                #region Send scraped data to DB And Generate Excel
                try
                {
                    ScrapeData jdata = GetScrapeDataFromJSONRecordFile();

                    if (jdata != null)
                    {
                        if (jdata.AllScrapedTillDate)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Storing about " + jdata.Reports.Count + " Records, Please wait ..");
                            string result = AddProductToDB(jdata);
                            Console.WriteLine(jdata.Reports.Count + " Records stored to AmazonDB.db !");
                            Console.WriteLine(result);
                            Console.WriteLine();
                            Console.WriteLine("Generating Excel Document ...");
                            GenereateExcel();
                            Console.WriteLine("Done !");
                            Thread.Sleep(TimeSpan.FromSeconds(2));
                            Environment.Exit(0);
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Blue;
                            Console.WriteLine("Not All Records Scraped Till Date ! Exitting ...");
                            Console.ResetColor();
                            Thread.Sleep(TimeSpan.FromSeconds(5));
                            Environment.Exit(0);
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Invalid JSON Records. Exitting ...");
                        Console.ResetColor();
                        Thread.Sleep(TimeSpan.FromSeconds(5));
                        Environment.Exit(0);

                    }

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

        private static void ScrapeDataTable(List<List<string>> rows, DateTime date)
        {
            if (rows.Count > 0)
            {
                ScrapeData data = new ScrapeData();
                foreach (var cols in rows)
                {
                    Report report = new Report();
                    report.ParentASIN = cols[0];
                    report.ChildASIN = cols[1];
                    report.Sessions = int.Parse(Regex.Replace(cols[3], @"[^0-9a-zA-Z]+", ""));
                    report.UnitsOrdered = int.Parse(Regex.Replace(cols[8], @"[^0-9a-zA-Z]+", ""));
                    report.ProductSales = decimal.Parse(Regex.Replace(cols[10], @"[^0-9a-zA-Z]+", ""));
                    report.TotalOrderItems = int.Parse(Regex.Replace(cols[11], @"[^0-9a-zA-Z]+", ""));
                    report.Date = date;
                    data.LastScraped = DateTime.Now;
                    data.LastScrapedDatePickerTime = date;
                    data.Reports.Add(report);
                }

                WriteToJson(data);
            }

            Thread.Sleep(1000);
        }




        #region Helper Methods
        public static DateTime GetDateTime(string day, string month, string year)
        {

            string date = day + "-" + month + "-" + year;
            return DateTime.ParseExact(date, "dd-MMMM-yyyy", CultureInfo.InvariantCulture);
        }


        public static ScrapeData GetScrapeDataFromJSONRecordFile()
        {
            string startupPath = Directory.GetCurrentDirectory();
            string collectionHistoryPath = Path.Combine(startupPath, "collections.json");
            if (File.Exists(collectionHistoryPath))
            {
                string json = string.Empty;
                using (StreamReader r = new StreamReader(collectionHistoryPath))
                {
                    json = r.ReadToEnd();
                }
                if (!string.IsNullOrEmpty(json))
                {
                    return JsonConvert.DeserializeObject<ScrapeData>(json);

                }
                return null;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("JSON Record File Doesnt Exist. Exitting ....");
                Console.ResetColor();
                Thread.Sleep(TimeSpan.FromSeconds(3));
                Environment.Exit(0);
            }
            return null;
        }
        public static void FlagScrapeStatusToJson(bool status)
        {
            string startupPath = Directory.GetCurrentDirectory();
            string collectionHistoryPath = Path.Combine(startupPath, "collections.json");

            if (File.Exists(collectionHistoryPath))
            {
                string json = string.Empty;
                using (StreamReader r = new StreamReader(collectionHistoryPath))
                {
                    json = r.ReadToEnd();
                }
                if (!string.IsNullOrEmpty(json))
                {
                    ScrapeData jdata = JsonConvert.DeserializeObject<ScrapeData>(json);
                    jdata.AllScrapedTillDate = status;
                    var convertedJson = JsonConvert.SerializeObject(jdata, Formatting.Indented);
                    File.WriteAllText(collectionHistoryPath, convertedJson);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Got all informations till date. Flagging as COMPLETED ");
                    Console.ResetColor();
                    Thread.Sleep(TimeSpan.FromSeconds(3));
                }
            }
        }
        public static void WriteToJson(ScrapeData data)
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
                ScrapeData jdata = GetScrapeDataFromJSONRecordFile();
                if (jdata != null)
                {
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
                else
                {
                    var json = JsonConvert.SerializeObject(data, Formatting.Indented);
                    File.WriteAllText(collectionHistoryPath, json);

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
        public static string AddProductToDB(ScrapeData data)
        {
            try
            {
                AmazonDBContext amazonDBContext = new AmazonDBContext();
                if (amazonDBContext.Database.EnsureCreated())
                {
                    if (amazonDBContext.Database.CanConnect())
                    {
                        var uniqueProductASINs =
                            data.Reports
                            .GroupBy(s => s.ChildASIN)
                            .Select(s => new UniqueProductASIN { ChildAsinID = s.Key })
                            .ToList();

                        var availableProductInfoOfDates =
                            data.Reports
                            .GroupBy(s => s.Date)
                            .Select(s => new AvailableProductInfoOfDate { DatePickerDate = s.Key })
                            .ToList();

                        if (amazonDBContext.AvailableProductInfoOfDates.Any())
                        {
                            var productInfoOfDates = amazonDBContext.AvailableProductInfoOfDates.AsQueryable();
                            var lastCollectionDateFromDB = productInfoOfDates.OrderByDescending(s => s.DatePickerDate).FirstOrDefault();
                            var lastCollectionDateFromScraper = availableProductInfoOfDates.OrderByDescending(s => s.DatePickerDate).FirstOrDefault();
                            if (lastCollectionDateFromScraper.DatePickerDate > lastCollectionDateFromDB.DatePickerDate)
                            {
                                foreach (var product in uniqueProductASINs)
                                {
                                    amazonDBContext.UniqueProductASINs.Add(product);
                                }
                                foreach (var infoOfDate in availableProductInfoOfDates)
                                {
                                    amazonDBContext.AvailableProductInfoOfDates.Add(infoOfDate);
                                }
                            }
                            else
                            {
                                return "Already have informations till date. Not adding to database.";
                            }
                        }
                        else
                        {

                            foreach (var product in uniqueProductASINs)
                            {
                                amazonDBContext.UniqueProductASINs.Add(product);
                            }
                            foreach (var infoOfDate in availableProductInfoOfDates)
                            {
                                amazonDBContext.AvailableProductInfoOfDates.Add(infoOfDate);
                            }
                        }
                        amazonDBContext.SaveChanges();

                        var childASINSessions =
                            data.Reports
                            .Select(x => new ChildASINSession
                            {
                                ChildASINId = amazonDBContext.UniqueProductASINs.FirstOrDefault(s => s.ChildAsinID == x.ChildASIN).Id,
                                DateID = amazonDBContext.AvailableProductInfoOfDates.FirstOrDefault(s => s.DatePickerDate == x.Date).Id,
                                SessionValue = x.Sessions
                            }).ToList();


                        var unitsOrderedByAsinId =
                            data.Reports
                            .Select(x => new UnitsOrderedByASINID
                            {
                                ChildASINId = amazonDBContext.UniqueProductASINs.FirstOrDefault(s => s.ChildAsinID == x.ChildASIN).Id,
                                DateID = amazonDBContext.AvailableProductInfoOfDates.FirstOrDefault(s => s.DatePickerDate == x.Date).Id,
                                UnitsOrdered = x.UnitsOrdered
                            }).ToList();


                        var productSalesByAsinId =
                            data.Reports
                            .Select(x => new ProductSalesByChildASINID
                            {
                                ChildASINId = amazonDBContext.UniqueProductASINs.FirstOrDefault(s => s.ChildAsinID == x.ChildASIN).Id,
                                DateID = amazonDBContext.AvailableProductInfoOfDates.FirstOrDefault(s => s.DatePickerDate == x.Date).Id,
                                Earning = x.ProductSales
                            }).ToList();

                        var totlaOrderedItemsByAsinId =
                            data.Reports
                            .Select(x => new TotalOrderItemsByASINID
                            {
                                ChildASINId = amazonDBContext.UniqueProductASINs.FirstOrDefault(s => s.ChildAsinID == x.ChildASIN).Id,
                                DateID = amazonDBContext.AvailableProductInfoOfDates.FirstOrDefault(s => s.DatePickerDate == x.Date).Id,
                                TotalOrders = x.UnitsOrdered
                            }).ToList();

                        foreach (var session in childASINSessions)
                        {
                            amazonDBContext.ChildASINSessions.Add(session);
                        }
                        foreach (var unit in unitsOrderedByAsinId)
                        {
                            amazonDBContext.UnitsOrderedByASINIDs.Add(unit);
                        }
                        foreach (var sales in productSalesByAsinId)
                        {
                            amazonDBContext.ProductSalesByChildASINIDs.Add(sales);
                        }
                        foreach (var ordered in totlaOrderedItemsByAsinId)
                        {
                            amazonDBContext.TotalOrderItemsByASINIDs.Add(ordered);
                        }
                        amazonDBContext.SaveChanges();
                        return "New Informations Added To Database";
                    }
                    else
                    {
                        return "Error Storing Data";
                    }
                }
                else
                {
                    return "Error Creating Database";
                }
            }
            catch
            {
                return "Error Connecting To Database";
            }
        }
        public static async Task<string> AddProductToUnicorpAsync(ScrapeData data)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(UnicorpURI);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));



                var content = new StringContent(JsonConvert.SerializeObject(data).ToString(), Encoding.UTF8, "application/json");
                var result = client.PostAsync("api/AddSalesCentralScrapeData", content).Result;
                if (result.IsSuccessStatusCode)
                {
                    return await result.Content.ReadAsStringAsync();
                }
            }
            return "";
        }

        public static void GenereateExcel()
        {

            AmazonDBContext amazonDBContext = new AmazonDBContext();
            string startupPath = Directory.GetCurrentDirectory();
            string configPath = Path.Combine(startupPath, "AmazonScraper.xlsx");
            FileInfo excelFile = new FileInfo(configPath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (amazonDBContext.Database.CanConnect())
            {
                using (ExcelPackage excel = new ExcelPackage())
                {
                    if (amazonDBContext.UniqueProductASINs.Any())
                    {
                        #region Products Excel  
                        excel.Workbook.Worksheets.Add("Products");
                        var headerRowProducts = new List<string[]>()
                          {
                            new string[] { "ID","ChildASINId" }
                          };
                        string pRange = "A1:" + Char.ConvertFromUtf32(headerRowProducts[0].Length + 64) + "1";
                        var worksheetProducts = excel.Workbook.Worksheets["Products"];
                        worksheetProducts.Cells[pRange].LoadFromArrays(headerRowProducts);
                        var pdata = new List<object[]>();
                        foreach (var product in amazonDBContext.UniqueProductASINs.ToList())
                        {
                            pdata.Add(new object[] { product.Id, product.ChildAsinID });
                        }
                        worksheetProducts.Cells[2, 1].LoadFromArrays(pdata);

                        #endregion
                    }
                    if (amazonDBContext.AvailableProductInfoOfDates.Any())
                    {
                        #region DatePickerDates Excel  
                        excel.Workbook.Worksheets.Add("DatePickerDates");
                        var headerRowDates = new List<string[]>()
                          {
                            new string[] { "ID","Date" }
                          };
                        string dRange = "A1:" + Char.ConvertFromUtf32(headerRowDates[0].Length + 64) + "1";
                        var worksheetDates = excel.Workbook.Worksheets["DatePickerDates"];
                        worksheetDates.Cells[dRange].LoadFromArrays(headerRowDates);
                        var ddata = new List<object[]>();
                        foreach (var date in amazonDBContext.AvailableProductInfoOfDates.ToList())
                        {
                            ddata.Add(new object[] { date.Id, date.DatePickerDate.ToShortDateString() });
                        }
                        worksheetDates.Cells[2, 1].LoadFromArrays(ddata);
                        #endregion
                    }
                    if (amazonDBContext.ChildASINSessions.Any())
                    {
                        #region Sessions Excel  
                        excel.Workbook.Worksheets.Add("Sessions");
                        var headerRowSession = new List<string[]>()
                          {
                            new string[] { "ID", "DateId", "ChildAsinId", "SessionValue" }
                          };
                        string sRange = "A1:" + Char.ConvertFromUtf32(headerRowSession[0].Length + 64) + "1";
                        var worksheetSessions = excel.Workbook.Worksheets["Sessions"];
                        worksheetSessions.Cells[sRange].LoadFromArrays(headerRowSession);
                        var sdata = new List<object[]>();
                        foreach (var session in amazonDBContext.ChildASINSessions.ToList())
                        {
                            sdata.Add(new object[] { session.Id, session.DateID, session.ChildASINId, session.SessionValue });
                        }
                        worksheetSessions.Cells[2, 1].LoadFromArrays(sdata);
                        #endregion
                    }
                    if (amazonDBContext.ProductSalesByChildASINIDs.Any())
                    {
                        #region Product sales by ChildASIN
                        excel.Workbook.Worksheets.Add("Product Sales By ChildASINId");
                        var headerRowPSales = new List<string[]>()
                        {
                            new string[] { "ID", "DateId", "ChildAsinId", "Earning" }
                        };
                        string psRange = "A1:" + Char.ConvertFromUtf32(headerRowPSales[0].Length + 64) + "1";
                        var worksheetProductSales = excel.Workbook.Worksheets["Product Sales By ChildASINId"];
                        worksheetProductSales.Cells[psRange].LoadFromArrays(headerRowPSales);
                        var psdata = new List<object[]>();
                        foreach (var sales in amazonDBContext.ProductSalesByChildASINIDs.ToList())
                        {
                            psdata.Add(new object[] { sales.Id, sales.DateID, sales.ChildASINId, sales.Earning });
                        }
                        worksheetProductSales.Cells[2, 1].LoadFromArrays(psdata);
                        #endregion

                    }
                    if (amazonDBContext.TotalOrderItemsByASINIDs.Any())
                    {
                        #region Ordered Items
                        excel.Workbook.Worksheets.Add("Ordered Items");
                        var ordItemsHeaderRow = new List<string[]>()
                  {
                    new string[] { "ID", "DateId", "ChildAsinId", "Total Orders" }
                  };
                        string oIRange = "A1:" + Char.ConvertFromUtf32(ordItemsHeaderRow[0].Length + 64) + "1";
                        var worksheetOrderedItems = excel.Workbook.Worksheets["Ordered Items"];
                        worksheetOrderedItems.Cells[oIRange].LoadFromArrays(ordItemsHeaderRow);
                        var oIData = new List<object[]>();
                        foreach (var sales in amazonDBContext.TotalOrderItemsByASINIDs.ToList())
                        {
                            oIData.Add(new object[] { sales.Id, sales.DateID, sales.ChildASINId, sales.TotalOrders });
                        }
                        worksheetOrderedItems.Cells[2, 1].LoadFromArrays(oIData);
                        #endregion

                    }
                    if (amazonDBContext.UnitsOrderedByASINIDs.Any())
                    {
                        #region Units Ordered
                        excel.Workbook.Worksheets.Add("Units Ordered");
                        var unitsOrderedHeaderRow = new List<string[]>()
                  {
                    new string[] { "ID", "DateId", "ChildAsinId", "Units Ordered" }
                  };
                        string uORange = "A1:" + Char.ConvertFromUtf32(unitsOrderedHeaderRow[0].Length + 64) + "1";
                        var worksheetUnitsOrdered = excel.Workbook.Worksheets["Units Ordered"];
                        worksheetUnitsOrdered.Cells[uORange].LoadFromArrays(unitsOrderedHeaderRow);
                        var uOData = new List<object[]>();
                        foreach (var sales in amazonDBContext.UnitsOrderedByASINIDs.ToList())
                        {
                            uOData.Add(new object[] { sales.Id, sales.DateID, sales.ChildASINId, sales.UnitsOrdered });
                        }
                        worksheetUnitsOrdered.Cells[2, 1].LoadFromArrays(uOData);

                        #endregion
                    }
                    excel.SaveAs(excelFile);
                }
            }

        }
        #endregion
    }
}
