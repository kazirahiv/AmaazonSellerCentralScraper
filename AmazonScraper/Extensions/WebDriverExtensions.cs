using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;

namespace AmazonScraper.Extensions
{
    public static class WebDriverExtensions
    {
        public static IWebElement FindElement(this IWebDriver driver, By by, int timeoutInSeconds)
        {
            if (timeoutInSeconds > 0)
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                return wait.Until(drv => drv.FindElement(by));
            }
            return driver.FindElement(by);
        }

        public static void WaitForDocumentReady(this IWebDriver driver)
        {
            Console.WriteLine("Waiting for five instances of document.readyState returning 'complete' at 100ms intervals.");
            IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;
            int i = 0; // Count of (document.readyState === complete) && (ae.isProcessing === false)
            int j = 0; // Count of iterations in the while() loop.
            int k = 0; // Count of times i was reset to 0.
            bool readyState = false;
            while (i < 5)
            {
                System.Threading.Thread.Sleep(100);
                readyState = (bool)jse.ExecuteScript("return ((document.readyState === 'complete') && (ae.isProcessing === false))");
                if (readyState) { i++; }
                else
                {
                    i = 0;
                    k++;
                }
                j++;
                if (j > 300) { throw new TimeoutException("Timeout waiting for document.readyState to be complete."); }
            }
            j *= 100;
            Console.WriteLine("Waited " + j.ToString() + " milliseconds. There were " + k + " resets.");
        }
    }
}
