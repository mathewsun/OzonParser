using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace OzonParser;

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
    
    /// <summary>
    /// Из-за непостоянности верстки ozon приходится искать один элемент через несколько XPath
    /// </summary>
    /// <param name="driver">browser driver</param>
    /// <param name="bys">By's</param>
    /// <param name="timeoutInSeconds">Timeout in seconds</param>
    /// <returns></returns>
    public static IWebElement? TryFindWhileBy(this IWebDriver driver, By[] bys, int timeoutInSeconds)
    {
        foreach (var by in bys)
        {
            try
            {
                var elem = driver.FindElement(by, timeoutInSeconds);

                return elem;
            }
            catch
            {
                // ignored
            }
        }

        return null;
    }
}