
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class Create
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void addtoqueue()


        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);



               

                xrmApp.Navigation.OpenApp(UCIAppName.CustomerService);

                xrmApp.Navigation.OpenSubArea("Service", "Cases");
                xrmApp.Grid.SwitchView("All Cases");
                xrmApp.Grid.Search("testing automation");
                xrmApp.Grid.OpenRecord(0);
                xrmApp.CommandBar.ClickCommand("Add to Queue");



                var xrmBrowser = client.Browser;

                var win = xrmBrowser.Driver.SwitchTo().Window(xrmBrowser.Driver.CurrentWindowHandle);
                var ele1 = win.FindClickable(OpenQA.Selenium.By.XPath("//input[@aria-label='Queue, Lookup']"));
                xrmApp.ThinkTime(1000);
                ele1.SendKeys("default");
                xrmApp.Lookup.OpenRecord(0);
                xrmApp.ThinkTime(1000);
                xrmBrowser.Driver.WaitForPageToLoad();
                

              // var ele3 =  win.FindElement(OpenQA.Selenium.By.XPath("//div[contains(id='dialogFooterContainer_12'])"));
               win.FindClickable(OpenQA.Selenium.By.XPath("//button[@data-id='ok_id']")).Click();
                

                xrmApp.Entity.Save();

            }

        }

        [TestMethod]
        public void commandbar()


        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);



                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.Navigation.OpenSubArea("Sales", "Contacts");
               // xrmApp.CommandBar.ClickCommand("More Command for contact");
                xrmApp.CommandBar.GetCommandValues(true);
                xrmApp.ThinkTime(4000);
                //xrmApp.CommandBar.ClickCommand("Export to Excel");
                xrmApp.CommandBar.ClickCommand("Import from Excel");
                xrmApp.ThinkTime(4000);
                var xrmBrowser = client.Browser;
                var win = xrmBrowser.Driver.SwitchTo().Window(xrmBrowser.Driver.CurrentWindowHandle);

                var ele1 = win.FindElement(OpenQA.Selenium.By.XPath(".//input[@aria-label=\"File Upload\"]")) ;
                
                xrmApp.ThinkTime(4000);
                ele1.Click();
                ele1.SendKeys("My Active Contacts 9-12-2023 11-36-06 AM.xlsx");
                
               // ele1.SendKeys("");
               // xrmBrowser.Driver.






                xrmApp.Entity.Save();

            }

        }
    }
}