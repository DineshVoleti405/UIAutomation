
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using OpenQA.Selenium;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class tasks
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void UCITestCreateAccount11()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Contacts");
                xrmApp.Grid.SwitchView("All Contacts");
                xrmApp.Grid.Search("saleem");
                xrmApp.Grid.OpenRecord(0);
                xrmApp.CommandBar.ClickCommand("Assign");
                xrmApp.Dialogs.Assign(Dialogs.AssignTo.User, "CDSReportService-CHE@onmicrosoft.com");
                xrmApp.Lookup.OpenRecord(0);


              // xrmApp.Entity.SetValue(new OptionSet { Name = "rdoMe_id", Value = "1" });
                var xrmBrowser = client.Browser;
                xrmBrowser.Driver.FindAvailable(OpenQA.Selenium.By.XPath("//select[contains(@aria-label='Assign to')"));
                xrmBrowser.Driver.ClickIfVisible(OpenQA.Selenium.By.XPath("//select[contains(@value='1')]"));
 

                var win = xrmBrowser.Driver.SwitchTo().Window(xrmBrowser.Driver.CurrentWindowHandle);

                var elemet = win.FindClickable(OpenQA.Selenium.By.XPath("//input[@aria-label='User or team, Lookup']"));
                elemet.SendKeys("CDSReportService-CHE@onmicrosoft.com");
                xrmApp.Lookup.OpenRecord(0);

                xrmApp.Navigation.OpenSubArea("Sales", "Activities");

                xrmApp.CommandBar.ClickCommand("Email");




                xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));

                xrmApp.Entity.Save();

            }

        }
        [TestMethod]
        public void UCITestCreateAccount1()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Activities");

                xrmApp.CommandBar.ClickCommand("Email");
                xrmApp.Entity.SetValue("to", "saleem");
                xrmApp.Lookup.OpenRecord(0);
                xrmApp.CommandBar.ClickCommand("Insert Signature");
                xrmApp.Entity.SetValue("signatures_id", "signd");
                xrmApp.Lookup.OpenRecord(0);




               // xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));

                xrmApp.Entity.Save();

            }



        }
        [TestMethod]
        public void routing()


        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);



                xrmApp.Navigation.OpenApp(UCIAppName.CustomerService);

                xrmApp.Navigation.OpenSubArea("Service", "Cases");
                xrmApp.Grid.SwitchView("All Cases");
                xrmApp.Grid.Search("testing case");
                xrmApp.Grid.OpenRecord(0);
                xrmApp.CommandBar.ClickCommand("Add to Queue");
              

               
                var xrmBrowser = client.Browser;

                var win = xrmBrowser.Driver.SwitchTo().Window(xrmBrowser.Driver.CurrentWindowHandle);
               var  ele1 = win.FindClickable(OpenQA.Selenium.By.XPath("//input[@aria-label='Queue, Lookup']"));
                ele1.SendKeys("default entity queue");
                xrmApp.Lookup.OpenRecord(0);

             




              



                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));

                xrmApp.Entity.Save();

            }

        }
    }
}