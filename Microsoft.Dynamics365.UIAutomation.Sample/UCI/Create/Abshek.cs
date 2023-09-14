
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class practiceA
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void methodpractice()


        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);



                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.Navigation.OpenSubArea("Sales", "Contacts");

                xrmApp.Grid.OpenRecord(0);
                xrmApp.Entity.SelectTab("Related","Activities");
                xrmApp.RelatedGrid.ClickCommand("New Activity", "Task");
                //xrmApp.RelatedGrid.SelectTab("New Activity","Task");
                xrmApp.QuickCreate.SetValue("subject", "Sample");
                xrmApp.QuickCreate.Save();
                xrmApp.ThinkTime(2000);

            }


        }
        [TestMethod]
        public void methodpractice1()


        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);



                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.Navigation.OpenSubArea("Sales", "Leads");

                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.Entity.SetValue("subject", TestSettings.GetRandomString(5, 15));

                xrmApp.Entity.SetValue("lastname", TestSettings.GetRandomString(5, 15));

                xrmApp.Entity.Save();

                xrmApp.CommandBar.ClickCommand("Qualify");
                xrmApp.ThinkTime(2000);

            }


        }

        [TestMethod]
        public void UCITestCreateAccount()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);


                xrmApp.Navigation.OpenApp(UCIAppName.CustomerService);

                xrmApp.Navigation.OpenSubArea("Service", "Cases");
                xrmApp.CommandBar.ClickCommand("New Case");
                xrmApp.Entity.SetValue("title", "Automation2");
                xrmApp.Entity.SetValue("customerid", "Fabrikam, Inc.");
                xrmApp.Lookup.OpenRecord(0);


                xrmApp.CommandBar.ClickCommand("Save & Close");
                xrmApp.Grid.HighLightRecord(0);
                xrmApp.CommandBar.ClickCommand("Edit");
               // xrmApp.CommandBar.c
              


                
                

                //xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                //xrmApp.Navigation.OpenSubArea("Sales", "Competitors");

                //xrmApp.CommandBar.ClickCommand("New");

                //xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));

                //xrmApp.Entity.Save();

            }

        }


    }
}