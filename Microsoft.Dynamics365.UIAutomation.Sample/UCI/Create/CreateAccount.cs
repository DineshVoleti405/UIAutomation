// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class CreateAccountUCI
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void UCITestCreateAccount123()


        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);


              
                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                xrmApp.Navigation.OpenSubArea("Sales", "Activities");

                xrmApp.CommandBar.ClickCommand("Email");
                var xrmBrowser = client.Browser;
                
                var win = xrmBrowser.Driver.SwitchTo().Window(xrmBrowser.Driver.CurrentWindowHandle);

                var elemet = win.FindClickable(OpenQA.Selenium.By.XPath("//input[@aria-label='To, Multiple Selection Lookup']"));
                elemet.SendKeys("saleem");
                xrmApp.Lookup.OpenRecord(0);
                xrmApp.CommandBar.ClickCommand("Insert Template");
                var elemet3 = win.FindClickable(OpenQA.Selenium.By.XPath("//input[@id = 'ChoiceGroup204-OtherSectionID0_2,radio']"));
                elemet3.Click();

               
                xrmApp.CommandBar.ClickCommand("Insert Signature");
                var elemet1 = win.FindClickable(OpenQA.Selenium.By.XPath("//input[@aria-label='Search Signature , Lookup']"));
                elemet1.SendKeys("signd");
                xrmApp.Lookup.OpenRecord(0);
                

                var elemet2 = win.ClickIfVisible(OpenQA.Selenium.By.XPath("//button[contains(@data-id,'select_id')]"));
                //elemet2.Click();
                xrmApp.CommandBar.ClickCommand("Attach File");
               // xrmBrowser.Driver.attachfile();
                xrmApp.Entity.SetValue("filename", "My Active Contacts 9-12-2023 11-36-06 AM.xlsx");
                xrmApp.Entity.Save();
                





           ;




                //xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                //xrmApp.CommandBar.ClickCommand("New");

               // xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5,15));

               // xrmApp.Entity.Save();
                
            }
            
        }
    }
}