using  TestBase;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using Actions = OpenQA.Selenium.Interactions.Actions;


namespace Claims_Testsuite.Claims
{

    [TestFixture]
    public class ClaimsTest : Base
    {

            private string sheet;
            [OneTimeSetUp]
            public void startBrowser()
            {
                _driver = SiteConnection();
                sheet = "Claims";
            }
        private void policySearch(string contractRef)
        {
            Delay(4);
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
            //Click on contract search 
            try
            {
                _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            }
            catch (Exception ex)
            {
                clickOnMainMenu();
                _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            }
            Delay(2);
            //Type in contract ref 
            _driver.FindElement(By.Name("frmContractReference")).SendKeys(contractRef);
            Delay(4);
            //Click on Search Icon 
            _driver.FindElement(By.Name("btncbcts0")).Click();
            Delay(5);
            _driver.FindElement(By.XPath("//*[@id='AppArea']/table[2]/tbody/tr[2]/td[1]/a")).Click();
            Delay(5);
        }
        private void clickOnMainMenu()
            {
                try
                {
                    //find the contract search option
                    var search = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td/div[7]/table[4]/tbody/tr/td/a"));
                }
                catch
                {
                    //clickOnMainMenu
                    _driver.FindElement(By.Name("CBWeb")).Click();
                }
            }

        [Test, TestCaseSource("GetTestData", new object[] { "SSFP_Claim" })]
        public void SSFP_Claim(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }


            string errMsg = String.Empty;
            string results = String.Empty;
            try
            {
                
                policySearch(contractRef);
                Delay(2);
                SetproductName();
                string Claimant = String.Empty, Cause_of_incident = String.Empty, BI_Number = String.Empty, Roleplayer = String.Empty ,SubClaimType = String.Empty, ClaimType = String.Empty, IdNum = String.Empty;

                OpenDBConnection("SELECT * FROM SSLPData");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    ClaimType = reader["ClaimType"].ToString().Trim();
                    Claimant = reader["Claimant"].ToString().Trim();
                    Cause_of_incident = reader["Cause_of_incident"].ToString().Trim();
                    BI_Number = reader["BI_Number"].ToString().Trim();
                    Roleplayer = reader["Roleplayer"].ToString().Trim();
                    IdNum = reader["RolePlayer_idNum"].ToString().Trim();
                }
                connection.Close();
                Delay(2);
                for (int i = 0; i < 24; i++)
                {
                    IWebElement comp;
                    var xPath = "";
                    try
                    { 
                    
                        xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/span/big/b/a";
                        comp = _driver.FindElement(By.XPath(xPath));
                    }
                    catch (Exception ex)
                    {
                        xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/span/big/b/a";
                        comp = _driver.FindElement(By.XPath(xPath));
                    }
                    var compTxt = comp.Text;
                    if (compTxt.Contains(Roleplayer))
                    {
                        Delay(2);
                        comp.Click();
                        var idComp = _driver.FindElement(By.XPath("//html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[4]/td[4]")).Text;
                        if (!(idComp.Contains(IdNum)))
                        {
                            _driver.Navigate().Back();
                        }
                        break;
                    }
                }



                Delay(2);
                //click on add new claim
                _driver.FindElement(By.Name("btnAddNewClaim")).Click();
                Delay(4);

                var value = String.Empty;
                //Insert Reinstate reason
                switch (ClaimType)
                {
                    case ("Accidental Death Claim"):
                        value = "Death(Acc),02104,966307650.488";
                        break;
                    case ("Manual(Mem)"):
                        value = "Manual(Mem),03112,966307657.488";
                        break;
                    case ("Non-Accidental Death"):
                        value = "Death(Non-Acc),01104,966307650.488";
                        break;
                    default:
                        break;
                }
                SelectElement dropDown = new SelectElement(_driver.FindElement(By.Name("frmClaimType")));
                dropDown.SelectByValue(value);
                Delay(5);

                _driver.FindElement(By.XPath("//div[3]/a/img")).Click();
                Delay(4);
                //Date of incident:
                _driver.FindElement(By.Name("frmIncidentDatee")).Clear();
                Delay(2);
                _driver.FindElement(By.Name("frmIncidentDate")).SendKeys("");
                Delay(2);

                //First Contact Date:
                _driver.FindElement(By.Name("frmReceivedDate")).Clear();
                _driver.FindElement(By.Name("frmReceivedDate")).SendKeys("");


                //Insert Reinstate reason
                switch (SubClaimType)
                {
                    case ("Accidental Death Claim"):
                        value = "Death(Acc),02104,966307650.488";
                        break;
                    case ("Manual(Mem)"):
                        value = "Manual(Mem),03112,966307657.488";
                        break;
                    case ("Non-Accidental Death"):
                        value = "Death(Non-Acc),01104,966307650.488";
                        break;
                    default:
                        break;
                }
                SelectElement dropDown2 = new SelectElement(_driver.FindElement(By.Name("frmClaimType")));
                dropDown2.SelectByText(SubClaimType);

                Delay(4);
                _driver.FindElement(By.Name("btncbmcc13")).Click();
                Delay(4);
                _driver.FindElement(By.Name("btncbmcc17")).Click();





            }
            catch (Exception ex)
            {
                if (ex.Message.Length > 255)
                {
                    errMsg = ex.Message.Substring(0, 255);
                }
                else
                {
                    errMsg = ex.Message;
                }
                results = "Failed";
            }
            writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
            Assert.IsTrue(results.Equals("Passed"));
        }


        [OneTimeTearDown]
        public void closeBrowser()
        {
            DisconnectBrowser();
        }
    }
}