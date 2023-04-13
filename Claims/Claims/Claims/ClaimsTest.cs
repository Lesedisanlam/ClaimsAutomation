using Claims_Testsuite;
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
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office.CustomUI;
using System.Globalization;
using System.Net.Mail;
using OpenQA.Selenium.DevTools.V108.Debugger;
using AngleSharp.Dom;

namespace Claims_Testsuite.Claims
{
   

    [TestFixture]
    public class ClaimsTest : Base
    {
        //Life Assured Details
        public static string Role, ClaimType, Claimant, Cause_of_incident, IdNum, Date_of_incident, Contact_Date, Email_Address, Mobile_Number, ClaimDescription, Gender, Title, Claim_Amount;
        //Banking Details
        public static string Effective_Date, Bank, Branch, Account_Number, Name, Account_Type, Expiry_date, DebitOrderDay;
        //Method Variables
        public static string errMsg, results, Product, PolicyStatus1, PolicyStatus2, LifeA_Relationship, SingleBenefit, PayableAmount, BankDetails_check, Comp_check, Incidents, EventName, ListedEvent, NettInvestment, ClosingBalance, ClosingBalance_short, Transaction_check, Arrears, AmountCalculation, Movement, ClaimStatus, ClaimpaymentStatus;
        public static decimal ClosingBalance_value;
    
    [OneTimeSetUp]
        public void startBrowser()
        {
            createExclReportFile();
            _driver = SiteConnection();
     
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
        private void Nav_ContractSummary()
        {
            try
            {
                //Click Contract Summary
                _driver.FindElement(By.Name("2000175333.8")).Click();
                Delay(2);
            }
            catch
            {
                //Click Main Menu then Contract Summary
                _driver.FindElement(By.Name("CBWeb")).Click();
                Delay(2);
                _driver.FindElement(By.Name("2000175333.8")).Click();
                Delay(2);
            }
        }
        private void Nav_Components()
        {
            try
            {
                //Click Components
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
                Delay(2);
            }
            catch
            {
                try
                {
                    //Expand Contract Summary then click on Components
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
                    Delay(2);
                }
                catch
                {
                    //Click Main Menu then expand Contract Summary then click Components
                    _driver.FindElement(By.Name("CBWeb")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
                    Delay(2);
                }
            }
        }
        private void Nav_Events()
        {
            try
            {
                //Click Events
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();
                Delay(2);
            }
            catch
            {
                try
                {
                    //Expand Contract Summary then click on Events
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();
                    Delay(2);
                }
                catch
                {
                    //Click Main Menu then expand Contract Summary then click Events
                    _driver.FindElement(By.Name("CBWeb")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();
                    Delay(2);
                }
            }
        }
        private void Nav_Transactions()
        {
            try
            {
                //click on transactions
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[7]/tbody/tr/td/a")).Click();
                Delay(3);
            }
            catch (Exception ex)
            {
                try
                {
                    //Expand Contract Summary then click on Transactions
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[7]/tbody/tr/td/a")).Click();
                    Delay(2);
                }
                catch
                {
                    //Click Main Menu then expand Contract Summary then click Events
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/table[7]/tbody/tr/td/a")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[7]/tbody/tr/td/a")).Click();

                }
            }
        }
        private void Nav_AddNewClaim(string scenarioID)
        {
            //Get Product Name
            Product = _driver.FindElement(By.XPath("//html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td[2]")).Text;

            OpenDBConnection("SELECT * FROM ClaimDetails_Data WHERE Scenario_ID = '" + scenarioID + "' ");
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                Role = reader["Role"].ToString().Trim();
                ClaimType = reader["ClaimType"].ToString().Trim();
                Claimant = reader["Claimant"].ToString().Trim();
                Cause_of_incident = reader["Cause_of_incident"].ToString().Trim();
                IdNum = reader["RolePlayer_idNum"].ToString().Trim();
                Date_of_incident = reader["Date_of_incident"].ToString().Trim();
                Contact_Date = reader["Contact_Date"].ToString().Trim();
                Email_Address = reader["Email_Address"].ToString().Trim();
                Mobile_Number = reader["Mobile_Number"].ToString().Trim();
                ClaimDescription = reader["ClaimDescription"].ToString().Trim();
                Gender = reader["Gender"].ToString().Trim();
                Title = reader["Title"].ToString().Trim();
                Claim_Amount = reader["Claim_Amount"].ToString().Trim();
    }
            connection.Close();

            //PolicyStatus
            PolicyStatus1 = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/u/font")).Text;

            Delay(2);

            //Search through list of role players
            for (int i = 0; i < 24; i++)
            {
                IWebElement comp;
                var xPath = "";
                try
                {
                    xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/ td[1]/span/big/b/a";
                    comp = _driver.FindElement(By.XPath(xPath));
                }
                catch (Exception ex)
                {
                    xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/span/a";
                    comp = _driver.FindElement(By.XPath(xPath));
                }
                var compTxt = comp.Text;
                if (compTxt.Contains(Role))
                {
                    Delay(2);
                    comp.Click();
                    var idComp = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[4]/td[4]")).Text;
                    if (!(idComp.Contains(IdNum)))
                    {
                        _driver.Navigate().Back();
                    }
                    else
                    {
                        break;
                    }
                }
            }

            // Life Validation
            LifeA_Relationship = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[8]/td[2]")).Text;

            Delay(2);
            //click on add new claim
            _driver.FindElement(By.Name("btnAddNewClaim")).Click();
            Delay(4);

            //FILL IN CLAIM DETAILS
            //Date of incident:
            _driver.FindElement(By.Name("frmIncidentDate")).Clear();
            Delay(2);
            _driver.FindElement(By.Name("frmIncidentDate")).SendKeys(Date_of_incident);
            Delay(2);

            //First Contact Date:
            _driver.FindElement(By.Name("frmReceivedDate")).Clear();
            Delay(2);
            _driver.FindElement(By.Name("frmReceivedDate")).SendKeys(Contact_Date);
            Delay(6);

            //ClaimType
            SelectElement dropDown = new SelectElement(_driver.FindElement(By.Name("frmClaimType")));
            dropDown.SelectByText(ClaimType);
            Delay(3);

            //Select claimant
            SelectElement dropDown1 = new SelectElement(_driver.FindElement(By.Name("frmClaimant")));
            dropDown1.SelectByText(Claimant);
            Delay(2);

            //Click next
            _driver.FindElement(By.Name("btncbmin2")).Click();
            Delay(2);

            //Select cause incident 
            try
            {
                Delay(2);
                _driver.FindElement(By.XPath("//*[@id='frmCbmin']/tbody/tr[9]/td[2]/nobr/input[2]")).SendKeys(Cause_of_incident);
                Delay(2);
                _driver.FindElement(By.XPath("//*[@id='frmCbmin']/tbody/tr[9]/td[2]/nobr/img")).Click();
                //Mutimediad pop
                String test_url_4_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";
                Assert.AreEqual(2, _driver.WindowHandles.Count);
                var newWindowHandle1 = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle1));
                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                string expectedNewWindowTitle2 = test_url_4_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle1).Title, expectedNewWindowTitle2);

                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr/td/center[2]/table[2]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[2]")).Click();

                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(2);
            }
            catch
            {
            }

            //Click next
            _driver.FindElement(By.Name("btncbmin5")).Click();
            Delay(2);
        }
        private void Nav_PayClaim(string scenarioID)
        {
            Delay(4);

            //new Claim validation
            String claimstatus = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

            if (claimstatus == "New Claim")
            {

                //Hover on claim options
                IWebElement NewClaimElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action.MoveToElement(NewClaimElement).Perform();


            }
            else
            {
                Delay(30);
                _driver.Navigate().Refresh();
                //Hover on claim options
                IWebElement ClaimsOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action.MoveToElement(ClaimsOptionElement).Perform();
            }
            if (Product == "Safrican Siyakhula Invest (5000)" || Product == "Safrican Legacy Saver (6000)")
            {
                //Click authorise claim
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[5]/td/div/div[3]/a/img")).Click();
                Delay(1);
            }
            else
            {
                //Click authorise claim
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[6]/td/div/div[3]/a/img")).Click();
                Delay(1);
            }
            Delay(2);

            ////Click on authorise
            //for (int i = 0; i < 15; i++)
            //{
            //    try
            //    {
            //        _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[{i + 6}]/td/div/div[3]/a")).Click();
            //    }

            //    catch (Exception ex)
            //    {
            //        _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[{i + 5}]/td/div/div[3]/a")).Click();
            //    }
            //    break;
            //}
            //Delay(2);

            //Click on authorise button
            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td/table")).Click();
            Delay(5);

            //Validate Claim status
            string actualvalue2 = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

            actualvalue2.Contains("Authorised Claim");

            //Authorize payment 
            //Hover on claim options
            IWebElement PaymentOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
            //Creating object of an Actions class
            Actions action3 = new Actions(_driver);
            //Performing the mouse hover action on the target element.
            action3.MoveToElement(PaymentOptionElement).Perform();
            Delay(1);

            if (Product == "Safrican Siyakhula Invest (5000)" || Product == "Safrican Legacy Saver (6000)")
            {
                //Click authorise claim
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[6]/td/div/div[3]/a/img")).Click();
                Delay(1);
            }
            else
            {
                //click authorise payment
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[7]/td/div/div[3]/a")).Click();
                Delay(1);
            }
            Delay(2);

            //click "Next" button
            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();
            Delay(2);

            PayableAmount = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[13]/td[2]")).Text;

            //Look to see if there are any beneficiaries without banking details
            OpenDBConnection("SELECT * FROM BankDetails WHERE Scenario_ID = '" + scenarioID + "' ");
            reader = command.ExecuteReader();
            while (reader.Read())
            {

                Effective_Date = reader["Effective_Date"].ToString().Trim();
                Bank = reader["Bank"].ToString().Trim();
                Branch = reader["Branch"].ToString().Trim();
                Account_Number = reader["Account_Number"].ToString().Trim();
                Name = reader["Name"].ToString().Trim();
                Account_Type = reader["Account_Type"].ToString().Trim();
                Expiry_date = reader["Expiry_date"].ToString().Trim();
                DebitOrderDay = reader["Debit_Order_Day"].ToString().Trim();

            }
            connection.Close();

            //Loop through list of payable beneficiaries and add bank details to those who do not have
            try
            {
                for (int i = 2; i < 23; i++)
                {
                    BankDetails_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[{i.ToString()}]/td[7]/em")).Text;
                    if (BankDetails_check == "* Bank Account Required *")
                    {
                        //Click on payment maintenance
                        Delay(2);
                        _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[{i.ToString()}]/td[14]/a")).Click();
                        //Add Additional Bank Account
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[2]/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();

                        //Bank / Retailer:
                        SelectElement dropDown2 = new SelectElement(_driver.FindElement(By.Name("frmEntityObj")));
                        dropDown2.SelectByText(Bank);
                        Delay(2);

                        //Branch:
                        _driver.FindElement(By.Name("frmBranchCode")).SendKeys(Branch);
                        Delay(2);
                        try
                        {
                            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td[2]/nobr/img")).Click();
                            //Mutimediad pop
                            String test_url_4_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";
                            Assert.AreEqual(2, _driver.WindowHandles.Count);
                            var newWindowHandle2 = _driver.WindowHandles[1];
                            Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle2));
                            /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                            string expectedNewWindowTitle2 = test_url_4_title;
                            Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle2).Title, expectedNewWindowTitle2);


                            Delay(2);
                            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr/td/center[2]/table[2]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]")).Click();

                            /* Return to the window with handle = 0 */
                            _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                        }
                        catch
                        {
                        }
                        Delay(8);
                        //Effective Date
                        _driver.FindElement(By.Name("frmEffectiveDate")).SendKeys(Effective_Date);
                        Delay(2);

                        //Account Number
                        _driver.FindElement(By.Name("frmAccountNumber")).SendKeys(Account_Number);
                        Delay(2);

                        //Name
                        _driver.FindElement(By.Name("frmAccountName")).SendKeys(Name);
                        Delay(2);

                        //Account Type
                        SelectElement dropDown3 = new SelectElement(_driver.FindElement(By.Name("frmBankAccountType")));
                        dropDown3.SelectByText(Account_Type);
                        Delay(2);

                        //Click on submit (The Form page)
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[13]/td/table/tbody/tr/td/table")).Click();

                        //Click on submit (To confirm the banking details for that beneficiary)
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                    }
                    //Remember to remove when you want it to look through multiple Beneficiaries
                    //break;
                }
            }
            catch
            {
            }

            //Click next
            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center[1]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();

            //Click on Authorize
            Delay(2);
            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[2]/table")).Click();
            Delay(7);

            //Validate claim status
            ClaimStatus = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

            ClaimStatus.Contains("Payments Created");

            //Hover on claim options
            IWebElement AuthoriseElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
            //Creating object of an Actions class
            Actions action2 = new Actions(_driver);
            //Performing the mouse hover action on the target element.
            action2.MoveToElement(AuthoriseElement).Perform();

            if (Product == "Safrican Siyakhula Invest (5000)" || Product == "Safrican Legacy Saver (6000)")
            {
                //Click on process payment (Investment)
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[9]/td/div/div[3]/a/img")).Click();
                Delay(1);
            }
            else
            {
                //Click on process payment
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[10]/td/div/div[3]/a/img")).Click();
                Delay(4);
            }
            

            //Tick all payable beneficiaries
            try
            {
                for (int i = 2; i < 23; i++)
                {
                    //Click on Confirm Payment textbox
                    _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[{i.ToString()}]/td[5]/input")).Click();
                }

            }
            catch
            {
            }
            Delay(2);

            //Click on "Pay Claim"
            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[1]/td[3]/table")).Click();
            Delay(3);


            //Validate claim status
            ClaimpaymentStatus = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;
            ClaimpaymentStatus.Contains("Claim Payment Raised");


            //Navigate to contract summary screen
            Nav_ContractSummary();
        }
        private void Nav_DataCollection()
        {
            //Store current Policy Status
            PolicyStatus2 = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/u/font")).Text;

            //Product 5000 Accommodation
            try
            {
                NettInvestment = _driver.FindElement(By.XPath(" /html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[3]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;
            }
            catch
            {
            }

            //Check components
            try
            {
                Comp_check = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[6]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td")).Text;
                Assert.That(Comp_check, Is.EqualTo("There are no Active (as at TODAY) components for this policy."));
            }
            catch
            {
                //Navigate to components screen
                Nav_Components();

                //Check if there are still components
                Comp_check = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td")).Text;

                //Navigate back to contract summary
                Nav_ContractSummary();
            }

            //Movement  valdation
            Movement = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[11]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;

            //Incidents  valdation
            Incidents = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;
            EventName = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[5]")).Text;

            //Navigate to Events screen
            Nav_Events();

            //Get today's date for events
            DateTime today = DateTime.Today;
            string Date_check = String.Empty;
            try
            {
                for (int i = 2; i < 23; i++)
                {
                    ListedEvent = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div[2]/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[1]")).Text;
                    Date_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div[2]/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i.ToString()}]/td[2]")).Text;
                    if (ListedEvent == "Death(Acc)" && Date_check == today.ToString("yyyy/MM/dd"))
                    {
                        break;
                    }
                }
            }
            catch
            {
                results = "Failed";
                TakeScreenshot("Claim_EventValidation");
                errMsg = "Correct event was not found";
            }

            //Transactions navigation
            Nav_Transactions();

            if (ClaimType == "PartSurrender" || ClaimType == "Surrender")
            {
                // Select investment Account Type
                Delay(1);
                SelectElement Selectinvestment = new SelectElement(_driver.FindElement(By.Name("frmAccountTypeObj")));
                Selectinvestment.SelectByText("Investment Account (Individual) (SPI)");

                Delay(1);
                _driver.FindElement(By.Name("btncbta20")).Click();

                Delay(1);
                //Get CB
                ClosingBalance = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/center[1]/b")).Text;
                //Remove "R" from CB
                ClosingBalance_short = ClosingBalance.Substring(1);
                //Remove "," from CB
                ClosingBalance_value = decimal.Parse(ClosingBalance_short, CultureInfo.InvariantCulture);
            }
            else
            {
                //Dropdown Selection
                SelectElement dropDown4 = new SelectElement(_driver.FindElement(By.Name("frmAccountTypeObj")));
                dropDown4.SelectByValue("55134.19");
                //Submit
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                //Search through list for Premium Debt amount
                try
                {
                    for (int i = 2; i < 23; i++)
                    {
                        Transaction_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[3]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[{i.ToString()}]/td[2]")).Text;

                        if (Transaction_check == "Premium Debt/Balance")
                        {
                            //Store Premium Debt amount for calculation
                            Arrears = _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[3]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                            break;
                        }
                    }
                }
                catch
                {
                    results = "Failed";
                    TakeScreenshot("Claim_Calculation");
                    errMsg = "Premium Debt/Balance was not found";
                }

                //Calculation
                AmountCalculation = (Convert.ToDecimal(SingleBenefit) - Convert.ToDecimal(Arrears)).ToString("#,##0.00");

            }
        }
        private void Validations(string scenarioID, string contractRef)
        {
            //VALIDATIONS!!!!
            //Check if status is correct
            switch (PolicyStatus2)
            {
                case "In-Force":
                    //In-Force + Components Present = Passed
                    if (Comp_check != "There are no Active (as at TODAY) components for this policy." || Comp_check != "There are no components for this policy.")
                    {
                        results = "Passed";
                    }
                    else
                    {
                        results = "Failed";
                        TakeScreenshot(contractRef);
                        errMsg = "Incorrect Policy Status";
                    }
                    break;
                case "Out-of-Force":
                    //Out-of-Force + No Components = Passed
                    if (Comp_check == "There are no Active (as at TODAY) components for this policy." || Comp_check == "There are no components for this policy.")
                    {
                        results = "Passed";
                    }
                    else
                    {
                        results = "Failed";
                        TakeScreenshot(contractRef);
                        errMsg = "Incorrect Policy Status";
                    }
                    break;
                case "Premium Waiver":
                    //Premium Waiver + Components Present = Passed
                    if (Comp_check != "There are no Active (as at TODAY) components for this policy." || Comp_check != "There are no components for this policy.")
                    {
                        results = "Passed";
                    }
                    else
                    {
                        results = "Failed";
                        TakeScreenshot(contractRef);
                        errMsg = "Incorrect Policy Status";
                    }
                    break;
                case "Surrendered":
                    //Surrendered + Components Present = Passed
                    if (Comp_check == "There are no Active (as at TODAY) components for this policy." || Comp_check == "There are no components for this policy.")
                    {
                        results = "Passed";
                    }
                    else
                    {
                        results = "Failed";
                        TakeScreenshot(contractRef);
                        errMsg = "Incorrect Policy Status";
                    }
                    break;
            }

            //Check if claim was succesfully done
            //Funeral Products
            switch (Product)
            {
                case "Safrican Serenity Funeral Premium (1000)":
                case "Safrican Serenity Funeral (2000)":
                case "Safrican Just Funeral (3000)":
                    //Product 1000, 2000 & 3000
                    if ((ClaimpaymentStatus == "Claim Payment Raised" || ClaimpaymentStatus == "Authorised Claim") & (Incidents == "Claim Payment Raised" || Incidents == "Authorised Claim") & (Movement == "Death Claim" || Movement == "Death(Acc)") & (EventName == ListedEvent) & (AmountCalculation == PayableAmount))
                    {
                        //Successful Claim
                        results = "Passed";
                    }
                    else
                    {
                        results = "Failed";
                        TakeScreenshot(contractRef);
                        errMsg = "Claim  did not meet all validation criteria";
                    }
                    break;
                case "Safrican Provide and Protect Plan (4000)":
                    //Product 4000 STILL NEED TO DO!!!!!!!!!!!!!!!

                    break;
                case "Safrican Siyakhula Invest (5000)":
                case "Safrican Legacy Saver (6000)":
                    //Product 5000 & 6000
                    if ((PolicyStatus2 == "Surrendered") && (Incidents == "Claim Payment Raised") && (Movement == "Surrender") && (ClosingBalance_short == NettInvestment))
                    {
                        //Successful Claim)
                        results = "Passed";
                    }
                    else if ((PolicyStatus2 == "In-Force") && (Incidents == "PartSurrender") && (Movement == "Part Surrender") && (ClosingBalance_value >= 1000.00m))
                    {
                        //Successful Claim)
                        results = "Passed";
                    }
                    else
                    {
                        results = "Failed";
                        TakeScreenshot(contractRef);
                        errMsg = "Claim was not successfull (Product 5000)";
                    }
                    break;
            }
        }

            //Shaq method - works
            [Test, TestCaseSource("GetTestData", new object[] { "SSFP_Claim" })]
        public void SSFP_Claim(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }

            try
            {
                //Navigate to desired policy
                policySearch(contractRef);
                
                //Navigate to Add New Claim screen
                Nav_AddNewClaim(scenarioID); 

                //Code Input unique to SSFP
                //Select ARL-BI_Number
                Random rnd = new Random();
                int myRandomNo = rnd.Next(1000, 9999); // creates a 8 digit random no.
                myRandomNo.ToString();
                _driver.FindElement(By.Name("frmCriterionValue1_1")).SendKeys("BI-1663" + myRandomNo.ToString());
                Delay(2);

                //Select ID-Number	
                _driver.FindElement(By.Name("frmCriterionValue1_2")).SendKeys(IdNum);
                Delay(2);

                //Click Next
                _driver.FindElement(By.Name("btncbmin9")).Click();
                Delay(2);

                SingleBenefit = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[4]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[2]")).Text;

                //Click Finish
                _driver.FindElement(By.Name("btncbmin12")).Click();
                Delay(4);

                //Navigation through claim payment
                Nav_PayClaim(scenarioID);

                //Navigation through data collection for validations
                Nav_DataCollection();

                //Run collected data through product specific validations
                Validations(scenarioID, contractRef);
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

        //Shaq method - works
        [Test, TestCaseSource("GetTestData", new object[] { "SSF_Claim" })]
        public void SSF_Claim(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }

            try
            {
                //Navigate to desired policy
                policySearch(contractRef);

                //Navigate to Add New Claim screen
                Nav_AddNewClaim(scenarioID);

                //Code Input unique to SSFP
                //Select ARL-BI_Number
                Random rnd = new Random();
                int myRandomNo = rnd.Next(1000, 9999); // creates a 8 digit random no.
                myRandomNo.ToString();
                _driver.FindElement(By.Name("frmCriterionValue1_1")).SendKeys("BI-1663" + myRandomNo.ToString());
                Delay(2);

                //Select ID-Number	
                _driver.FindElement(By.Name("frmCriterionValue1_2")).SendKeys(IdNum);
                Delay(2);

                //Click Next
                _driver.FindElement(By.Name("btncbmin9")).Click();
                Delay(2);

                SingleBenefit = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[4]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[2]")).Text;

                //Click Finish
                _driver.FindElement(By.Name("btncbmin12")).Click();
                Delay(4);

                //Navigation through claim payment
                Nav_PayClaim(scenarioID);

                //Navigation through data collection for validations
                Nav_DataCollection();

                //Run collected data through product specific validations
                Validations(scenarioID, contractRef);
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

        //Shaq method - works
        [Test, TestCaseSource("GetTestData", new object[] { "SJF_Claim" })]
        public void SJF_Claim(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }

            try
            {
                //Navigate to desired policy
                policySearch(contractRef);

                //Navigate to Add New Claim screen
                Nav_AddNewClaim(scenarioID);

                //Code Input unique to SSFP
                //Select ARL-BI_Number
                Random rnd = new Random();
                int myRandomNo = rnd.Next(1000, 9999); // creates a 8 digit random no.
                myRandomNo.ToString();
                _driver.FindElement(By.Name("frmCriterionValue1_1")).SendKeys("BI-1663" + myRandomNo.ToString());
                Delay(2);

                //Select ID-Number	
                _driver.FindElement(By.Name("frmCriterionValue1_2")).SendKeys(IdNum);
                Delay(2);

                //Click Next
                _driver.FindElement(By.Name("btncbmin9")).Click();
                Delay(2);

                SingleBenefit = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[4]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[2]")).Text;

                //Click Finish
                _driver.FindElement(By.Name("btncbmin12")).Click();
                Delay(4);

                //Navigation through claim payment
                Nav_PayClaim(scenarioID);

                //Navigation through data collection for validations
                Nav_DataCollection();

                //Run collected data through product specific validations
                Validations(scenarioID, contractRef);
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

        //Kamo method - pretty sure doesn't work
        [Test, TestCaseSource("GetTestData", new object[] { "SPPP_Claim" })]
        public void SPPP_Claim(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }


            string errMsg = String.Empty;
            string results = String.Empty;
            try
            {

                var Arrears = String.Empty;
                var SingleBenefit = String.Empty;
                var PayableAmount = String.Empty;
                var Policystatus1 = String.Empty;
                var Policystatus2 = String.Empty;
                var Product = String.Empty;
                var amountCalculation = String.Empty;


                Delay(2);
                //SetproductName();
                string Role = String.Empty, Claimant = String.Empty, Cause_of_incident = String.Empty, BI_Number = String.Empty, Roleplayer = String.Empty, SubClaimType = String.Empty, ClaimType = String.Empty,
                IdNum = String.Empty, Date_of_incident = String.Empty, Contact_Date = String.Empty, Email_Address = String.Empty, Mobile_Number = String.Empty, ClaimDescription = String.Empty, Gender = String.Empty, Title = String.Empty;
                string Comp_check = String.Empty;
                string Description_check = String.Empty;

                policySearch(contractRef);
                Product = _driver.FindElement(By.XPath("//html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td[2]")).Text;

                OpenDBConnection("SELECT * FROM ClaimDetails_Data WHERE Scenario_ID = '" + scenarioID + "' ");
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    Role = reader["Role"].ToString().Trim();
                    ClaimType = reader["ClaimType"].ToString().Trim();
                    Claimant = reader["Claimant"].ToString().Trim();
                    Cause_of_incident = reader["Cause_of_incident"].ToString().Trim();
                    IdNum = reader["RolePlayer_idNum"].ToString().Trim();
                    Date_of_incident = reader["Date_of_incident"].ToString().Trim();
                    Contact_Date = reader["Contact_Date"].ToString().Trim();
                    Email_Address = reader["Email_Address"].ToString().Trim();
                    Mobile_Number = reader["Mobile_Number"].ToString().Trim();
                    ClaimDescription = reader["ClaimDescription"].ToString().Trim();
                    Gender = reader["Gender"].ToString().Trim();
                    Title = reader["Title"].ToString().Trim();

                }
                connection.Close();


                // PolicyStatus
                Policystatus1 = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/u/font")).Text;


                Delay(2);

                for (int i = 0; i < 24; i++)
                {
                    IWebElement comp;
                    var xPath = "";
                    try
                    {
                        xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/ td[1]/span/big/b/a";
                        comp = _driver.FindElement(By.XPath(xPath));
                    }
                    catch (Exception ex)
                    {
                        xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/span/a";
                        comp = _driver.FindElement(By.XPath(xPath));
                    }
                    var compTxt = comp.Text;
                    if (compTxt.Contains(Role))
                    {
                        Delay(2);
                        comp.Click();
                        var idComp = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[4]/td[4]")).Text;
                        if (!(idComp.Contains(IdNum)))
                        {
                            _driver.Navigate().Back();
                        }
                        else
                        {
                            break;
                        }
                    }
                }



                // Life Validation
                string LifeA_Relationship = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[8]/td[2]")).Text;

                Delay(2);
                //click on add new claim
                _driver.FindElement(By.Name("btnAddNewClaim")).Click();


                Delay(4);

                //Date of incident:
                _driver.FindElement(By.Name("frmIncidentDate")).Clear();
                Delay(2);
                _driver.FindElement(By.Name("frmIncidentDate")).SendKeys(Date_of_incident);
                Delay(2);


                //First Contact Date:
                _driver.FindElement(By.Name("frmReceivedDate")).Clear();
                Delay(2);
                _driver.FindElement(By.Name("frmReceivedDate")).SendKeys(Contact_Date);
                Delay(2);

                //ClaimType
                SelectElement dropDown = new SelectElement(_driver.FindElement(By.Name("frmClaimType")));
                dropDown.SelectByText(ClaimType);
                Delay(3);

                //Select claimant
                SelectElement dropDown1 = new SelectElement(_driver.FindElement(By.Name("frmClaimant")));
                dropDown1.SelectByText(Claimant);
                Delay(2);

                //Click next
                _driver.FindElement(By.Name("btncbmin2")).Click();
                Delay(2);


                //Select cause incident 
                try
                {
                    Delay(2);
                    _driver.FindElement(By.XPath("//*[@id='frmCbmin']/tbody/tr[9]/td[2]/nobr/input[2]")).SendKeys(Cause_of_incident);
                    Delay(2);
                    _driver.FindElement(By.XPath("//*[@id='frmCbmin']/tbody/tr[9]/td[2]/nobr/img")).Click();
                    //Mutimediad pop
                    String test_url_4_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";
                    Assert.AreEqual(2, _driver.WindowHandles.Count);
                    var newWindowHandle1 = _driver.WindowHandles[1];
                    Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle1));
                    /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                    string expectedNewWindowTitle2 = test_url_4_title;
                    Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle1).Title, expectedNewWindowTitle2);


                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr/td/center[2]/table[2]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[2]")).Click();

                    /* Return to the window with handle = 0 */
                    _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                }
                catch
                {
                }

                Delay(2);

                //Click next
                _driver.FindElement(By.Name("btncbmin5")).Click();
                Delay(2);






                //Select ARL-BI_Number
                Random rnd = new Random();
                int myRandomNo = rnd.Next(1000, 9999); // creates a 8 digit random no.
                myRandomNo.ToString();
                _driver.FindElement(By.Name("frmCriterionValue1_1")).SendKeys("BI-1663" + myRandomNo.ToString());
                Delay(2);

                //Select ID-Number	
                _driver.FindElement(By.Name("frmCriterionValue1_2")).SendKeys(IdNum);
                Delay(2);


                //Click Next
                _driver.FindElement(By.Name("btncbmin9")).Click();
                Delay(2);

                SingleBenefit = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[4]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[2]")).Text;

                //Click Finish
                _driver.FindElement(By.Name("btncbmin12")).Click();
                Delay(4);

                //new Claim validation
                String claimstatus = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

                if (claimstatus == "New Claim")
                {

                    //Hover on claim options
                    IWebElement NewClaimElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                    //Creating object of an Actions class
                    Actions action = new Actions(_driver);
                    //Performing the mouse hover action on the target element.
                    action.MoveToElement(NewClaimElement).Perform();


                }
                else
                {


                    Delay(30);
                    _driver.Navigate().Refresh();
                    //Hover on claim options
                    IWebElement ClaimsOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                    //Creating object of an Actions class
                    Actions action = new Actions(_driver);
                    //Performing the mouse hover action on the target element.
                    action.MoveToElement(ClaimsOptionElement).Perform();



                }

                Delay(2);
                //Click on authorise
                for (int i = 0; i < 15; i++)
                {
                    try
                    {
                        _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[{i + 6}]/td/div/div[3]/a")).Click();
                    }

                    catch (Exception ex)
                    {
                        _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[{i + 5}]/td/div/div[3]/a")).Click();
                    }
                    break;
                }
                Delay(2);

                //Click on authorise button
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td/table")).Click();
                Delay(5);

                //Validate Claim status
                string actualvalue2 = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

                actualvalue2.Contains("Authorised Claim");



                //Authorize payment 
                //Hover on claim options
                IWebElement PaymentOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action3 = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action3.MoveToElement(PaymentOptionElement).Perform();
                Delay(1);


                //click authorise payment
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[7]/td/div/div[3]/a")).Click();
                Delay(1);

                //click "Next" button
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();
                Delay(2);

                PayableAmount = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[13]/td[2]")).Text;

                try
                {
                    //Click on payment maintenance
                    Delay(2);
                    _driver.FindElement(By.Name("hl_AuthPay")).Click();

                    //Click on submit
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table")).Click();

                    //click on "Next" button
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center[1]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();


                }
                catch
                {
                    //Go back two screens
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr/td/div/div/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[4]/td/div/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();

                    string bankdetails = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[7]/em")).Text;

                    //Validate bank details 
                    if (bankdetails == "* Bank Account Required *")

                    {
                        //Authorise payment
                        string Effective_Date = String.Empty, Bank = String.Empty, Branch = String.Empty, Account_Number = String.Empty, Name = String.Empty, Account_Type = String.Empty,
                        credit_Card = String.Empty, DebitOrderDay = String.Empty, Expiry_date = String.Empty;

                        OpenDBConnection("SELECT * FROM BankDetails");
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {

                            Effective_Date = reader["Effective_Date"].ToString().Trim();
                            Bank = reader["Bank"].ToString().Trim();
                            Branch = reader["Branch"].ToString().Trim();
                            Account_Number = reader["Account_Number"].ToString().Trim();
                            Name = reader["Name"].ToString().Trim();
                            Account_Type = reader["Account_Type"].ToString().Trim();
                            Expiry_date = reader["Expiry_date"].ToString().Trim();
                            DebitOrderDay = reader["Debit_Order_Day"].ToString().Trim();

                        }
                        connection.Close();

                        //Click on payment maintenance
                        Delay(2);
                        _driver.FindElement(By.Name("hl_AuthPay")).Click();
                        //Add Additional Bank Account
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[2]/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();

                        //Bank / Retailer:
                        SelectElement dropDown2 = new SelectElement(_driver.FindElement(By.Name("frmEntityObj")));
                        dropDown2.SelectByText(Bank);
                        Delay(5);

                        //Branch:
                        _driver.FindElement(By.Name("frmBranchCode")).SendKeys(Branch);
                        Delay(2);
                        try
                        {
                            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td[2]/nobr/img")).Click();
                            //Mutimediad pop
                            String test_url_4_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";
                            Assert.AreEqual(2, _driver.WindowHandles.Count);
                            var newWindowHandle2 = _driver.WindowHandles[1];
                            Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle2));
                            /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                            string expectedNewWindowTitle2 = test_url_4_title;
                            Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle2).Title, expectedNewWindowTitle2);


                            Delay(2);
                            _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr/td/center[2]/table[2]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]")).Click();

                            /* Return to the window with handle = 0 */
                            _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                        }
                        catch
                        {
                        }

                        //Effective Date
                        _driver.FindElement(By.Name("frmStopDate")).SendKeys(Effective_Date);
                        Delay(2);

                        //Account Number:	
                        _driver.FindElement(By.Name("frmAccountNumber")).SendKeys(Account_Number);
                        Delay(4);

                        //Name:	
                        _driver.FindElement(By.Name("frmAccountName")).SendKeys(Name);
                        Delay(4);

                        //Clear Stop Date
                        _driver.FindElement(By.Name("frmStopDate")).Clear();
                        Delay(2);

                        //Click on submit
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[13]/td/table/tbody/tr/td/table")).Click();

                        //Click on submit
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();

                        //Click next
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();

                    }
                    else
                    {


                        //Click on  submit
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center[1]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();


                    }

                }

                //Click on  Authorize
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[2]/table")).Click();
                Delay(1);

                //Validate claim status
                string ClaimStatus = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

                ClaimStatus.Contains("Payments Created");

                //Process Payment


                //Hover on claim options
                IWebElement AuthoriseElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action2 = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action2.MoveToElement(AuthoriseElement).Perform();




                //Click on process payment
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[10]/td/div/div[3]/a/img")).Click();
                Delay(4);


                //Click on Confirm Payment textbox

                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[5]/input")).Click();
                Delay(3);

                //Click on process payment button
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[1]/td[3]/table")).Click();
                Delay(3);




                //Validate claim status
                string ClaimpaymentStatus = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

                ClaimpaymentStatus.Contains("Claim Payment Raised");


                clickOnMainMenu();



                //Click on contract summary
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[3]")).Click();
                Delay(3);


                Policystatus2 = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/u/font")).Text;

                //Check components
                try
                {
                    Comp_check = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[6]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td")).Text;
                    Assert.That(Comp_check, Is.EqualTo("There are no Active (as at TODAY) components for this policy."));
                }
                catch
                {
                    //Navigate to components screen
                    try
                    {
                        //click on components
                        _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
                        Delay(3);
                    }
                    catch (Exception ex)
                    {
                        //expand contract sumary
                        _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                        //click on components
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[3]/tbody/tr/td/a")).Click();
                    }
                    Comp_check = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td")).Text;
                    //Navigate back to contract summary
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[3]/a")).Click();
                }

                // movement  valdation
                string movement = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[11]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;

                // Incidents  valdation
                string Incidents = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;

                string eventname = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[5]")).Text;

                //Navigate to Events screen
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();
                string events = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div[2]/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;

                //Transactions navigation
                try
                {
                    //click on transactions
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[7]/tbody/tr/td/a")).Click();
                    Delay(3);
                }
                catch (Exception ex)
                {
                    //expand Main Menu
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/table[7]/tbody/tr/td/a")).Click();
                    //expand contract sumary
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();
                    //click on transactions
                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[7]/tbody/tr/td/a")).Click();
                }
                //Dropdown Selection
                SelectElement dropDown4 = new SelectElement(_driver.FindElement(By.Name("frmAccountTypeObj")));
                dropDown4.SelectByValue("55134.19");
                //Submit
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td/span/a")).Click();
                //Search through list for Premium Debt amount
                try
                {
                    for (int i = 2; i < 23; i++)
                    {
                        Description_check = _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[3]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[{i.ToString()}]/td[2]")).Text;

                        if (Description_check == "Premium Debt/Balance")
                        {
                            //Store Premium Debt amount for calculation
                            Arrears = _driver.FindElement(By.XPath($"/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[3]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[{i.ToString()}]/td[3]")).Text;
                            break;
                        }
                    }
                }
                catch
                {
                    results = "Failed";
                    TakeScreenshot("Claim_Calculation");
                    errMsg = "Premium Debt/Balance was not found";
                }

                //Calculation
                string AmountCalculation = (Convert.ToDecimal(SingleBenefit) - Convert.ToDecimal(Arrears)).ToString("#,##0.00");

                //Check if product has prem waver


                //VALIDATIONS!

                //PRODUCT 1000 VALIDATIONS
                //Prod 1000, Self, with other components
                if ((LifeA_Relationship == "Self" & Product == "Safrican Serenity Funeral Premium (1000)") && (Comp_check != "There are no Active (as at TODAY) components for this policy." & Comp_check != "There are no components for this policy."))
                {
                    if (ClaimType == "PremWaiver-Death")
                    {
                        Assert.That(Policystatus2, Is.EqualTo("Premium Waiver"));
                    }
                    else
                    {
                        Assert.That(Policystatus2, Is.EqualTo("Out-of-Force"));
                    }
                }

                //PRODUCT 3000 VALIDATIONS
                else if (LifeA_Relationship == "Self" & Product == "Safrican Just Funeral (3000)")
                {

                    Assert.That(Policystatus2, Is.EqualTo("Out-of-Force"));

                }

                //VALIDATION THAT APPLIES FOR ALL PRODUCTS
                else if (Comp_check == "There are no Active (as at TODAY) components for this policy." || Comp_check == "There are no components for this policy.")
                {

                    Assert.That(Policystatus2, Is.EqualTo("Out-of-Force"));

                }

                else

                {

                    Assert.That(Policystatus2, Is.EqualTo("In-Force"));

                }

                if ((ClaimpaymentStatus == "Claim Payment Raised") & (Incidents == "Claim Payment Raised") & (movement == "Death Claim" || movement == "Death(Acc)") & (eventname == events) & (AmountCalculation == PayableAmount))
                {
                    //Sucessful Claim
                    results = "Passed";
                }
                else
                {
                    results = "Failed";
                    TakeScreenshot(contractRef);
                    errMsg = "Claim  did not meet all validation criteria";
                }
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

        //Minimum Part Surrender amount is 1000!
        [Test, TestCaseSource("GetTestData", new object[] { "SSI_Claim" })]
        public void SSI_Manual_Claim(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            try
            {
                //Navigate to desired policy
                policySearch(contractRef);

                //Navigate to Add New Claim screen
                Nav_AddNewClaim(scenarioID);

                //IF statement accomodating PartSurrender
                if (ClaimType == "PartSurrender")
                {
                    _driver.FindElement(By.Name("frmDisinvestAmount")).Clear();
                    Delay(2);
                    _driver.FindElement(By.Name("frmDisinvestAmount")).SendKeys(Claim_Amount);
                    Delay(2);
                }

                //Click next on screen showing surrender amount
                _driver.FindElement(By.Name("btncbmin9")).Click();
                Delay(2);

                //Click Finish
                _driver.FindElement(By.Name("btncbmin12")).Click();
                Delay(2);

                //Authorise and Pay
                Nav_PayClaim(scenarioID);

                //Navigation through data collection for validation
                Nav_DataCollection();

                //PRODUCT 5000 VALIDATIONS (SSI)
                Validations(scenarioID, contractRef);

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

        //Minimum Part Surrender amount is 1000!
        [Test, TestCaseSource("GetTestData", new object[] { "SLS_Claim" })]
        public void SLS_Claim(string contractRef, string scenarioID)
        {
            if (String.IsNullOrEmpty(contractRef))
            {
                Assert.Ignore();
            }
            try
            {
                //Navigate to desired policy
                policySearch(contractRef);

                //Navigate to Add New Claim screen
                Nav_AddNewClaim(scenarioID);

                //IF statement accomodating PartSurrender
                if (ClaimType == "PartSurrender")
                {
                    _driver.FindElement(By.Name("frmDisinvestAmount")).Clear();
                    Delay(2);
                    _driver.FindElement(By.Name("frmDisinvestAmount")).SendKeys(Claim_Amount);
                    Delay(2);
                }

                //Click next on screen showing surrender amount
                _driver.FindElement(By.Name("btncbmin9")).Click();
                Delay(2);

                //Click Finish
                _driver.FindElement(By.Name("btncbmin12")).Click();
                Delay(2);

                //Authorise and Pay
                Nav_PayClaim(scenarioID);

                //Navigation through data collection for validation
                Nav_DataCollection();

                //PRODUCT 5000 VALIDATIONS (SSI)
                Validations(scenarioID, contractRef);

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
           //EmailReport();
            DisconnectBrowser();
          
        }
    }
}