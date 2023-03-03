﻿using Claims_Testsuite;
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

namespace Claims_Testsuite.Claims
{

     [TestFixture]
    public class ClaimsTest : Base
    {

        private string sheet;
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
        [Test, TestCaseSource("GetTestData", new object[] { "SSF_Claim" })]
        public void SSF_Claim(string contractRef, string scenarioID)
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

                policySearch(contractRef);
                Product = _driver.FindElement(By.XPath("//html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[1]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td[2]")).Text;

                Delay(2);
                Arrears = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[3]/td/div/table/tbody/tr/td/span/table/tbody/tr/td[2]/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td[2]")).Text;

                OpenDBConnection("SELECT * FROM ClaimDetails_Data WHERE Scenario_ID = '" + scenarioID + "' ");
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    Role = reader["Role"].ToString().Trim();
                    ClaimType = reader["ClaimType"].ToString().Trim();
                    Claimant = reader["Claimant"].ToString().Trim();
                    Cause_of_incident = reader["Cause_of_incident"].ToString().Trim();
                    Roleplayer = reader["Roleplayer"].ToString().Trim();
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
                        xPath = $"/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[5]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[{i + 2}]/td[1]/span/big/b/a";
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



                // Life  valdation
                string LifeA = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[8]/td[2]")).Text;

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
                Delay(5);



                //Select claimant
                SelectElement dropDown1 = new SelectElement(_driver.FindElement(By.Name("frmClaimant")));
                dropDown1.SelectByText(Claimant);

                Delay(4);



                //Click next
                _driver.FindElement(By.Name("btncbmin2")).Click();
                Delay(4);


                //go to incedent 
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

                //validate payout amount


                //Click on authorise
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td/table")).Click();
                Delay(5);



                // Authorise Claim authorization

                //Validate Claim status
                string actualvalue2 = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

                actualvalue2.Contains("Authorised Claim");



                //Aurthorize papyment 
                //Hover on claim options
                IWebElement PaymentOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action3 = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action3.MoveToElement(PaymentOptionElement).Perform();
                Delay(1);


                //click authrise payment
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[7]/td/div/div[3]/a")).Click();
                Delay(1);



                //click authrise Next
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();

                Delay(2);

                PayableAmount = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[13]/td[2]")).Text;




                try

                {


                    //Click on  submit
                    Delay(2);
                    _driver.FindElement(By.Name("hl_AuthPay")).Click();

                    Delay(2);
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table")).Click();


                    //click authrise Next
                    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center[1]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();


                }
                catch
                {

                    string bankdetails = _driver.FindElement(By.XPath("/html/bodchy/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[7]/em")).Text;

                    //Validate ank details 
                    if (bankdetails == "* Bank Account Required *")

                    {
                        //Authorisze payment
                        string Effective_Date = String.Empty, Bank = String.Empty, Branch = String.Empty, Account_Number = String.Empty, Name = String.Empty, Account_Type = String.Empty,
                        Stop_Date = String.Empty, Cheque_Stale_Months = String.Empty, credit_Card = String.Empty, Expiry_date = String.Empty;

                        OpenDBConnection("SELECT * FROM ClaimBankdetails");
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {

                            Effective_Date = reader["Effective_Date"].ToString().Trim();
                            Bank = reader["Bank"].ToString().Trim();
                            Branch = reader["Branch"].ToString().Trim();
                            Account_Number = reader["Account_Number"].ToString().Trim();
                            Name = reader["Name"].ToString().Trim();
                            Account_Type = reader["Account_Type"].ToString().Trim();
                            Stop_Date = reader["Stop_Date"].ToString().Trim();
                            Cheque_Stale_Months = reader["Cheque_Stale_Months"].ToString().Trim();
                            credit_Card = reader["credit_Card"].ToString().Trim();
                            Expiry_date = reader["Expiry_date"].ToString().Trim();

                        }

                        connection.Close();


                        //Add  payments 
                        //Click on  submit
                        Delay(2);
                        _driver.FindElement(By.Name("hl_AuthPay")).Click();

                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table")).Click();


                        //add bank details 
                        //Click 
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center[1]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();


                        //Bank / Retailer:
                        SelectElement dropDown2 = new SelectElement(_driver.FindElement(By.Name("frmEntityObj")));
                        dropDown2.SelectByText(bankdetails);
                        Delay(5);

                        //Branch:
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td[2]/nobr/img")).Click();


                        //Mutimediad pop
                        String test_url_5_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";


                        Assert.AreEqual(2, _driver.WindowHandles.Count);
                        var newWindowHandle2 = _driver.WindowHandles[1];
                        Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle1));
                        /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                        string expectedNewWindowTitle3 = test_url_5_title;
                        Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle2).Title, expectedNewWindowTitle3);


                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr/td/center[2]/table[2]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[1]")).Click();

                        /* Return to the window with handle = 0 */
                        _driver.SwitchTo().Window(_driver.WindowHandles[0]);

                        Delay(2);


                        //Account Number:	
                        _driver.FindElement(By.Name("frmAccountNumber")).SendKeys(Account_Number);
                        Delay(2);

                        //Name:	
                        _driver.FindElement(By.Name("frmAccountName")).SendKeys(Name);
                        Delay(2);

                        //Type:	
                        //Cheque Stale Months:	
                        //Default for Owner?	
                        //Stop Date:
                        _driver.FindElement(By.Name("frmStopDate")).SendKeys(Stop_Date);
                        Delay(2);
                        //Pick a date
                        //Credit Card Type:	
                        //Credit Card Expiry Date:	

                        //Click on  submit
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table/tbody/tr[4]/td[2]/span/table/tbody/tr[13]/td/table/tbody/tr/td/table")).Click();


                        //Click on  submit
                        Delay(2);
                        _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center[1]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[1]/table")).Click();
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



                // Authorise Claim validation

                //Validate calim status
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


                // movement  valdation
                string movement = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[11]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;


                // Incidents  valdation
                string Incidents = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;



                string eventname = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[5]")).Text;


                //viewed events valdation

                //
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[1]/a")).Click();

                //

                //*[@id="t0_756"]/table/tbody/tr/td[1]/a
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/div[1]/table[9]/tbody/tr/td/a")).Click();


                string events = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/div/center/div[2]/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[1]")).Text;
                
                string AmountCalculation = (Convert.ToDecimal(SingleBenefit) - Convert.ToDecimal(Arrears)).ToString("#,##0.00");




                if (LifeA == "Self" & Product == "Safrican Just Funeral (3000)")
                {
                    //policy status
                    // PolicyStatus

                    Assert.That(Policystatus2, Is.EqualTo("Out-of-Force"));



                }
                else if ((LifeA == "Self" & Product == "Safrican Serenity Funeral Premium"))
                {

                    Assert.That(Policystatus2, Is.EqualTo("Premium Waiver"));


                }

                else

                {


                    Assert.That(Policystatus2, Is.EqualTo("In-Force"));


                }

                if ((ClaimpaymentStatus == "Claim Payment Raised") & (Incidents == "Claim Payment Raised") & (movement == "Death Claim") & (eventname == events) & (AmountCalculation == PayableAmount))
                {
                    //Sucessful Claim)
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

        //[Test, TestCaseSource("GetTestData", new object[] { "Workflow_Claim" })]
        //public void Workflow_Claim(string contractRef, string scenarioID)
        //{
        //    if (String.IsNullOrEmpty(contractRef))
        //    {
        //        Assert.Ignore();
        //    }


        //    string errMsg = String.Empty;
        //    string results = String.Empty;
        //    try
        //    {

        //        policySearch(contractRef);
        //        Delay(2);
        //        //SetproductName();
        //        string Claimant = String.Empty, Cause_of_incident = String.Empty, BI_Number = String.Empty, Roleplayer = String.Empty, SubClaimType = String.Empty, ClaimType = String.Empty,
        //        IdNum = String.Empty, Date_of_incident = String.Empty, Contact_Date = String.Empty, Email_Address = String.Empty, Mobile_Number = String.Empty, ClaimDescription = String.Empty, Gender = String.Empty, Title = String.Empty;

        //        OpenDBConnection("SELECT * FROM SSLP_Data");
        //        reader = command.ExecuteReader();
        //        while (reader.Read())
        //        {

        //            ClaimType = reader["ClaimType"].ToString().Trim();
        //            Claimant = reader["Claimant"].ToString().Trim();
        //            Cause_of_incident = reader["Cause_of_incident"].ToString().Trim();
        //            BI_Number = reader["BI_Number"].ToString().Trim();
        //            Roleplayer = reader["Roleplayer"].ToString().Trim();
        //            IdNum = reader["RolePlayer_idNum"].ToString().Trim();
        //            Date_of_incident = reader["Date_of_incident"].ToString().Trim();
        //            Contact_Date = reader["Contact_Date"].ToString().Trim();
        //            Email_Address = reader["Email_Address"].ToString().Trim();
        //            Mobile_Number = reader["Mobile_Number"].ToString().Trim();
        //            ClaimDescription = reader["ClaimDescription"].ToString().Trim();
        //            Gender = reader["Gender"].ToString().Trim();
        //            Title = reader["Title"].ToString().Trim();

        //        }
        //        connection.Close();

        // Delay(2);
        // //click on add Digital  Stack
        // _driver.FindElement(By.Name("sv_Main")).Click();
        // Delay(2);

        // //click on Call centre
        // _driver.FindElement(By.Name("cc_Main")).Click();
        // Delay(2);

        // //click on add Add call log  
        // _driver.FindElement(By.Name("cc_Proc_cctcl")).Click();
        // Delay(2);
        // //click on  Call template 
        // _driver.FindElement(By.Name("cc_Template")).Click();
        // Delay(2);
        // //click on MIP Sanlam
        // _driver.FindElement(By.Name("cc_MIP")).Click();
        // Delay(2);
        // //click on   Claims
        // _driver.FindElement(By.Name("cc_MIP_Claims")).Click();
        // Delay(2);
        // //click on   claim
        // _driver.FindElement(By.Name("cc_MIP_Claims_AC_NewClaim")).Click();
        // Delay(2);
        // //click on complete
        // _driver.FindElement(By.Name("btnComplete")).Click();
        // Delay(2);



        // //ClaimType
        // SelectElement dropDown = new SelectElement(_driver.FindElement(By.Name("refActivityLogRefsMainReqClaimType")));
        // dropDown.SelectByText(ClaimType);
        // Delay(5);

        // //click on Yes BI Number
        // _driver.FindElement(By.XPath("//*[@id='refActivityLogRefsMainHaveBINumber1']")).Click();
        // Delay(1);

        // //click on YES Death Certificate
        // _driver.FindElement(By.XPath("//*[@id='refActivityLogRefsMainHaveDeathCertificate1']")).Click();
        // Delay(2);


        // //click on YES ID Document
        // _driver.FindElement(By.XPath("//*[@id='refActivityLogRefsMainHaveIDDocument1']")).Click();
        // Delay(2);


        // //click on Contract selection

        // _driver.FindElement(By.Name("remlActivityLogRemsMaincbmct1")).Click();
        // Delay(2);


        // //Claims  

        // String test_url_2_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";


        // Assert.AreEqual(2, _driver.WindowHandles.Count);
        // var newWindowHandle = _driver.WindowHandles[1];
        // Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle));
        // /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
        // string expectedNewWindowTitle = test_url_2_title;
        // Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle).Title, expectedNewWindowTitle);
        // Delay(1);
        // _driver.FindElement(By.Name("frmContractReference")).SendKeys(contractRef);
        // Delay(2);
        // _driver.FindElement(By.Name("frmContractReference")).Click();

        // Delay(2);
        // _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[2]")).Click();//*[@id="lkpResultsTable"]/tbody/tr[2]
        // /* Return to the window with handle = 0 */
        // _driver.SwitchTo().Window(_driver.WindowHandles[0]);
        // Delay(5);

        // //Click on Complete
        // _driver.FindElement(By.XPath("//*[@id='stateimg-5']")).Click();

        // //Click on Related Entities
        // Delay(2);
        // _driver.FindElement(By.XPath("//*[@id='stateimg-6']")).Click();


        // try
        // {
        //     //Click on Add new
        //     Delay(2);
        //     _driver.FindElement(By.Name("fcRemLabel1")).Click();
        // }
        // catch
        // {

        //     //Click on Related Entities
        //     Delay(2);
        //     _driver.FindElement(By.XPath("//*[@id='stateimg-6']")).Click();
        // }
        // //*[@id="stateimg-6"]

        // //Click on Add new
        // Delay(2);
        // _driver.FindElement(By.Name("fcRemLabel1")).Click();

        // //Mutimediad pop
        // String test_url_3_title = "SANLAM RM - Safrican Retail";


        // Assert.AreEqual(2, _driver.WindowHandles.Count);
        // var newWindowHandle1 = _driver.WindowHandles[1];
        // Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle1));
        // /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
        // string expectedNewWindowTitle1 = test_url_3_title;
        // Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle1).Title, expectedNewWindowTitle1);


        ////add death certificate 


        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/div/span[1]")).Click();
        // //Click on 
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/div/span[2]")).Click();
        // //Click on 
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/div/span[2]")).Click();
        // //Click on 
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/ul/li[6]/div/span")).Click();
        // //Click on 
        // Delay(2);
        // IWebElement element = _driver.FindElement(By.Name("ffFile"));
        // element.SendKeys("C:\\Users\\G992107\\Downloads\\UPload file.pdf");


        // Delay(4);
        // _driver.FindElement(By.Name("btnSubmit")).Click();

        // //click on add 
        // Delay(4);
        // _driver.FindElement(By.Name("btnAdd")).Click();

        // //Add Decease  Idetification


        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/div/span[1]")).Click();
        // //Click on 
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/div/span[2]")).Click();
        // //Click on 
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/div/span[2]")).Click();

        // //Click on Decease  Idetification
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/ul/li[5]/div/span")).Click();
        // //Click on 
        // Delay(2);
        // IWebElement element2 = _driver.FindElement(By.Name("ffFile"));
        // element2.SendKeys("C:\\Users\\G992107\\Downloads\\UPload file.pdf");


        // Delay(4);
        // _driver.FindElement(By.Name("btnSubmit")).Click();

        // //click on add 
        // Delay(4);  
        // _driver.FindElement(By.Name("btnAdd")).Click();


        // //Add BI-1663

        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/div/span[1]")).Click();

        // //Click on 
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/div/span[2]")).Click();
        // //Click on 
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/div/span[2]")).Click();

        // //Click on Add BI-1663
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/ul/li[3]/div/span")).Click();
        // //Click on 
        // Delay(2);
        // IWebElement element1 = _driver.FindElement(By.Name("ffFile"));
        // element1.SendKeys("C:\\Users\\G992107\\Downloads\\UPload file.pdf");

        // //cxlick on close
        // Delay(3);
        // _driver.FindElement(By.Name("btnSubmit")).Click();

        // //cxlick on close
        // Delay(3);
        // _driver.FindElement(By.Name("btnClose")).Click();
        // /* Return to the window with handle = 0 */
        // _driver.SwitchTo().Window(_driver.WindowHandles[0]);


        // //click on complete
        // Delay(2);
        //_driver.FindElement(By.Name("btnComplete")).Click();
        // //Click on Add new



        // Delay(2);
        // _driver.FindElement(By.Name("fcIDNumber")).SendKeys(IdNum);


        // //Click on search
        // Delay(2);
        // _driver.FindElement(By.Name("fcPersonLkp")).Click();


        // Delay(4);

        // //Mutimediad pop
        // String test_url_4_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";


        // Assert.AreEqual(2, _driver.WindowHandles.Count);
        // var newWindowHandle3 = _driver.WindowHandles[1];
        // Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle1));
        // /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
        // string expectedNewWindowTitle4 = test_url_4_title;
        // Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle3).Title, expectedNewWindowTitle4);

        // var Firstname = _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[2]/td[2]")).Text;

        // var Surname = _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[2]/td[3]")).Text;

        // //Click on Add new
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr/td/center[2]/table[2]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]")).Click();


        // /* Return to the window with handle = 0 */
        // _driver.SwitchTo().Window(_driver.WindowHandles[0]);

        // //incednt Date
        // Delay(2);
        // _driver.FindElement(By.Name("fcIncidentDate")).SendKeys(Date_of_incident);

        // //incednt Date
        // Delay(2);
        // _driver.FindElement(By.Name("fcFirstContact")).SendKeys(Date_of_incident);

        // //click on add 
        // Delay(4);
        // //_driver.FindElement(By.Name("fcLifeAssured")).Click();




        // try

        // {


        //     //Gender 
        //     SelectElement userGender = new SelectElement(_driver.FindElement(By.Name("fcGender")));
        //     userGender.SelectByText(Gender);
        //     Delay(2);

        //     //Firstname
        //     Delay(2);
        //     _driver.FindElement(By.Name("fcFirstName")).SendKeys(Firstname);

        //     //Surname
        //     Delay(2);
        //     _driver.FindElement(By.Name("fcLastName")).SendKeys(Surname);


        //     //Title
        //     SelectElement dropDown5 = new SelectElement(_driver.FindElement(By.Name("fcTitle")));
        //     dropDown5.SelectByText(Title);

        // }
        // catch
        // {


        // }


        // IWebElement ele = _driver.FindElement(By.Name("fcEmailAddress")); //find the text field

        // if (ele.GetAttribute("value") == "")
        // {

        //     //Email_Address
        //     Delay(2);
        //     _driver.FindElement(By.Name("fcEmailAddress")).SendKeys(Email_Address);

        //     //Mobile_Number 
        //     Delay(2);
        //     _driver.FindElement(By.Name("fcMobileNumber")).SendKeys(Mobile_Number);
        //     Delay(2);
        // }
        // else
        // {
        //     //Store the value
        //     String store = ele.GetAttribute("value");
        // }



        // //Life assured
        // SelectElement dropDown2 = new SelectElement(_driver.FindElement(By.Name("fcLifeAssured")));
        // dropDown2.SelectByText(Claimant);
        // Delay(2);


        // //ClaimType Person
        // SelectElement dropDown1 = new SelectElement(_driver.FindElement(By.Name("fcClaimType")));
        // dropDown1.SelectByText(Cause_of_incident);
        // Delay(2);

        // //Cause of Incident
        // SelectElement dropDown3 = new SelectElement(_driver.FindElement(By.Name("fcIncidentCause")));
        // dropDown3.SelectByText(ClaimDescription);
        // Delay(2);

        // //BI-number 
        // Delay(2);
        // _driver.FindElement(By.Name("fcBINumber")).SendKeys(BI_Number);




        // //Click submit
        // Delay(2);
        // _driver.FindElement(By.Name("btnSubmit")).Click();



        // Delay(3);

        // //process claim

        // clickOnMainMenu();



        // //click on contrct summary
        // Delay(2);
        // _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[3]/a")).Click();


        // Delay(2);

        // //Validate calim status
        // string actualvalue = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]")).Text;

        // actualvalue.Contains("New Claim");





        // Delay(90);
        // _driver.Navigate().Refresh();



        // //go to workflow 
        // Delay(8);



        //    //go to workflow 
        //    String expectedtext = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[19]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[3]")).Text;

        //    if (expectedtext == "Active")
        //    {

        //        _driver.FindElement(By.Name("fcReference1")).Click();


        //    }
        //    else
        //    {


        //        Delay(90);
        //        _driver.Navigate().Refresh();
        //        _driver.FindElement(By.Name("fcReference1")).Click();


        //    }

        //    //click component om yellow element 


        //    IWebElement iframe = _driver.FindElement(By.XPath("//div[@id='AppArea']//center"));

        //    //Now we can click the button

        //    IWebElement mapObject = _driver.FindElement(By.XPath("//*[name()='svg']//*[local-name()='g']//*[local-name()='path']"));
        //    Actions builder = new Actions(_driver);
        //    builder.ContextClick(mapObject).Perform();


        //    //click on red error messagew
        //    IWebElement m = _driver.FindElement(By.XPath("//*[local-name()='svg']"));
        //    // Action class to move and click element
        //    Actions a = new Actions(_driver);
        //    a.MoveToElement(m).
        //    Click().Build().Perform();





        //    Delay(3);
        //    _driver.FindElement(By.Id("a_node23")).Click();


        //    //click on pick
        //    Delay(3);
        //    _driver.FindElement(By.Name("btnPick")).Click();

        //    //click on Death Certificate tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainDeath Certificate")).Click();


        //    //click on Certified Copy of Identity Document tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainCertified Copy of Identity Document")).Click();


        //    //click on Marriage Certificate or Proof of Customary Union	 tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainMarriage Certificate or Proof of Customary Union")).Click();


        //    //click on Application Form - Claim tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainApplication Form - Claim")).Click();



        //    //click on Medical Report or Medical Attendance Certificate tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainDeath Notification Form(BI - 1663 or DHA - 1663)")).Click();




        //    //click on Additional Medicals  tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainMedical Report or Medical Attendance Certificate")).Click();

        //    //click on Additional Medicals (eg MRI/Failures/Cancer/Transplant)tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainAuthority or Executorship Notification")).Click();
        //    //click on  complete
        //    Delay(3);
        //    _driver.FindElement(By.Name("btnComplete")).Click();


        //    //click on pick tickbox
        //    Delay(3);
        //    _driver.FindElement(By.Name("")).Click();


        //    string Informant_Information = String.Empty, Undertaker_Name = String.Empty, BI_SerialNumber = String.Empty, PlaceOfDeath = String.Empty, Primary_CauseOfDeath = String.Empty, Secondary_CauseOfDeath = String.Empty,
        //    Medical_SAMDC = String.Empty, DoctorName = String.Empty, Doctor_PracticeNumber = String.Empty, Doctor_ContactNumber = String.Empty, Traditional_Healer = String.Empty;

        //    OpenDBConnection("SELECT * FROM Reference_Values");
        //    reader = command.ExecuteReader();
        //    while (reader.Read())
        //    {


        //        Informant_Information = reader["Informant_Information"].ToString().Trim();
        //        Undertaker_Name = reader["Undertaker_Name"].ToString().Trim();
        //        BI_SerialNumber = reader["BI_SerialNumber"].ToString().Trim();
        //        PlaceOfDeath = reader["PlaceOfDeath"].ToString().Trim();
        //        Primary_CauseOfDeath = reader["Primary_CauseOfDeath"].ToString().Trim();
        //        Secondary_CauseOfDeath = reader["Secondary_CauseOfDeath"].ToString().Trim();
        //        Medical_SAMDC = reader["Medical_SAMDC"].ToString().Trim();
        //        DoctorName = reader["DoctorName"].ToString().Trim();
        //        Doctor_PracticeNumber = reader["Doctor_PracticeNumber"].ToString().Trim();
        //        Doctor_ContactNumber = reader["Doctor_ContactNumber"].ToString().Trim();
        //        Traditional_Healer = reader["Traditional_Healer"].ToString().Trim();


        //    }
        //    connection.Close();


        //    //Informant_Information	
        //    SelectElement dropDown4 = new SelectElement(_driver.FindElement(By.Name("refActivityLogRefsMainInformantSameClaimant")));
        //    dropDown4.SelectByText(Informant_Information);
        //    Delay(2);


        //    //BI Certificate Information	
        //    //Undertaker Name:	
        //    Delay(2);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainUndertakerName")).SendKeys(Undertaker_Name);

        //    //BI_SerialNumber:	
        //    Delay(2);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainBISerialNumber")).SendKeys(BI_SerialNumber);

        //    //PlaceOfDeath:	
        //    Delay(2);
        //    _driver.FindElement(By.Name("refActivityLogRefsMainBIPlaceOfDeath")).SendKeys(PlaceOfDeath);

        //    //Primary_CauseOfDeath:	
        //    SelectElement CauseOfDeath = new SelectElement(_driver.FindElement(By.Name("refActivityLogRefsMainBIPlaceOfDeath")));
        //    CauseOfDeath.SelectByText(Primary_CauseOfDeath);
        //    Delay(2);

        //    //Secondary_CauseOfDeath:	
        //    Delay(2);
        //    _driver.FindElement(By.Name("fcBINumber")).SendKeys(Secondary_CauseOfDeath);

        //    //Medical_SAMDC:	
        //    Delay(2);
        //    _driver.FindElement(By.Name("fcBINumber")).SendKeys(Medical_SAMDC);




        //    //Practitioner Information
        //    //DoctorName:
        //    Delay(2);
        //    _driver.FindElement(By.Name("fcBINumber")).SendKeys(DoctorName);

        //    //Doctor_PracticeNumber:	

        //    Delay(2);
        //    _driver.FindElement(By.Name("fcBINumber")).SendKeys(Doctor_PracticeNumber);

        //    //Doctor_ContactNumber:	
        //    Delay(2);
        //    _driver.FindElement(By.Name("fcBINumber")).SendKeys(Doctor_ContactNumber);

        //    //Traditional_Healer:
        //    SelectElement dropDown8 = new SelectElement(_driver.FindElement(By.Name("fcGender")));
        //    dropDown8.SelectByText(Traditional_Healer);
        //    Delay(2);




        //    //validation of claim



        //    //Claim




        //    //Hover on claim options
        //    IWebElement ClaimsOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
        //    //Creating object of an Actions class
        //    Actions action = new Actions(_driver);
        //    //Performing the mouse hover action on the target element.
        //    action.MoveToElement(ClaimsOptionElement).Perform();
        //    Delay(3);

        //    //Click on authorise
        //    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[7]/td/div/div[3]/a/img")).Click();
        //    Delay(5);


        //    // Authorise Claim authorization

        //    //Validate calim status
        //    string actualvalue2 = _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[5]/td[2]")).Text;

        //    actualvalue2.Contains("Authorised Claim");


        //    //Add  payments 
        //    //Click on  payment maintence
        //    Delay(2);
        //    _driver.FindElement(By.Name("hl_AuthPay")).Click();

        //    Delay(2);
        //    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table")).Click();



        //    //Authorisze payment  
        //    string Effective_Date = String.Empty, Bank = String.Empty, Branch = String.Empty, Account_Number = String.Empty, Name = String.Empty, Account_Type = String.Empty,
        //       Stop_Date = String.Empty, Cheque_Stale_Months = String.Empty, credit_Card = String.Empty, Expiry_date = String.Empty;

        //    OpenDBConnection("SELECT * FROM SSLP_Data");
        //    reader = command.ExecuteReader();
        //    while (reader.Read())
        //    {

        //        Effective_Date = reader["Effective_Date"].ToString().Trim();
        //        Bank = reader["Bank"].ToString().Trim();
        //        Branch = reader["Branch"].ToString().Trim();
        //        Account_Number = reader["Account_Number"].ToString().Trim();
        //        Name = reader["Name"].ToString().Trim();
        //        Account_Type = reader["Account_Type"].ToString().Trim();
        //        Stop_Date = reader["Date_of_incident"].ToString().Trim();
        //        Cheque_Stale_Months = reader["Cheque_Stale_Months"].ToString().Trim();
        //        credit_Card = reader["credit_Card"].ToString().Trim();
        //        Expiry_date = reader["Expiry_date"].ToString().Trim();


        //    }
        //    connection.Close();

        //    //Bank account if details


        //    //Click next

        //    Delay(2);
        //    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table")).Click();

        //    //Click Authorize
        //    Delay(2);
        //    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[16]/td/table/tbody/tr/td[2]/table")).Click();




        //    // Authorise Claim validation

        //    //Validate calim status
        //    string ClaimStatus = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/span/em")).Text;

        //    ClaimStatus.Contains("Payments Created");

        //    //Process Payment



        //    //Hover on claim options
        //    IWebElement ClaimOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
        //    //Creating object of an Actions class
        //    Actions action2 = new Actions(_driver);
        //    //Performing the mouse hover action on the target element.
        //    action2.MoveToElement(ClaimOptionElement).Perform();
        //    Delay(5);

        //    //Click on process payment
        //    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[10]/td/div/div[3]/a/img")).Click();
        //    Delay(5);


        //    //Click on Confirm Payment textbox

        //    _driver.FindElement(By.Name("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td[5]/input")).Click();
        //    Delay(5);

        //    //Click on process payment button
        //    _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/center/table[2]/tbody/tr[1]/td[3]/table")).Click();
        //    Delay(5);




        //    //Validate claim status
        //    string ClaimpaymentStatus = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[3]/center/div/table/tbody/tr/td/span/table/tbody/tr[7]/td/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/span/em")).Text;

        //    ClaimpaymentStatus.Contains("Claim Payment Raised");


        //    //workflow valdation and checks 


        //}
        //catch (Exception ex)
        //{
        //    if (ex.Message.Length > 255)
        //    {
        //        errMsg = ex.Message.Substring(0, 255);
        //    }
        //    else
        //    {
        //        errMsg = ex.Message;
        //    }
        //    results = "Failed";
        //}
        //writeResultsToDB(results, Int32.Parse(scenarioID), errMsg);
        //Assert.IsTrue(results.Equals("Passed"));
        //}

        [OneTimeTearDown]
        public void closeBrowser()
        {
           EmailReport();
            DisconnectBrowser();
          
        }
    }
}