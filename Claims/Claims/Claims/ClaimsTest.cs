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
                //SetproductName();
                string Claimant = String.Empty, Cause_of_incident = String.Empty, BI_Number = String.Empty, Roleplayer = String.Empty, SubClaimType = String.Empty, ClaimType = String.Empty, 
                    IdNum = String.Empty, Date_of_incident = String.Empty, Contact_Date = String.Empty, Email_Address = String.Empty, Mobile_Number = String.Empty, ClaimDescription = String.Empty, Gender = String.Empty, Title = String.Empty; 

                OpenDBConnection("SELECT * FROM SSLP_Data");
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    ClaimType = reader["ClaimType"].ToString().Trim();
                    Claimant = reader["Claimant"].ToString().Trim();
                    Cause_of_incident = reader["Cause_of_incident"].ToString().Trim();
                    BI_Number = reader["BI_Number"].ToString().Trim();
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
                
                Delay(2);
                //click on add Digital  Stack
                _driver.FindElement(By.Name("sv_Main")).Click();
                Delay(2);

                //click on Call centre
                _driver.FindElement(By.Name("cc_Main")).Click();
                Delay(2);

                //click on add Add call log  
                _driver.FindElement(By.Name("cc_Proc_cctcl")).Click();
                Delay(2);
                //click on  Call template 
                _driver.FindElement(By.Name("cc_Template")).Click();
                Delay(2);
                //click on MIP Sanlam
                _driver.FindElement(By.Name("cc_MIP")).Click();
                Delay(2);
                //click on   Claims
                _driver.FindElement(By.Name("cc_MIP_Claims")).Click();
                Delay(2);
                //click on   claim
                _driver.FindElement(By.Name("cc_MIP_Claims_AC_NewClaim")).Click();
                Delay(2);
                //click on complete
                _driver.FindElement(By.Name("btnComplete")).Click();
                Delay(2);



                //ClaimType
                SelectElement dropDown = new SelectElement(_driver.FindElement(By.Name("refActivityLogRefsMainReqClaimType")));
                dropDown.SelectByText(ClaimType);
                Delay(5);

                //click on Yes BI Number
                _driver.FindElement(By.XPath("//*[@id='refActivityLogRefsMainHaveBINumber1']")).Click();
                Delay(1);

                //click on YES Death Certificate
                _driver.FindElement(By.XPath("//*[@id='refActivityLogRefsMainHaveDeathCertificate1']")).Click();
                Delay(2);


                //click on YES ID Document
                _driver.FindElement(By.XPath("//*[@id='refActivityLogRefsMainHaveIDDocument1']")).Click();
                Delay(2);


                //click on Contract selection

                _driver.FindElement(By.Name("remlActivityLogRemsMaincbmct1")).Click();
                Delay(2);


                //Claims  

                String test_url_2_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";


                Assert.AreEqual(2, _driver.WindowHandles.Count);
                var newWindowHandle = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle));
                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                string expectedNewWindowTitle = test_url_2_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle).Title, expectedNewWindowTitle);
                Delay(1);
                _driver.FindElement(By.Name("frmContractReference")).SendKeys(contractRef);
                Delay(2);
                _driver.FindElement(By.Name("frmContractReference")).Click();

                Delay(2);
                _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[2]")).Click();//*[@id="lkpResultsTable"]/tbody/tr[2]
                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);
                Delay(5);

                //Click on Complete
                _driver.FindElement(By.XPath("//*[@id='stateimg-5']")).Click();

                //Click on Related Entities
                Delay(2);
                _driver.FindElement(By.XPath("//*[@id='stateimg-6']")).Click();


                try
                {
                    //Click on Add new
                    Delay(2);
                    _driver.FindElement(By.Name("fcRemLabel1")).Click();
                }
                catch
                {

                    //Click on Related Entities
                    Delay(2);
                    _driver.FindElement(By.XPath("//*[@id='stateimg-6']")).Click();
                }
                //*[@id="stateimg-6"]

                //Click on Add new
                Delay(2);
                _driver.FindElement(By.Name("fcRemLabel1")).Click();

                //Mutimediad pop
                String test_url_3_title = "SANLAM RM - Safrican Retail";


                Assert.AreEqual(2, _driver.WindowHandles.Count);
                var newWindowHandle1 = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle1));
                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                string expectedNewWindowTitle1 = test_url_3_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle1).Title, expectedNewWindowTitle1);
               

               //add death certificate 
                
       
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/div/span[1]")).Click();
                //Click on 
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/div/span[2]")).Click();
                //Click on 
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/div/span[2]")).Click();
                //Click on 
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/ul/li[6]/div/span")).Click();
                //Click on 
                Delay(2);
                IWebElement element = _driver.FindElement(By.Name("ffFile"));
                element.SendKeys("C:\\Users\\G992107\\Downloads\\UPload file.pdf");


                Delay(4);
                _driver.FindElement(By.Name("btnSubmit")).Click();

                //click on add 
                Delay(4);
                _driver.FindElement(By.Name("btnAdd")).Click();

                //Add Decease  Idetification


                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/div/span[1]")).Click();
                //Click on 
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/div/span[2]")).Click();
                //Click on 
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/div/span[2]")).Click();

                //Click on Decease  Idetification
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/ul/li[5]/div/span")).Click();
                //Click on 
                Delay(2);
                IWebElement element2 = _driver.FindElement(By.Name("ffFile"));
                element2.SendKeys("C:\\Users\\G992107\\Downloads\\UPload file.pdf");
  

                Delay(4);
                _driver.FindElement(By.Name("btnSubmit")).Click();

                //click on add 
                Delay(4);  
                _driver.FindElement(By.Name("btnAdd")).Click();


                //Add BI-1663

                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/div/span[1]")).Click();

                //Click on 
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/div/span[2]")).Click();
                //Click on 
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/div/span[2]")).Click();

                //Click on Add BI-1663
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr/td/div/div/div/div/ul/li/ul/li[26]/ul/li[2]/ul/li[3]/div/span")).Click();
                //Click on 
                Delay(2);
                IWebElement element1 = _driver.FindElement(By.Name("ffFile"));
                element1.SendKeys("C:\\Users\\G992107\\Downloads\\UPload file.pdf");

                //cxlick on close
                Delay(3);
                _driver.FindElement(By.Name("btnSubmit")).Click();

                //cxlick on close
                Delay(3);
                _driver.FindElement(By.Name("btnClose")).Click();
                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);


                //click on complete
                Delay(2);
               _driver.FindElement(By.Name("btnComplete")).Click();
                //Click on Add new
                Delay(2);
                _driver.FindElement(By.Name("fcIDNumber")).SendKeys(IdNum);


                //Click on search
                Delay(2);
                _driver.FindElement(By.Name("fcPersonLkp")).Click();


                Delay(4);
                  
                //Mutimediad pop
                String test_url_4_title = "SANLAM RM - Safrican Retail - Warpspeed Lookup Window";

               
                Assert.AreEqual(2, _driver.WindowHandles.Count);
                var newWindowHandle3 = _driver.WindowHandles[1];
                Assert.IsTrue(!string.IsNullOrEmpty(newWindowHandle1));
                /* Assert.AreEqual(driver.SwitchTo().Window(newWindowHandle).Url, http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?); */
                string expectedNewWindowTitle4 = test_url_4_title;
                Assert.AreEqual(_driver.SwitchTo().Window(newWindowHandle3).Title, expectedNewWindowTitle4);

                var Firstname = _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[2]/td[2]")).Text;

                var Surname = _driver.FindElement(By.XPath("//*[@id='lkpResultsTable']/tbody/tr[2]/td[3]")).Text;

                //Click on Add new
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr/td/center[2]/table[2]/tbody/tr[4]/td[2]/span/center/table/tbody/tr[2]/td/center/table/tbody/tr[2]")).Click();
                
            
                /* Return to the window with handle = 0 */
                _driver.SwitchTo().Window(_driver.WindowHandles[0]);

                //incednt Date
                Delay(2);
                _driver.FindElement(By.Name("fcIncidentDate")).SendKeys(Date_of_incident);

                //incednt Date
                Delay(2);
                _driver.FindElement(By.Name("fcFirstContact")).SendKeys(Date_of_incident);

                //click on add 
                Delay(4);
                //_driver.FindElement(By.Name("fcLifeAssured")).Click();

                //Life assured
                SelectElement dropDown2 = new SelectElement(_driver.FindElement(By.Name("fcLifeAssured")));
                dropDown2.SelectByText(Claimant);
                Delay(5);


             
                //ClaimType Person
                SelectElement dropDown1 = new SelectElement(_driver.FindElement(By.Name("fcClaimType")));
                dropDown1.SelectByText(ClaimDescription);
                Delay(5);

                //Cause of Incident
                SelectElement dropDown3= new SelectElement(_driver.FindElement(By.Name("fcIncidentCause")));
                dropDown3.SelectByText(Cause_of_incident);
                Delay(5);
                try

                {

                    //Firstname
                    Delay(2);
                    _driver.FindElement(By.Name("fcFirstName")).SendKeys(Firstname);

                    //Surname
                    Delay(2);
                    _driver.FindElement(By.Name("fcLastName")).SendKeys(Surname);

                    //Gender 
                    SelectElement dropDown4 = new SelectElement(_driver.FindElement(By.Name("fcGender")));
                    dropDown4.SelectByText(Gender);
                    Delay(5);

                    //Title
                    SelectElement dropDown5 = new SelectElement(_driver.FindElement(By.Name("fcTitle")));
                    dropDown5.SelectByText(Title);
                    Delay(5);
                }
                catch
                { 
                    



                
                }
                

                //BI-number 
                Delay(2);
                _driver.FindElement(By.Name("fcBINumber")).SendKeys(BI_Number);

                //Email_Address
                Delay(2);
                _driver.FindElement(By.Name("fcEmailAddress")).SendKeys(Email_Address);

                //Mobile_Number 
                Delay(2);
                _driver.FindElement(By.Name("fcMobileNumber")).SendKeys(Mobile_Number);

                clickOnMainMenu();

                //Click submit
                Delay(2);
                _driver.FindElement(By.Name("btnSubmit")).Click();


                
                Delay(3);


                //Validate calim status
                string actualvalue = _driver.FindElement(By.Id("fcClaimStatus1_span")).Text;

                actualvalue.Contains("Authorised Claim");

                _driver.FindElement(By.Name("fcClaimRef1")).Click();


                //validation of claim





                //Hover on claim options
                IWebElement policyOptionElement = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));
                //Creating object of an Actions class
                Actions action = new Actions(_driver);
                //Performing the mouse hover action on the target element.
                action.MoveToElement(policyOptionElement).Perform();
                Delay(5);
                //Click on authorise
                _driver.FindElement(By.XPath("/html/body/center/center/form[3]/table/tbody/tr[2]/td[3]/table[1]/tbody/tr[4]/td[2]/span/table/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/div[2]/table/tbody/tr[7]/td/div/div[3]/a/img")).Click();
                Delay(5);


                //click on contrct summary
                Delay(2);
                _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div[7]/table[5]/tbody/tr/td/table/tbody/tr/td[3]/a")).Click();





                // Authorise Claim


                //Authorisze payment  
                string Effective_Date = String.Empty, Bank = String.Empty, Branch = String.Empty, Account_Number = String.Empty, Name = String.Empty, Account_Type = String.Empty,
                   Stop_Date = String.Empty, Cheque_Stale_Months = String.Empty, credit_Card = String.Empty, Expiry_date = String.Empty;

                OpenDBConnection("SELECT * FROM SSLP_Data");
                reader = command.ExecuteReader();
                while (reader.Read())
                {

                    Effective_Date = reader["Effective_Date"].ToString().Trim();
                    Bank = reader["Bank"].ToString().Trim();
                    Branch = reader["Branch"].ToString().Trim();
                    Account_Number = reader["Account_Number"].ToString().Trim();
                    Name = reader["Name"].ToString().Trim();
                    Account_Type = reader["Account_Type"].ToString().Trim();
                    Stop_Date = reader["Date_of_incident"].ToString().Trim();
                    Cheque_Stale_Months = reader["Cheque_Stale_Months"].ToString().Trim();
                    credit_Card = reader["credit_Card"].ToString().Trim();
                    Expiry_date = reader["Expiry_date"].ToString().Trim();
                  
                }
                connection.Close();

                //Bank account if details


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