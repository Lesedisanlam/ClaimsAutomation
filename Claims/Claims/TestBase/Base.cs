using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data;
using System.Data.SqlClient;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;
using System.Net.Mail;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Data.OleDb;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System;

namespace ILR_TestSuite

{
    public class Base


    {

        private ChromeOptions _chromeOptions;

        public IWebDriver _driver, _webDriver;


        private string _userName;

        private string _password;

        public string _screenShotFolder, _screenShotFolderOutput;

        public string sheet = "Policy-Servicing";
        public static int currentMethod { get; set; }

        public static string connectionxls = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:/Users/e697642/Documents/GitHub/new/new/ilrsafricanautopolicyservicing/Policy Servicing/TestResults.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES""";

        public static string excelReportFilePath = $@"{Directory.GetCurrentDirectory()}\ClaimsTestResults.csv";
        // Check if file already exists. If yes, delete it.
        //public static string Excelconnection = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fileName};Extended Properties=""Excel 12.0 Xml;HDR=YES""";
        public static string provider =
                       "Microsoft.Jet.OLEDB.4.0";
        // "Microsoft.ACE.OLEDB.12.0";
        public static string excelConnectionString = $"Provider={provider};Data Source={excelReportFilePath};Extended Properties=\"text;HDR=Yes;FMT=Delimited\"";
        //public static string excelConnectionString = $"Provider={provider};Data Source={excelReportFilePath};Extended Properties='Excel 12.0 Xml;HDR=Yes';";

        public static string connectionString = @"Data Source='SRV007232, 1455';Initial Catalog=Automation;Integrated Security=True";
        public static SqlConnection connection = new SqlConnection(connectionString);
        public static SqlCommand command { get; set; }
        public string sqlString { get; set; }
        public static SqlDataReader reader { get; set; }



        [OneTimeSetUp]
        public void StartBrowser()
        {

            //  _driver = new ChromeDriver("C:/Code/bin");

            new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser);

            _chromeOptions = new ChromeOptions();
            _chromeOptions.AddArguments("--incognito");
            _chromeOptions.AddArguments("--ignore-certificate-errors");
            _driver = new ChromeDriver();


            //InternetExplorerOptions options = new InternetExplorerOptions();
            //options.IgnoreZoomLevel = true;
            //options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            //_driver = new InternetExplorerDriver(options);

            _screenShotFolder = $@"{AppDomain.CurrentDomain.BaseDirectory}Failed_ScreenShots​{ScreenShotDailyFolderName()}​\";

            _screenShotFolderOutput = $@"{AppDomain.CurrentDomain.BaseDirectory}Failed_ScreenShots​{ScreenShotDailyFolderName()}​.zip";
            if (Directory.Exists(_screenShotFolder))
            {
                Directory.Delete(_screenShotFolder, true);
            }
            new DirectoryInfo(_screenShotFolder).Create();






            //// Create FileStream for output ZIP archive
            //using (FileStream zipFile = File.Open("archive.zip", FileMode.Create))
            //{
            //    // Create archive
            //    using (va archive = new Archive())
            //    {
            //        // Add file to ZIP archive
            //        archive.CreateEntry("data.bin", "file.dat");
            //        archive.Save(zipFile);
            //    }
            //}


        }
        public void createExclReportFile()
        {

            if (File.Exists(excelReportFilePath))
            {
                File.Delete(excelReportFilePath);
            }
            using (StreamWriter fs = File.CreateText(excelReportFilePath))
            {

                String separator = ",";
                StringBuilder output = new StringBuilder();
                String headings = "Policy_Number, ExpectedResults, Test_Results, Comments, Test_Date, Product_name, Created_at ";
                //  output.AppendLine(string.Join(separator, headings));
                fs.WriteLine(headings);

            }

            //using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            //{

            //    connection.Open();

            //    execOleDbCommand(connection, $"create table PS_Results (Policy_Number varchar,ExpectedResults varchar ,Test_Results varchar ,Comments varchar,Test_Date varchar," +
            //        $"Function_name varchar,ProductName varchar,Created_at varchar);");
            //    connection.Close();
            //}
        }

        static void execOleDbCommand(OleDbConnection conn, string sqlText)
        {
            using (OleDbCommand command = new OleDbCommand())
            {

                command.Connection = conn;
                command.CommandText = sqlText;
                command.ExecuteNonQuery();

            }
        }
        public static void OpenDBConnection(string sqlCmnd)
        {
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }
            command = new SqlCommand(sqlCmnd, connection);
        }

        public string SetproductName()
        {
            var product = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv5']/table/tbody/tr[1]/td[2]")).Text;

            try
            {
                var cmd = $"UPDATE Claims_Scenarios SET productName = @product WHERE FunctionID = ID ";
                OpenDBConnection(cmd);
                command.Parameters.AddWithValue("@product", product);
                command.ExecuteNonQuery();




                return product;



            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                connection.Close();
            }




        }

        public static int getFuctionID(string funcName)
        {
            int id = 0;
            try
            {
                OpenDBConnection("SELECT ID FROM ClaimsFunction WHERE function_name = '" + funcName + "'");
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    id = (int)reader["ID"];

                }

                connection.Close();
            }

            catch (Exception ex)
            {
                Console.WriteLine("Exception:" + ex.ToString());

            }
            return id;
        }
        public static string getFuncName(int id)
        {
            string funcName;
            funcName = String.Empty;

            try
            {
                OpenDBConnection("SELECT function_name FROM ClaimsFunction WHERE ID =" + id);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    funcName = reader["function_name"].ToString();

                }
                connection.Close();

            }

            catch (Exception ex)
            {
                Console.WriteLine("Exception:" + ex.ToString());

            }


            return funcName;
        }

        private static string ScreenShotDailyFolderName()

        {

            return DateTime.Now.ToString("yyyyMMdd").Replace("AM", string.Empty).Replace("PM", string.Empty);

        }

        public void TakeScreenshot(string fileName)
        {
            var filePath = $@"{_screenShotFolder}\Failed_Scenarios\";

            if (!Directory.Exists(filePath))

                new DirectoryInfo(filePath).Create();


            ITakesScreenshot ssdriver = _driver as ITakesScreenshot;

            Screenshot screenshot = ssdriver.GetScreenshot();

            fileName = $"{fileName}{ScreenShotTime()}.png";

            screenshot.SaveAsFile($"{filePath}{fileName}", ScreenshotImageFormat.Png);

        }

        private static string ScreenShotTime()

        {

            return DateTime.Now.TimeOfDay.ToString().Replace(":", "_").Replace(".", "_");

        }

        public IWebDriver SiteConnection()
        {


            _driver.Url = "http://ilr-tst.safrican.co.za/";

            // IAlert alert = _driver.SwitchTo().Alert();
            //  alert.Dismiss();

            _userName = "G992092";//TODO add your user name and password


            _password = "G992092saftst";

            _driver.Manage().Window.Maximize();

            System.Threading.Thread.Sleep(2000);


            System.Threading.Thread.Sleep(2000);

            IWebElement loginTextBox = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[2]/td[2]/input"));
            IWebElement passwordTextBox = _driver.FindElement(By.XPath("/html/body/center/center/form[2]/div/table/tbody/tr/td/span/table/tbody/tr/td/div/center/div/table/tbody/tr[4]/td[2]/span/table/tbody/tr[3]/td[2]/input"));
            IWebElement loginBtn = _driver.FindElement(By.CssSelector("#GBLbl-1 > span > a"));
            loginTextBox.SendKeys(_userName);
            System.Threading.Thread.Sleep(6000);
            passwordTextBox.SendKeys(_password);

            System.Threading.Thread.Sleep(4000);
            //Check if password field is empty
            String textInsideInputBox = passwordTextBox.GetAttribute("value");
            if (String.IsNullOrEmpty(textInsideInputBox))
            {
                passwordTextBox.SendKeys(_password);
                System.Threading.Thread.Sleep(4000);
                loginBtn.Click();
                System.Threading.Thread.Sleep(2000);
                return _driver;
            }
            else
            {
                loginBtn.Click();
                System.Threading.Thread.Sleep(2000);
                return _driver;
            }
        }

        public Decimal getPremuimFromRateTable(string idNumber, string rolePlayer, string sumAsured, string product)
        {
            var premium = String.Empty;
            var age = String.Empty;
            //Calculate age based on IdNo
            var thisYear = DateTime.Now.Year.ToString().Substring(2);
            thisYear = DateTime.Now.Year.ToString();
            var id_year = Int32.Parse(idNumber.Substring(0, 2));
            if (id_year >= 00 && id_year <= Int32.Parse(DateTime.Now.Year.ToString().Substring(2)))
            {
                age = (DateTime.Now.Year - Int32.Parse("200" + id_year)).ToString();
            }
            else
            {
                age = (DateTime.Now.Year - Int32.Parse("19" + id_year)).ToString();
            }
            //Get product name for the rate table
            switch (product.Trim())
            {
                case "Safrican Serenity Funeral Premium (1000)":
                    product = "Serenity_Premium";
                    break;
                case "Safrican Serenity Funeral (2000)":
                    product = "Safrican_Serenity_Funeral";
                    break;
                case "Safrican Just Funeral (3000)":
                    product = "Safrican_Just_Funeral";
                    break;
                default:
                    break;
            }
            //Get roleplayer ref for DB table
            if ((rolePlayer.Trim()).Contains("Parent"))
            {
                rolePlayer = "Parent";
            }
            else if ((rolePlayer.Trim()).Contains("Child"))
            {
                rolePlayer = "Children";
            }
            else if ((rolePlayer.Trim()).Contains("Spouse"))
            {
                rolePlayer = "Spouse";
            }
            else if ((rolePlayer.Trim()).Contains("Wider"))
            {
                rolePlayer = "Extended";
            }
            else
            {
                rolePlayer = "ML";
            }
            var cover = rolePlayer + "_" + sumAsured;
            OpenDBConnection($"SELECT {cover} FROM {product} WHERE AGE = " + age);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                premium = reader[cover].ToString();
            }
            connection.Close();
            return Convert.ToDecimal(premium);
        }

        public void EmailReport()
        {
            try
            {

                string funcName, PolicyNo, ExpectedResults, TestResults, Comments, Test_Date, FunctionID, UserID, product_name, Created_at; int ID;
                funcName = String.Empty; PolicyNo = String.Empty; ExpectedResults = String.Empty; TestResults = String.Empty; Comments = String.Empty; Test_Date = String.Empty;
                product_name = String.Empty; UserID = String.Empty; FunctionID = String.Empty; Created_at = String.Empty;

                List<TestResultObject> results = new List<TestResultObject>();
                Delay(5);
                try
                {
                    OpenDBConnection("SELECT * FROM Claims_Scenarios WHERE  Created_at > DATEADD(HOUR, -12, GETDATE());");
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {

                        PolicyNo = reader["PolicyNo"].ToString();
                        ExpectedResults = reader["ExpectedResults"].ToString();
                        TestResults = reader["Test_Results"].ToString();
                        Comments = reader["Comments"].ToString();
                        Test_Date = DateTime.Now.ToString();
                        FunctionID = reader["FunctionID"].ToString();
                        UserID = reader["UserID"].ToString();
                        product_name = reader["ProductName"].ToString();
                        Created_at = reader["Created_at"].ToString();

                        if (Comments.Length > 50)
                        {

                            Comments = Comments.Substring(0, 50);
                        }

                        TestResultObject tstResult = new TestResultObject(PolicyNo, ExpectedResults, TestResults, Comments, Test_Date, FunctionID,
                            UserID, product_name, Created_at);
                        results.Add(tstResult);

                    }
                    connection.Close();
                    StringBuilder strBldr = new StringBuilder();

                    using (StreamWriter file = new StreamWriter(excelReportFilePath, true))
                    {

                        foreach (var item in results)
                        {

                            //  using (OleDbConnection conn = new OleDbConnection(excelConnectionString))

                            var line = (item.policyNo, item.expectedResults, item.testResults, item.comment, item.test_Date, product_name,
                                   item.created_at);
                            file.WriteLine(line);
                            //// conn.Open();

                            //  execOleDbCommand(conn, $"Insert into [PS_Results$] (Policy_Number,ExpectedResults,Test_Results ,Comments ,Test_Date ,Function_name,ProductName ,Created_at ) " +
                            //          $"values('{item.policyNo}','{item.expectedResults}'," +
                            //          $"'{item.testResults}','{item.comment}','{item.test_Date}'," +
                            //          $"'{getFuncName(Int32.Parse(item.functionID))}','{item.product_name}','{item.created_at}')");
                            // // conn.Close();


                        }
                    }
                }
                catch (Exception ex)
                {

                    Console.WriteLine("Exception:" + ex.ToString());
                }
                //    MailMessage mail = new MailMessage();
                //SmtpClient smtp = new SmtpClient("mail.sanlam.co.za");
                //    mail.Subject = "Sanlam Report";
                //    mail.Body = "Testing";

                //mail.From = new MailAddress(mail.Subject, mail.Body);
                //smtp.Port = 25;
                //smtp.Credentials = new System.Net.NetworkCredential("AutoResult@sanlamsky.co.za","");
                //smtp.EnableSsl = true;
                //smtp.Send(mail);

                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment("");

                //if (File.Exists(_screenShotFolderOutput))
                //{
                //    File.Delete(_screenShotFolderOutput);
                //}

                //using (var archive = ZipFile.Open(_screenShotFolderOutput, ZipArchiveMode.Create))
                //{
                //    archive.CreateEntryFromFile(_screenShotFolder, Path.GetFileName(_screenShotFolder));

                //}


                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("mail.sanlam.co.za");
                mail.From = new MailAddress("Autoresult@sanlamsky.co.za");
                mail.To.Add("napeb@sanlamsky.co.za");
                //,kamogelomo@sanlamsky.co.za,lesedima@sanlamsky.co.za,linda.zondi@sanlamsky.co.za,Shaquille.Bandura@sanlamsky.co.za
                //,kamogelomo@sanlamsky.co.za,lesedima@sanlamsky.co.za
                mail.Subject = "Claims Auto Test Results";
                mail.Body = @"Please see the attached Policy Servicing Automation Test Results.

        Kind Regards";


                //var report = new byte[300];
                //using (MemoryStream memoryStream = new MemoryStream())
                //{
                //    using (ZipArchive zipArchive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                //    {
                //        ZipArchiveEntry zipArchiveEntry = zipArchive.CreateEntry(_screenShotFolder);
                //        using (StreamWriter streamWriter = new StreamWriter(zipArchiveEntry.Open()))
                //        {
                //            streamWriter.Write(Encoding.Default.GetString(report));
                //        }
                //    }

                //     MemoryStream attachmentStream = new MemoryStream(memoryStream.ToArray());

                //    Attachment Zipattachment = new Attachment(attachmentStream, _screenShotFolder + ".zip", MediaTypeNames.Application.Zip);


                //  Attachment Zipattachment = new Attachment(_screenShotFolder);

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(excelReportFilePath);
                mail.Attachments.Add(attachment);
                //   mail.Attachments.Add(Zipattachment);
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("Autoresult@sanlamsky.co.za", "P@ssword987951");
                SmtpServer.EnableSsl = true;

                ServicePointManager.ServerCertificateValidationCallback =
                  delegate (
                  object s,
                  X509Certificate certificate,
                  X509Chain chain,
                  SslPolicyErrors sslPolicyErrors
       )
                  {
                      return true;
                  };
                SmtpServer.Send(mail);
                attachment.Dispose();

                // }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception:" + ex.ToString());


            }

            if (File.Exists(excelReportFilePath))
            {
                File.Delete(excelReportFilePath);
            }



        }



        public void writeResultsToDB(string results, int scenario_id, string comments)
        {
            OpenDBConnection($"UPDATE Claims_Scenarios SET Test_Results = @results, Test_Date =@testDate, Comments = @comments WHERE ID = {scenario_id}");
            var testDate = DateTime.Now.ToString();
            command.Parameters.AddWithValue("@results", results);
            command.Parameters.AddWithValue("@testDate", testDate);
            command.Parameters.AddWithValue("@comments", comments);

            command.ExecuteNonQuery();
        }



        public void DisconnectBrowser()

        {

            if (_driver != null)

                _driver.Quit();

        }



        public string GetSeleniumFormatTag(string inputControlName)

        {

            var result = $"//*[@id=\"{inputControlName}\"]";

            return result;

        }



        public void Delay(int delaySeconds)

        {

            Thread.Sleep(delaySeconds * 1000);

        }
        public static IEnumerable<string[]> GetTestData(string methodName)
        {
            var conractRef = String.Empty;
            var scenarioID = String.Empty;
            int id = getFuctionID(methodName);
            OpenDBConnection($"SELECT PolicyNo,ID FROM Claims_Scenarios WHERE FunctionID = {id}");
            reader = command.ExecuteReader(); while (reader.Read())
            {
                scenarioID = reader["ID"].ToString();
                conractRef = reader["PolicyNo"].ToString();
                break;
            }
            connection.Close();
            yield return new[] { conractRef, scenarioID };
        }


    }


    [TestFixture]


    public class TestResultObject
    {
        public string policyNo;
        public string expectedResults;
        public string testResults;
        public string comment;
        public string test_Date;
        public string functionID;
        public string userID;
        public string product_name;
        public string created_at;

        public TestResultObject(string policyNo, string expectedResults, string testResults, string comment, string test_Date,
            string functionID, string userID, string product_name, string created_at)
        {
            this.policyNo = policyNo;
            this.expectedResults = expectedResults;
            this.testResults = testResults;
            this.comment = comment;
            this.test_Date = test_Date;
            this.functionID = functionID;
            this.userID = userID;
            this.product_name = product_name;
            this.created_at = created_at;
        }

    }

}