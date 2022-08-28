using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Hybrid_Framework
{
    [TestClass]
    public class UnitTest1
    {
        [DynamicData(nameof(ReadExcel), DynamicDataSourceType.Method)]
        //Getdata is a method that is defined by the user
        [TestMethod]
        public void DatadrivenusingExcelSheet(string Email, string Password)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "https://www.pggoodeveryday.com/login/";

            driver.FindElement(By.Id("login-email")).SendKeys(Email);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(50);
            driver.FindElement(By.Id("login-password")).SendKeys(Password);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(50);
            // driver.FindElement(By.Id("EnrollmentDate")).SendKeys(eDate);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(3000);
            driver.Quit();
        }
        public static IEnumerable<object[]> ReadExcel()
        {
            //creating the worksheet object
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo("test_data.xlsx")))
            {
                //worksheet object
                ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
                int rowcount = worksheet.Dimension.End.Row;//give number of rows in integer form
                for (int i = 2; i <= rowcount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                    };
                }
            }
        }
    }
}
