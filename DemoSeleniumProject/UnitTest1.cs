using DemoSeleniumProject.Data;
using DemoSeleniumProject.PageObjectModel;
using DemoSeleniumProject.SourceMain;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal.Execution;
using OpenQA.Selenium.Chrome;
using PageFactoryCore;
using System.Configuration;
using System.Data.Common;
using IronXL;
using System;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace DemoSeleniumProject
{
    public class Tests
    {
      
        //Settings1 set = new Settings1();
        [SetUp]
        public void Setup()
        {
            BaseTest.driver = new ChromeDriver();
            //login to application
            LoginPage loginPage = new LoginPage(BaseTest.driver);
            loginPage.loginToApplicationAndVerify();
        }

        [Test, Order(1),Category("PositiveTest")]
        public void fillSampleFormAndVerify()
        {      
            HomePage homePage = new HomePage(BaseTest.driver);
            homePage.fillSampleFormAndVerify();
        }

        [Test, Order(2), Category("PositiveTest")]
        public void performActionsOnLayOut()
        {
            HomePage homePage = new HomePage(BaseTest.driver);
            homePage.performActionsOnLayOut();
        }

        [Test, Order(3), Category("NegativeTest")]
        public void performActionsOnLayOutNegative()
        {
            HomePage homePage = new HomePage(BaseTest.driver);
            homePage.performActionsOnLayOutNegative();
        }

        [Test, Order(4), Category("PositiveTest")]
        public void verifyTableLayout()
        {
            //home page object declaration
            HomePage homePage = new HomePage(BaseTest.driver);
            homePage.verifyTableLayout();

        }

        [TearDown]
        public void OneTimeTearDown()
        {
            if (TestContext.CurrentContext.Result.Outcome != ResultState.Success)
            {
                CommonMethod.captureScreenShot();
            }
            BaseTest.driver.Quit();
        }
    }
}