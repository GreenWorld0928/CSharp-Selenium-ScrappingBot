using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium.Support.UI;
using AutoIt;
using System.Xml.Linq;

namespace ScrappingBot
{
    public partial class Form1 : Form
    {
        ChromeDriver driver;
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(Screen.PrimaryScreen.Bounds.Width-237, Screen.PrimaryScreen.Bounds.Height-381);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            ChromeOptions options = new ChromeOptions();
            //options.AddArguments("--headless");
            //driver.Manage().Window.Maximize();
            driver = new ChromeDriver(service, options);
            // Set the base website
            //string baseURL = "file:///E:/New%20Text%20Document.html";
            string baseURL = "https://app.rbassetsolutions.com/ims/login.aspx";
            driver.Navigate().GoToUrl(baseURL);
            var loginForm = driver.FindElement(By.Id("login_form"));
            var inputFields = loginForm.FindElements(By.TagName("input"));
            inputFields[0].SendKeys("selimdinc");
            inputFields[1].SendKeys("6465dincka");
            inputFields[2].Click();


        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            driver.Close();
        }

        private async void startBtn_Click(object sender, EventArgs e)
        {
            try
            {
                driver.FindElement(By.Id("rbas_place_sell_ad")).Click();
                driver.FindElement(By.Id("ef_brand")).SendKeys("Volvo");
                driver.FindElement(By.Id("ef_model")).SendKeys("EC55C");
                driver.FindElement(By.Id("category-browse")).Click();

                var catalog = driver.FindElement(By.Id("ef_catalog"));
                var selectElement = new SelectElement(catalog);
                selectElement.SelectByText("Construction");

                await Task.Delay(1000);

                var maincategory = driver.FindElement(By.Id("ef_maincategory"));
                selectElement = new SelectElement(maincategory);
                selectElement.SelectByText("Excavators");

                await Task.Delay(1000);
                var subcategory = driver.FindElement(By.Id("ef_subcategory"));
                selectElement = new SelectElement(subcategory);
                selectElement.SelectByText("Mini excavators < 7t (Mini diggers)");

                driver.FindElement(By.Id("form-continue-button")).Click();

                var yearofmanufacture = driver.FindElement(By.Id("ef_yearofmanufacture"));
                selectElement = new SelectElement(yearofmanufacture);
                selectElement.SelectByText("2009");

                driver.FindElement(By.Id("ef_meterreadouthours")).SendKeys("10");
                               

                driver.FindElement(By.Id("ef_otherinformation")).SendKeys("WorldWide Shipping");
                driver.FindElement(By.Id("otherinformation_addbutton")).Click();

                //driver.FindElements(By.ClassName("image_uploader_button"))[0].Click();
                string filePath = @"C:\Users\UserOSA\Documents\info\Screenshot_3.png";
                string allPath = filePath + " \n " + filePath + " \n " + filePath + " \n " + filePath;

                driver.FindElement(By.XPath("//input[@type='file']")).SendKeys(allPath);
                await Task.Delay(3000);

                driver.FindElement(By.Id("ef_priceoriginal")).SendKeys("1000");
                driver.FindElement(By.Id("ef_includes_vat_0")).Click();

                driver.FindElement(By.Id("form-continue-button")).Click();

            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            } 
        }

        private void loadBtn_Click(object sender, EventArgs e)
        {


}
    }
}
