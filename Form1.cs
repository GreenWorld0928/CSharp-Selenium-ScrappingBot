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

namespace ScrappingBot
{
    public partial class Form1 : Form
    {
        ChromeDriver driver;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            driver = new ChromeDriver(service);
            // Set the base website
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
            driver.FindElement(By.Id("rbas_place_sell_ad")).Click();
            driver.FindElement(By.Id("ef_brand")).SendKeys("Volvo");
            driver.FindElement(By.Id("ef_model")).SendKeys("EC55C");
            driver.FindElement(By.Id("category-browse")).Click();

            var catalog = driver.FindElement(By.Id("ef_catalog"));
            var selectElement = new SelectElement(catalog);
            selectElement.SelectByText("Construction");

            await Task.Delay(3000);

            var maincategory = driver.FindElement(By.Id("ef_maincategory"));
            selectElement = new SelectElement(maincategory);
            selectElement.SelectByText("Excavators");

            await Task.Delay(3000);
            var subcategory = driver.FindElement(By.Id("ef_subcategory"));
            selectElement = new SelectElement(subcategory);
            selectElement.SelectByText("Mini excavators < 7t (Mini diggers)");

            driver.FindElement(By.Id("form-continue-button")).Click();
        }

        private void loadBtn_Click(object sender, EventArgs e)
        {

        }
    }
}
