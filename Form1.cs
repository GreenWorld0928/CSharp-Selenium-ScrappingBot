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


            //driver.FindElement(By.Id("a3a0d7182")).SendKeys("selimdinc");
            //driver.FindElement(By.Id("a1e74233e")).SendKeys("6465dincka");
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            driver.Close();
        }
    }
}
