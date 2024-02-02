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
using System.Xml.Linq;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.IO;

namespace ScrappingBot
{
  public partial class Form1 : Form
    {
      ChromeDriver driver;
        Thread trd = null;
        bool isLoaded = false;
        bool isStarted = false;
        string[] dataEntry;
        string[] fieldValue;
        int entryNum = 0;
        string directoryPath="";
        private void ThreadTask()
        {
            try
            {
                entryNum = 0;
                for (int i = 0; i < dataEntry.Length - 1; i++)
                {
                    entryNum++;
                    fieldValue = dataEntry[i + 1].Split(',');
                    this.statusLabel.Text = "         Entry No"+entryNum+" adding..." ;

                    driver.FindElement(By.Id("rbas_place_sell_ad")).Click();
                    driver.FindElement(By.Id("ef_brand")).SendKeys(fieldValue[1]);
                    driver.FindElement(By.Id("ef_model")).SendKeys(fieldValue[2]);
                    driver.FindElement(By.Id("category-browse")).Click();

                    var catalog = driver.FindElement(By.Id("ef_catalog"));
                    var selectElement = new SelectElement(catalog);
                    selectElement.SelectByText(fieldValue[3]);

                    Thread.Sleep(1000);

                    var maincategory = driver.FindElement(By.Id("ef_maincategory"));
                    selectElement = new SelectElement(maincategory);
                    selectElement.SelectByText(fieldValue[4]);

                    Thread.Sleep(1000);
                    var subcategory = driver.FindElement(By.Id("ef_subcategory"));
                    selectElement = new SelectElement(subcategory);
                    selectElement.SelectByText(fieldValue[5]);
                    driver.FindElement(By.Id("form-continue-button")).Click();

                    var yearofmanufacture = driver.FindElement(By.Id("ef_yearofmanufacture"));
                    selectElement = new SelectElement(yearofmanufacture);
                    selectElement.SelectByText(fieldValue[6]);

                    driver.FindElement(By.Id("ef_meterreadouthours")).SendKeys(fieldValue[7]);


                    driver.FindElement(By.Id("ef_otherinformation")).SendKeys(fieldValue[8]);
                    driver.FindElement(By.Id("otherinformation_addbutton")).Click();

                    //driver.FindElements(By.ClassName("image_uploader_button"))[0].Click();
                                                         
                    string[] fileNames = fieldValue[9].Split(' ');
                    string fullFileNames = directoryPath + string.Join(" \n " + directoryPath, fileNames);

                    driver.FindElement(By.XPath("//input[@type='file']")).SendKeys(fullFileNames);
                    Thread.Sleep(3000);

                    driver.FindElement(By.Id("ef_priceoriginal")).SendKeys(fieldValue[10]);
                    driver.FindElement(By.Id("ef_includes_vat_0")).Click();

                    driver.FindElement(By.Id("form-continue-button")).Click();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Entry No" + entryNum + " adding was stopped");
                entryNum--;
            }            
            this.statusLabel.Text = "      " + entryNum + " Data Entry(s) added";
            isStarted = false;
            trd.Abort();
            trd = null;

        }
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
            if (trd != null)
            {
                trd.Abort();
                trd = null;
            }
            driver.Close();
        }

        private void startBtn_Click(object sender, EventArgs e)
        {
           if(isLoaded) {
                trd = new Thread(new ThreadStart(this.ThreadTask));
                trd.IsBackground = true;
                trd.Start();
                loadBtn.Enabled = false;
                isStarted = true;
            }
            else
            {
                MessageBox.Show("DataEntry not loaded.");
            }
        }

        private void loadBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            { 
                string str = File.ReadAllText(openFileDialog.FileName);
                string[] tempPath = openFileDialog.FileName.ToString().Split('\\');
                tempPath[tempPath.Length - 1] = "Images\\";
                directoryPath = string.Join("\\", tempPath);
                //MessageBox.Show(directoryPath);
                dataEntry = str.Split('\n');
                isLoaded = true;
                statusLabel.Text = "  " + (dataEntry.Length-1) + " Data Entry(s) was loaded";
            }
        }

        private void stopBtn_Click(object sender, EventArgs e)
        {
            if (isStarted)
            {
                if (trd != null) { 
                    trd.Abort();
                    trd = null;
                }
                isStarted = false;
            }
        }
    }
}
