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
using System.Collections;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;
using System.Reflection.Emit;

namespace ScrappingBot
{
  public partial class Form1 : Form
    {
      ChromeDriver driver;
        Thread trd = null;
        Thread trdChingInbox = null;
        bool isLoaded = false;
        bool isStarted = false;
        bool isCheckingInbox = false;
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

                    try
                    {
                        driver.FindElement(By.Id("ef_itemtype_2")).Click();
                    }
                    catch (Exception)
                    {
                    }

                    driver.FindElement(By.Id("category-browse")).Click();

                    var catalog = driver.FindElement(By.Id("ef_catalog"));
                    var selectElement = new SelectElement(catalog);
                    int count;

                    count = 0;
                    while (true)
                    {
                        try
                        {
                            Thread.Sleep(500);
                            selectElement.SelectByText(fieldValue[3]);
                            break;
                        }
                        catch (Exception)
                        {
                            count++;
                            if (count < 6) continue;
                            else break;
                        }
                    }

                    count = 0;
                    while (true)
                    {
                        try {
                            Thread.Sleep(500);
                            var maincategory = driver.FindElement(By.Id("ef_maincategory"));
                            selectElement = new SelectElement(maincategory);
                            selectElement.SelectByText(fieldValue[4]);
                            break;
                        } catch (Exception)
                        {  
                            count++;
                            if (count < 6) continue;
                            else break;
                        }
                    }

                    count = 0;
                    while (true)
                    {
                        try
                        {
                            Thread.Sleep(500);
                            var subcategory = driver.FindElement(By.Id("ef_subcategory"));
                            selectElement = new SelectElement(subcategory);
                            selectElement.SelectByText(fieldValue[5]);
                            break;
                        }
                        catch (Exception)
                        {
                            count++;
                            if (count < 6) continue;
                            else break;
                        }
                    }

                    driver.FindElement(By.Id("form-continue-button")).Click();

                    var yearofmanufacture = driver.FindElement(By.Id("ef_yearofmanufacture"));
                    selectElement = new SelectElement(yearofmanufacture);
                    selectElement.SelectByText(fieldValue[6]);

                    count = 0;
                    while (true)
                    {
                        try
                        {
                            Thread.Sleep(500);
                            var tractorType = driver.FindElement(By.Id("ef_tractortype"));
                            selectElement = new SelectElement(tractorType);
                            selectElement.SelectByValue("1000100");
                            break;
                        }
                        catch (Exception)
                        {
                            count++;
                            if (count < 2) continue;
                            else break;
                        }
                    }

                    driver.FindElement(By.Id("ef_meterreadouthours")).SendKeys(fieldValue[7]);


                    driver.FindElement(By.Id("ef_otherinformation")).SendKeys(fieldValue[8]);
                    driver.FindElement(By.Id("otherinformation_addbutton")).Click();

                    var locationInfo = driver.FindElement(By.Id("ef_useraddressbookid"));
                    selectElement = new SelectElement(locationInfo);
                    selectElement.SelectByValue("ChooseRbYard");

                    count = 0;
                    while (true)
                    {
                        try
                        {
                            Thread.Sleep(500);
                            var assetCustomYardLocation = driver.FindElement(By.Id("ef_assetCustomYardLocation"));
                            selectElement = new SelectElement(assetCustomYardLocation);
                            selectElement.SelectByValue("21");
                            driver.FindElement(By.Id("assetCustomUserLocationsSave")).Click();
                            break;
                        }
                        catch (Exception)
                        {
                            count++;
                            if (count < 6) continue;
                            else break;
                        }
                    }
                    //driver.FindElements(By.ClassName("image_uploader_button"))[0].Click();

                    string[] fileNames = fieldValue[9].Split('#');
                    string fullFileNames = directoryPath + string.Join(" \n " + directoryPath, fileNames);
                    try
                    {
                        driver.FindElement(By.XPath("//input[@type='file']")).SendKeys(fullFileNames);
                        Thread.Sleep(3000);
                    }catch(Exception) {
                        this.statusLabel.Text = "  Photos not loaded in Entry No " + entryNum;
                    }
                    driver.FindElement(By.Id("ef_priceoriginal")).SendKeys(fieldValue[10]);
                    driver.FindElement(By.Id("ef_includes_vat_0")).Click();

                    driver.FindElement(By.Id("form-continue-button")).Click();

                }
            }
            catch (Exception)
            {
                MessageBox.Show("Entry No" + entryNum + " adding was stopped");
                entryNum--;
            }            
            this.statusLabel.Text = "      " + entryNum + " Data Entry(s) added";
            isStarted = false;
            this.loadBtn.Enabled = true;            
            this.checkInboxBtn.Enabled = true;
            trd.Abort();
            trd = null;
        }

        private void ThreadCheckingInbox()
        {
            try
            {
                driver.FindElement(By.Id("noti-badge-3784")).Click();
                driver.FindElements(By.CssSelector("i[class='fas fa-angle-down']"))[0].Click();
                driver.FindElement(By.Id("divfilter_item_64")).Click();
                
                var totalMsg = driver.FindElements(By.CssSelector("span[class='total-count']"))[0].Text;
                
                int msgNum = Convert.ToInt32(totalMsg);
                int unReadMsgNum = 0;

                for (int i = 0; i < (msgNum - 1) / 30 + 1; i++)
                {
                    driver.FindElement(By.Id("selectAll")).Click();
                    try
                    {
                        var MsgElements = driver.FindElements(By.CssSelector("div[class='row form-inline bottomdashed p-3']"));
                        var contentElements = driver.FindElements(By.CssSelector("div[class='row form-inline bottomdashed grayed hidden']"));
                        for (int j = 0; j < MsgElements.Count; j++) {
                            try
                            {
                                var unReadElement = MsgElements[j].FindElements(By.CssSelector("i[class='fas fa-eye-slash']"))[0];
                                if (unReadElement != null) {
                                    MsgElements[j].Click();
                                    string[] productInfo = MsgElements[j].Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                    string[] contactInfo = contentElements[j].Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                    TextWriter tw = new StreamWriter("unReadMessages.csv", true);
                                    string name = contactInfo[0].Replace(",", "-");
                                    string str = name + "," + contactInfo[1] + "," + contactInfo[2] + "," + contactInfo[3] + "," + productInfo[4] + "," + productInfo[1];
                                    tw.WriteLine(str);
                                    tw.Close();
                                    unReadMsgNum++;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                    //Delete read messages
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    js.ExecuteScript("submitIfChecked('rowid', 1, 0, 'Please select at least #minimum# item(s).', 'Are you sure that you want to delete #number# selected item(s)?', 'deletemessage', '');");

                    try
                    {
                       // driver.SwitchTo().Alert().Accept();
                    }
                    catch (NoAlertPresentException ex)
                    {
                    }
                }
                this.msgStatusLabel.Text = unReadMsgNum + " unread message(s) saved";
            }
            catch (Exception)
            {
                this.msgStatusLabel.Text = "There are no new messages";
            }
            isStarted = false;
            this.startBtn.Enabled = true;
            trdChingInbox.Abort();
            trdChingInbox = null;
        }
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(Screen.PrimaryScreen.Bounds.Width-337, Screen.PrimaryScreen.Bounds.Height-481);
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
            //inputFields[0].SendKeys("christian@bks.dk");
            //inputFields[1].SendKeys("Cks00005159");
            //inputFields[2].Click();
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (trd != null)
            {
                trd.Abort();
                trd = null;
            }
            if (trdChingInbox != null)
            {
                trdChingInbox.Abort();
                trdChingInbox = null;
            }
            driver.Close();
        }
        private void startBtn_Click(object sender, EventArgs e)
        {
           if(isLoaded && !isStarted) {
                trd = new Thread(new ThreadStart(this.ThreadTask));
                trd.IsBackground = true;
                trd.Start();
                loadBtn.Enabled = false;
                checkInboxBtn.Enabled = false;
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
                    entryNum--;
                    statusLabel.Text = "      " + entryNum + " Data Entry(s) added";
                    loadBtn.Enabled = true;
                    checkInboxBtn.Enabled = true;
                }
                isStarted = false;
            }
        }
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        TimeSpan timeLeft;
        private void checkInboxBtn_Click(object sender, EventArgs e)
        {
            if (!isCheckingInbox)
            {
                isCheckingInbox = true;
                timer.Enabled = true;
                timer.Interval = 1000;
                timeLeft = TimeSpan.FromSeconds(1);
                timer.Tick += new System.EventHandler(timer_Tick);
                timer.Start();
            }
        }
        private void timer_Tick(object sender, EventArgs e)
        {
            timeLeft = timeLeft.Subtract(new TimeSpan(0, 0, 1));
            leftTimeLabel.Text = timeLeft.Minutes.ToString()+ "min " + timeLeft.Seconds.ToString()+"s " + "after rechecking inbox";
            if (timeLeft.Minutes == 0)
            {
                msgStatusLabel.Text = "Checking Inbox ... ... ...";
                trdChingInbox = new Thread(new ThreadStart(this.ThreadCheckingInbox));
                trdChingInbox.IsBackground = true;
                trdChingInbox.Start();
                startBtn.Enabled = false;
                timeLeft = TimeSpan.FromMinutes(30);
            }
        }
        private void stopCheckBtn_Click(object sender, EventArgs e)
        {
            if (isCheckingInbox)
            {
                timer.Stop();
                timer.Tick -= new System.EventHandler(timer_Tick);
                isCheckingInbox = false;
                leftTimeLabel.Text = "If checking, will run every 30 mins";
            }
        }
    }
}
