using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Drawing;
using OpenQA.Selenium.Remote;
using DeathByCaptcha;
using Exception = System.Exception;

namespace ClarkCountryCaseDownloader
{

    [Parallelizable]
    public class Locate
    {
        public string DownloadPath;
        private ChromeDriverEx _d;
        public static string logsPath;
        private bool destroyedDriver;
        public string UserDefinedPath;
        public string file_name_global = "";
        ITakesScreenshot screenshotDriver;
        bool take_screenshot = true;
        int sleeptime = 3000;

        public class ChromeOptionsWithPrefs : ChromeOptions
        {
            public Dictionary<string, object> prefs { get; set; }
        }

        public class ChromeDriverEx : ChromeDriver
        {
            private const string SendChromeCommandWithResult = "sendChromeCommandWithResponse";
            private const string SendChromeCommandWithResultUrlTemplate = "/session/{sessionId}/chromium/send_command_and_get_result";

            public ChromeDriverEx(ChromeDriverService service, ChromeOptions options)
                : base(service, options)
            {
                CommandInfo commandInfoToAdd = new CommandInfo(CommandInfo.PostCommand, SendChromeCommandWithResultUrlTemplate);
                this.CommandExecutor.CommandInfoRepository.TryAddCommand(SendChromeCommandWithResult, commandInfoToAdd);
            }

            public object ExecuteChromeCommandWithResult(string commandName, Dictionary<string, object> commandParameters)
            {
                if (commandName == null)
                {
                    throw new ArgumentNullException("commandName", "commandName must not be null");
                }

                Dictionary<string, object> parameters = new Dictionary<string, object>();
                parameters["cmd"] = commandName;
                parameters["params"] = commandParameters;
                Response response = this.Execute(SendChromeCommandWithResult, parameters);
                return response.Value;
            }

            private Dictionary<string, object> EvaluateDevToolsScript(string scriptToEvaluate)
            {
                // This code is predicated on knowing the structure of the returned
                // object as the result. In this case, we know that the object returned
                // has a "result" property which contains the actual value of the evaluated
                // script, and we expect the value of that "result" property to be an object
                // with a "value" property. Moreover, we are assuming the result will be
                // an "object" type (which translates to a C# Dictionary<string, object>).
                Dictionary<string, object> parameters = new Dictionary<string, object>();
                parameters["returnByValue"] = true;
                parameters["expression"] = scriptToEvaluate;
                object evaluateResultObject = this.ExecuteChromeCommandWithResult("Runtime.evaluate", parameters);
                Dictionary<string, object> evaluateResultDictionary = evaluateResultObject as Dictionary<string, object>;
                Dictionary<string, object> evaluateResult = evaluateResultDictionary["result"] as Dictionary<string, object>;

                // If we wanted to make this actually robust, we'd check the "type" property
                // of the result object before blindly casting to a dictionary.
                Dictionary<string, object> evaluateValue = evaluateResult["value"] as Dictionary<string, object>;
                return evaluateValue;
            }
        }

        public bool InitializeDriver()
        {
            bool res;
            try
            {
                ChromeDriverService defaultService = ChromeDriverService.CreateDefaultService("./");
                defaultService.HideCommandPromptWindow = true;
                ChromeOptionsWithPrefs options = new ChromeOptionsWithPrefs();
                options.BinaryLocation = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
                string val = "{\"version\":2,\"isGcpPromoDismissed\":false,\"selectedDestinationId\":\"Save as PDF\", \"recentDestinations\":[{\"id\":\"Save as PDF\", \"origin\":\"local\", \"account\":\"\"}]}";
                options.AddUserProfilePreference("printing.print_preview_sticky_settings.appState", val);
                options.AddUserProfilePreference("printing.default_destination_selection_rules", "{\"kind\":\"local\",\"namePattern\":\"Save as PDF\"}");
                options.AddArguments("--start-maximized");
                options.AddArgument("--kiosk-printing");
                options.AddArguments("--disable-gpu");
                options.AddArguments("--disable-extensions");
                options.AddArgument("headless");
                options.AddArgument("no-sandbox");
                options.AddArgument("test-type");
                options.AddArguments(new string[1]
                {
                    "disable-infobars"
                });
                try
                {
                    logsPath = Directory.GetCurrentDirectory() + "\\Logs\\";
                    if (Directory.Exists(logsPath) == false)
                    {
                        new DirectoryInfo(logsPath).Create();
                    }
                } catch (Exception e2)
                {
                    logsPath = Directory.GetCurrentDirectory() + "\\Logs Directory\\";
                    if (Directory.Exists(logsPath) == false)
                    {
                        new DirectoryInfo(logsPath).Create();
                    }
                }
                
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                try
                {
                    _d = new ChromeDriverEx(defaultService, options);
                } catch (Exception ex1)
                {
                    Console.WriteLine(ex1);
                    options.BinaryLocation = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe";
                    _d = new ChromeDriverEx(defaultService, options);
                }
                _d.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5.0);
                screenshotDriver = _d as ITakesScreenshot;
                res = true;
                destroyedDriver = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception in InitializeDriver", ex.Message);
                Thread.Sleep(sleeptime);
                res = false;
            }
            return res;
        }

        private void sleep(int time)
        {
            Thread.Sleep(time);
        }

        private void getScreenShot(string name)
        {
            if (this.take_screenshot)
            {
                string screenshotPath = Directory.GetCurrentDirectory() + "\\images\\";
                if (Directory.Exists(screenshotPath) == false)
                {
                    new DirectoryInfo(screenshotPath).Create();
                }
                Screenshot screenshot = screenshotDriver.GetScreenshot();
                screenshot.SaveAsFile(screenshotPath + name + ".png", ScreenshotImageFormat.Png);
            }
        }

        private string getCaptchaBase64()
        {
            string base_64 = "";
            try
            {
                string img_base_64 = (this._d as IJavaScriptExecutor).ExecuteScript(@"
                var canvas = document.createElement('canvas');
                var ctx = canvas.getContext('2d');

                function getMaxSize(srcWidth, srcHeight, maxWidth, maxHeight) {
                    var widthScale = null;
                    var heightScale = null;

                    if (maxWidth != null)
                    {
                        widthScale = maxWidth / srcWidth;
                    }
                    if (maxHeight != null)
                    {
                        heightScale = maxHeight / srcHeight;
                    }

                    var ratio = Math.min(widthScale || heightScale, heightScale || widthScale);
                    return {
                        width: Math.round(srcWidth * ratio),
                        height: Math.round(srcHeight * ratio)
                    };
                }
                function getBase64FromImage(img, width, height) {
                    var size = getMaxSize(width, height, 400, 400)
                    canvas.width = size.width;
                    canvas.height = size.height;
                    ctx.fillStyle = 'white';
                    ctx.fillRect(0, 0, size.width, size.height);
                    ctx.drawImage(img, 0, 0, size.width, size.height);
                    return canvas.toDataURL('image/jpeg', 0.9);
                }
                var img = document.querySelector('#search_samplecaptcha_CaptchaImage');
                    return getBase64FromImage(img, img.width, img.height);
                ") as string;

                    base_64 = img_base_64.Split(',').Last();
                }
            catch (Exception ex)
            {
                WriteLogFile.writeLog("logs.txt", "Exception in create capctcha base 64 function", ex.Message);
                Console.WriteLine("Exception in create captch base 64", ex.Message);
            }
            return base_64;
        }

        private void SaveImage(string base64, string file_name)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(base64)))
                {
                    using (Bitmap bm2 = new Bitmap(ms))
                    {
                        string capctahImagetPath = Directory.GetCurrentDirectory() + "\\solvedCaptcha\\";
                        if (Directory.Exists(capctahImagetPath) == false)
                        {
                            new DirectoryInfo(capctahImagetPath).Create();
                        }
                        bm2.Save(capctahImagetPath + file_name + ".png");
                    }
                }

            }
            catch (Exception ex)
            {
                WriteLogFile.writeLog("logs.txt", "Exception in SaveImage function", ex.Message);
                Console.WriteLine("Exception in SaveImage Function", ex.Message);
            }
        }

        public string captchaSolver(string base64, string cross_ref_number)
        {
            if (destroyedDriver == true)
            {
                return "";
            }
            string captcha_text = "";
            for (int i=1; i<=10; i++)
            {
                try
                {
                    Client client = (Client)new SocketClient("farukh", "webdir123R");
                    byte[] img = Convert.FromBase64String(base64);
                    WriteLogFile.writeLog("logs.txt", "Solving Captcha", cross_ref_number);
                    Captcha captcha;
                    try
                    {
                        captcha = client.Decode(img, 10);

                    } catch (Exception ex1)
                    {
                        WriteLogFile.writeLog("logs.txt", "Exception in captchaSolver", cross_ref_number);
                        Console.WriteLine("Exception in captchaSolver", ex1.Message);
                        throw ex1;
                    }
                    if ( captcha != null && captcha.Solved && captcha.Correct)
                    {
                        WriteLogFile.writeLog("logs.txt", "Captcha Solved " + captcha.Text, cross_ref_number);
                        Console.WriteLine("CAPTCHA {0}: {1}", captcha.Id, captcha.Text);
                        captcha_text = captcha.Text;
                        SaveImage(base64, captcha_text);
                    }
                    break;
                }
                catch (Exception ex)
                {
                    WriteLogFile.writeLog("logs.txt", "Exception in captchaSolver", cross_ref_number);
                    Console.WriteLine("Exception in captchaSolver", ex.Message);
                    if ( destroyedDriver == false)
                    {
                        captchaSolver(base64, cross_ref_number);
                    }
                    
                }
            }
            
            return captcha_text;
        }

        //public bool saveCaptcha()
        //{
        //    bool res = false;
        //    try
        //    {
        //        string img_base_64 = getCaptchaBase64();
        //        res = SaveImage(img_base_64);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Exception in saveCaptcha Function ", ex.Message);
        //    }
        //    return res;
        //}

        public bool enterCaptchaText(string captchaText, string crossRefnumber)
        {
            bool res = false;
            try
            {
                WriteLogFile.writeLog("logs.txt", "Inserting Captcha in fields...", crossRefnumber);
                IWebElement captchaField = FindElementIfExists(By.CssSelector("#CodeTextBox"));
                captchaField.SendKeys(captchaText);
                IWebElement crossRefRadio = FindElementIfExists(By.CssSelector("#CrossRefNumberOption"));
                crossRefRadio.Click();
                IWebElement crossRefField = FindElementIfExists(By.CssSelector("#CaseSearchValue"));
                crossRefField.SendKeys(crossRefnumber);
                getScreenShot(crossRefnumber + "_captchaScreen");
                IWebElement searchButton = FindElementIfExists(By.CssSelector("#SearchSubmit"));
                searchButton.Click();
                sleep(sleeptime);
                IWebElement incorrectCaptcha = null;
                try{
                    incorrectCaptcha = FindElementIfExists(By.XPath("//*[@id='MessageLabel']"));
                    WriteLogFile.writeLog("logs.txt", "Incorrect Captcha Label Found Trying Again...", crossRefnumber);
                }
                catch (Exception ex1){
                    WriteLogFile.writeLog("logs.txt", "Exception in Incorrect Captcha That means no incorrect captcha label is founded ..." + ex1.Message, crossRefnumber);
                }
                if (incorrectCaptcha != null){
                    getScreenShot(crossRefnumber + "_incorrectCaptcha");
                    res = false;
                }else
                {
                    WriteLogFile.writeLog("logs.txt", "Captcha is successfully applied Moving Next ...", crossRefnumber);
                    res = true;
                }
            }catch (Exception ex){
                WriteLogFile.writeLog("logs.txt", "Exception in Enter Captcha Text function..." + ex.Message, crossRefnumber);
                Console.WriteLine("Exception in enterCaptchaText function", ex.Message);
                getScreenShot(crossRefnumber + "e_in_captcha_text");
            }
            return res;
        }

        public bool openingMainPage()
        {
            bool res = false;
            try
            {
                _d.Navigate().GoToUrl("https://www.clarkcountycourts.us/Anonymous/default.aspx");
                sleep(sleeptime);
                IWebElement lasVegasLink = FindElementIfExists(By.CssSelector("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > a:nth-child(14)"));
                lasVegasLink.Click();
                //getScreenShot("lasvegas");
                sleep(sleeptime);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception in openingMainPage function", ex.Message);
            }
            return res;
        }

        public void ViewAndPayCriminalPage()
        {
            try
            {
                IWebElement viewAndPayCriminal = this.FindElementIfExists(By.CssSelector("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > a:nth-child(8)"));
                viewAndPayCriminal.Click();
                sleep(sleeptime);
            }
            catch (Exception ex)
            {

                Console.WriteLine("Exception in ViewAndPayCriminalPage  Function", ex.Message);
            }
        }

        public bool LookingAndSolvingCaptcha(string cross_ref_number)
        {
            bool navigated = false ;
            try
            {
                for (int i = 1; i <= 5; i++)
                {
                    string img_base_64 = getCaptchaBase64();
                    WriteLogFile.writeLog("logs.txt", "Getting Image Base 64...", cross_ref_number);
                    string captcha_text = captchaSolver(img_base_64, cross_ref_number);
                    if (captcha_text.Length > 0)
                    {
                        WriteLogFile.writeLog("logs.txt", "Captcha Text found ..." + captcha_text + "  ", cross_ref_number);
                        navigated = enterCaptchaText(captcha_text, cross_ref_number);
                        if (navigated)
                        {
                            WriteLogFile.writeLog("logs.txt", "Captcha Resolved On Result Page...", cross_ref_number);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogFile.writeLog("logs.txt", "Exception in LookingAndSolvingCaptcha Function" + ex.Message, cross_ref_number);
                Console.WriteLine("Exception in LookingAndSolvingCaptcha  Function", ex.Message);
            }
            return navigated;
        }

        public int extractData(string cross_ref_number)
        {
            int res = 0;
            try
            {
                WriteLogFile.writeLog("logs.txt", "Extracting Data ...", cross_ref_number);
                getScreenShot(cross_ref_number + "_extract_data_main");
                List<IWebElement> trsForLength = FindElementsIfExists(By.CssSelector("body > table:nth-child(5) > tbody > tr"));
                int counter = 1;
                for (int trIndex = 0; trIndex < trsForLength.Count; trIndex ++)
                {
                    List<IWebElement> trs = FindElementsIfExists(By.CssSelector("body > table:nth-child(5) > tbody > tr"));
                    if (trIndex > 1)
                    {
                        var tr = trs[trIndex];
                        IWebElement caseLink = null;
                        IWebElement defendantName = null;
                        try
                        {
                            caseLink = FindElementOnObject(By.CssSelector("td > a"), tr);
                            defendantName = FindElementOnObject(By.CssSelector("td:nth-child(3) > div"), tr);
                        }
                        catch (Exception exx) {
                            WriteLogFile.writeLog("logs.txt", "Exception in Finding Case Links " + exx.Message , cross_ref_number);
                            getScreenShot(cross_ref_number + "_unable_to_find_links");
                            res = 1;
                        }
                        if (caseLink != null && defendantName !=null )
                        {
                            string file_name = defendantName.Text;
                            caseLink.Click();
                            sleep(sleeptime);
                            getScreenShot(cross_ref_number + "_extract_data");
                            DateTime now = DateTime.Now;
                            var header = "<div style='width:100%;margin-left:0.5cm;text-align:left;font-size:30px;'>Public Access Downloader</div>";
                            var footer = "<div style='width:100%;margin-left:0.5cm;text-align:left;font-size:10px;'>Date:" + (now).ToString() +"</div>";
                            var commandParameters = new Dictionary<string, object>
                            {
                                {"displayHeaderFooter", true },
                                { "format", "A4" },
                                { "scale", 0.9 },
                                { "behavior", "allow" },
                                { "footerTemplate", footer},
                            };
                            var printOutput = (Dictionary<string, object>)_d.ExecuteChromeCommandWithResult("Page.printToPDF", commandParameters);
                            var pdf_data = Convert.FromBase64String(printOutput["data"] as string);
                            
                            WriteLogFile.writeLog("logs.txt", "Createing Pdf", cross_ref_number);
                            if (DownloadPath == null)
                            {
                                DownloadPath = Directory.GetCurrentDirectory() + "\\Cases Data\\";
                            }
                            //var sub_directory = DownloadPath + "\\" + cross_ref_number;
                            //if (Directory.Exists(sub_directory) == false)
                            //{
                            //    new DirectoryInfo(sub_directory).Create();
                            //}
                            string file_n = file_name + "_" + cross_ref_number + "_" + DateTime.Now.ToFileTime() + ".pdf";
                            file_name_global = file_name_global + "\n" + file_n;
                            string file = DownloadPath + "\\" + file_n;
                            WriteLogFile.writeLog("logs.txt", "Created pdf " + file_n, cross_ref_number);
                            File.WriteAllBytes(file, pdf_data);
                            res = 2;
                        }
                        
                        _d.Navigate().Back();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogFile.writeLog("logs.txt", "Exception in  Extract Data function" + ex.Message, cross_ref_number);
                Console.WriteLine("Exception in Extract Data function", ex.Message);
                getScreenShot(cross_ref_number + "_e_in_extract_data");
                res = 0;
            }
            return res;
        }


        private IWebElement FindElementIfExists(By by)
        {
            int times = 1;
            while (true)
            {
                try
                {
                    ReadOnlyCollection<IWebElement> elements = _d.FindElements(by);
                    return elements.Count >= 1 ? elements.First<IWebElement>() : (IWebElement)null;
                }
                catch (Exception ex)
                {
                    sleep(1000);
                    times += 1;
                    if (times == 10 || destroyedDriver == true)
                    {
                        WriteLogFile.writeLog("errors_logs.txt", "Exception in FindElementIfExists" + ex.Message, "");
                        WriteLogFile.writeLog("errors_logs.txt", "Network Issue", "");
                        getScreenShot("network_issues");
                        return null;
                    }
                }
            }
            
        }

        private List<IWebElement> FindElementsIfExists(By by)
        {
            int times = 1;
            while (true)
            {
                try
                {
                    var webElements = _d.FindElements(by);
                    return webElements.ToList();
                }
                catch (Exception ex)
                {
                    sleep(1000);
                    times += 1;
                    if (times == 10 || destroyedDriver == true)
                    {
                        WriteLogFile.writeLog("errors_logs.txt", "Exception in FindElementsIfExists" + ex.Message, "");
                        WriteLogFile.writeLog("errors_logs.txt", "Network Issue", "");
                        getScreenShot("network_issues");
                        return null;
                    }
                }
            }
        }

        private IWebElement FindElementOnObject(By by, IWebElement el)
        {
            int times = 1;
            while (true)
            {
                try
                {
                    var webElement = el.FindElement(by);
                    return webElement;
                }
                catch (Exception ex)
                {
                    sleep(1000);
                    times += 1;
                    if (times == 10 || destroyedDriver == true)
                    {
                        WriteLogFile.writeLog("errors_logs.txt", "Exception in FindElementOnObject" + ex.Message, "");
                        WriteLogFile.writeLog("errors_logs.txt", "Network Issue", "");
                        getScreenShot("network_issues");
                        return null;
                    }
                }
            }
        }

        #region Destructors
        ~Locate()
        {
            if (_d != null)
                _d.Quit();
        }

        public void Dispose()
        {
            try
            {
                _d.Quit();
                destroyedDriver = true;

            } catch (Exception e) { }
        }
        #endregion

        public class WriteLogFile
        {
            public static bool writeLog(string strFileName, string strMessage, string cross_ref_number)
            {
                try
                {
                    FileStream objFilestream = new FileStream(string.Format("{0}\\{1}", Locate.logsPath, cross_ref_number + strFileName), FileMode.Append, FileAccess.Write);
                    StreamWriter objStreamWriter = new StreamWriter((Stream)objFilestream);
                    objStreamWriter.WriteLine(strMessage);
                    objStreamWriter.Close();
                    objFilestream.Close();
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
        }

    }
}
