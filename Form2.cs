using System.Windows.Forms;
using System;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace ClarkCountryCaseDownloader
{
    public partial class Form2 : Form
    {
        Thread backgroudThread;
        delegate void showMessage(string msg);
        List<Locate> locateDrivers;
        Task task_factory;
        CancellationTokenSource cts; 

        public Form2()
        {
            locateDrivers = new List<Locate>();
            InitializeComponent();
            download_directory_input.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            threadInput.Text = "1";
            FormClosed += new FormClosedEventHandler(frmCountyDownloader_FormClosed);
        }

        private void Form2_UnLoad(object sender, EventArgs e)
        {
            backgroudThread.Abort();
        }



        private void frmCountyDownloader_FormClosed(object sender, FormClosedEventArgs e)
        {
            stop_function_process();
        }

        private void stop_function_process()
        {
            var th = new Thread(() => { CleanDrivers(); });
            th.Start();
        }

        protected void CleanDrivers()
        {
            // Handle the Event here.
            if (cts != null)
            {
                cts.Cancel();
                cts.Dispose();
            }
            if (locateDrivers.Count > 0)
            {
                for(int i = 0; i<locateDrivers.Count; i++)
                {
                    if (locateDrivers[i] != null)
                        locateDrivers[i].Dispose();
                }
            }
            
        }

        protected override void Dispose(bool disposing)
        {
            if ((!disposing ? false : this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        void showStatusMessage(string msg)
        {
            try
            {
                if (activityBar.InvokeRequired)
                {
                    activityBar.Invoke(new Action<string>(showStatusMessage), msg);
                    return;
                }
                activityBar.AppendText(msg + "\r\n");
            } catch (Exception e) { }
        }

        void showDownloadedFiles(string msg)
        {
            try
            {
                if (downloadedFilesBox.InvokeRequired)
                {
                    downloadedFilesBox.Invoke(new Action<string>(showDownloadedFiles), msg);
                    return;
                }
                downloadedFilesBox.AppendText(msg + "\r\n");
            }
            catch (Exception e) { }
        }

        void stop_btn(bool enable)
        {
            if (stop_button.InvokeRequired)
            {
                start_download_button.Invoke(new Action<bool>(stop_btn), enable);
                return;
            }
            stop_button.Enabled = enable;
        }

        void start_download_btn(bool enable)
        {
            if (start_download_button.InvokeRequired)
            {
                start_download_button.Invoke(new Action<bool>(start_download_btn), enable);
                return;
            }
            start_download_button.Enabled = enable;
        }

        void toggleBrowseButton(bool eanble)
        {
            try
            {
                if (download_path_button.InvokeRequired)
                {
                    download_path_button.Invoke(new Action<bool>(toggleBrowseButton), eanble);
                    return;
                }
                download_path_button.Enabled = eanble;
            } catch(Exception e) { }
            
        }

        void toggleCrossRef(bool enable)
        {
            try
            {
                if (cross_Ref.InvokeRequired)
                {
                    cross_Ref.Invoke(new Action<bool>(toggleCrossRef), enable);
                    return;
                }
                cross_Ref.Enabled = enable;
            } catch(Exception e) { }
            
        }

        void toggleThreadField(bool enable)
        {
            try
            {
                if (threadInput.InvokeRequired)
                {
                    threadInput.Invoke(new Action<bool>(toggleThreadField), enable);
                    return;
                }
                threadInput.Enabled = enable;
            }
            catch (Exception e) { }
        }

        void toggleExcelfileButton(bool enable)
        {
            try
            {
                if (excel_file_btn.InvokeRequired)
                {
                    excel_file_btn.Invoke(new Action<bool>(toggleExcelfileButton), enable);
                    return;
                }
                excel_file_btn.Enabled = enable;
            } catch (Exception e) { }
        }

        private void start_download_button_Click(object sender, System.EventArgs e)
        {
            toggleBrowseButton(false);
            start_download_btn(false);
            if (manual_cross_ref.Checked)
            {
                if (cross_Ref.Text.Length == 0)
                {
                    MessageBox.Show("Enter Cross Ref Number Then Press Download");
                    start_download_btn(true);
                    return;
                }
                   
            } else
            {
                if (excel_file_input.Text.Length == 0)
                {
                    MessageBox.Show("Browse Excel File Then Press Download");
                    start_download_btn(true);
                    return;
                }
            }
            backgroudThread = new Thread(() =>
            {
                ProcessDownload();
            });
            backgroudThread.Start();
        }

        private bool initializingDriver(Locate loc)
        {
            loc.DownloadPath = download_directory_input.Text;
            bool res = false;
            for (int i = 1; i<=5; i++)
            {
                bool initialized = loc.InitializeDriver();
                if (initialized)
                {
                    res = true;
                    break;
                } else
                {
                    showStatusMessage("Unable To Initializing Driver. Retrying ...");
                }
            }
            
            return res;
        }

        private void ProcessDownload()
        {
            mUiContext = new WindowsFormsSynchronizationContext();
            if (manual_cross_ref.Checked)
            {
                Locate locate = new Locate();
                locateDrivers.Add(locate);
                showStatusMessage("Initializing Driver...");
                if (initializingDriver(locate))
                {
                    showStatusMessage("Initalized");
                    toggleCrossRef(true);
                    toggleExcelfileButton(false);
                    string cross_ref_number = cross_Ref.Text;
                    Locate.WriteLogFile.writeLog("logs.txt", string.Format("{0} @ {1}", "Log is Created at", DateTime.Now), cross_ref_number);
                    Locate.WriteLogFile.writeLog("logs.txt", "Intialzed Driver Successfully", cross_ref_number);
                    toggleCrossRef(false);
                    showStatusMessage("Starting Scrapping...");
                    Locate.WriteLogFile.writeLog("logs.txt", "Starting Scrapping...", cross_ref_number);
                    Locate.WriteLogFile.writeLog("logs.txt", "Opening Las Vegas Page...", cross_ref_number);
                    showStatusMessage("Opening Las Vegas Page...");
                    locate.openingMainPage();
                    showStatusMessage("Opening View And Pay Criminal Page...");
                    locate.ViewAndPayCriminalPage();
                    Locate.WriteLogFile.writeLog("logs.txt", "Opening View And Pay Criminal Page...", cross_ref_number);
                    showStatusMessage("Solving Captcha and Inserting Cross Ref Number...");
                    Locate.WriteLogFile.writeLog("logs.txt", "Solving Captcha and Inserting Cross Ref Number...", cross_ref_number);
                    bool isOnDataPage = locate.LookingAndSolvingCaptcha(cross_ref_number);
                    if (isOnDataPage)
                    {
                        showStatusMessage("Extracting And Saving Data");
                        int res = locate.extractData(cross_ref_number);
                        if (res == 2)
                        {
                            Locate.WriteLogFile.writeLog("logs.txt", "Extraction Complete Cross Ref Number ", cross_ref_number);
                            showDownloadedFiles(locate.file_name_global);
                            showStatusMessage("Extraction Complete Cross Ref Number " + cross_ref_number + " Downloaded");
                        }
                        else if (res == 1)
                        {
                            Locate.WriteLogFile.writeLog("logs.txt", "Incorrect Cross Ref Number " + cross_ref_number, cross_ref_number);
                            showStatusMessage("Incorrect Cross Ref Number " + cross_ref_number);
                        } else
                        {
                            showStatusMessage("Unable To Extract Data");
                            Locate.WriteLogFile.writeLog("errors_logs.txt", "Failed To Initalize Driver...", "");
                        }
                    }
                    locate.Dispose();
                    showStatusMessage("Enter Cross Ref Number");
                    start_download_btn(true);
                    toggleCrossRef(true);
                }
                else
                {
                    showStatusMessage("Failed To Initalize Driver...");
                    Locate.WriteLogFile.writeLog("errors_logs.txt", "Failed To Initalize Driver...", "");
                    toggleThreadField(true);
                }
            }
            else
            {
                toggleCrossRef(false);
                start_download_btn(false);
                toggleThreadField(false);
                readExcelAndStartoperations(excel_file_input.Text);
            }
        }

        private void readExcelAndStartoperations(string sFile)
        {
            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(sFile);
                Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];
                Excel.Range xlRange = xlWorkSheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int event_number_index = -1;
                List<string> cross_ref_numbers = new List<string>();
                for (int j = 1; j <= colCount; j++)
                {
                    string col_name = xlRange.Cells[1, j].Value2.ToString();
                    if (col_name.ToLower() == "event number")
                    {
                        event_number_index = j;
                    }
                }

                for (int i = 2; i <= rowCount; i++)
                {
                    string cross_ref_number = Convert.ToString(xlRange.Cells[i, event_number_index].Value2);
                    if (cross_ref_number != null)
                    {
                        cross_ref_numbers.Add(cross_ref_number);
                    }

                }
                xlWorkBook.Close();
                xlApp.Quit();
                cts = new CancellationTokenSource();
                try
                {
                    var part = Partitioner.Create(cross_ref_numbers);
                    task_factory = Task.Factory.StartNew<ParallelLoopResult>(() => Parallel.ForEach(part, new ParallelOptions()
                    {
                        MaxDegreeOfParallelism = int.Parse(threadInput.Text),
                        CancellationToken = cts.Token

                    }, (refNum) =>
                    {
                        commonOperations(refNum);
                    }));
                    task_factory.Wait();
                    showStatusMessage("All Cross Ref Number Successfully Downloaded");
                    start_download_btn(true);
                    toggleThreadField(true);
                } catch (Exception ex) { }
                
            }
            catch (Exception ex)
            {
                Locate.WriteLogFile.writeLog("errors_logs.txt", "Exception in Reading Excel File Path" + ex.Message, "");
                Console.WriteLine("Exception in Reading Excel File Path, ", ex.Message);
            }

        }

        private void commonOperations(string cross_ref_number)
        {
            try
            {
                Locate locate = new Locate();
                locateDrivers.Add(locate);
                bool connect = initializingDriver(locate);
                if (connect)
                {
                    showStatusMessage("Downloading Cross Ref Number..." + (cross_ref_number).ToString());
                    showStatusMessage("Opening Las Vegas Page...");
                    locate.openingMainPage();
                    showStatusMessage("Opening View And Pay Criminal Page...");
                    locate.ViewAndPayCriminalPage();
                    showStatusMessage("Solving Captcha and Inserting Cross Ref Number...");
                    bool isOnDataPage = locate.LookingAndSolvingCaptcha(cross_ref_number);
                    if (isOnDataPage)
                    {
                        locate.DownloadPath = download_directory_input.Text;
                        showStatusMessage("Extracting And Saving Data...");
                        int res = locate.extractData(cross_ref_number);
                        if (res == 2)
                        {
                            Locate.WriteLogFile.writeLog("logs.txt", "Extraction Complete Cross Ref Number " + cross_ref_number + " Downloaded", cross_ref_number);
                            showDownloadedFiles(locate.file_name_global);
                            showStatusMessage("Extraction Complete Cross Ref Number " + cross_ref_number + " Downloaded");
                        } else if (res == 1){
                            Locate.WriteLogFile.writeLog("logs.txt", "Incorrect Cross Ref Number " , cross_ref_number);
                            showStatusMessage("Incorrect Cross Ref Number " + cross_ref_number);
                        }
                        else {
                            Locate.WriteLogFile.writeLog("errors_logs.txt", "Exception in Extract Data function", "");
                        }
                    }
                }
                locate.Dispose();
            } catch (Exception e) { 
                Locate.WriteLogFile.writeLog("errors_logs.txt", "Exception in Reading Excel File Path" + e.Message, "");
                Console.WriteLine("Exception in Reading Excel File Path, ", e.Message);
            }

        }

        private void excel_file_btn_Click(object sender, System.EventArgs e)
        {
            try
            {
                openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Browse Excel File";
                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;
                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Filter = "Excel File|*.xlsx;*.xls";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string file_name = openFileDialog1.FileName;
                        excel_file_input.Text = file_name;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Security error.\n\nError message: {ex.Message}\n\n" +
                        $"Details:\n\n{ex.StackTrace}");
                    }
                }
            } catch (Exception ex)
            {
                Locate.WriteLogFile.writeLog("errors_logs.txt", "Exception in excel_file_btn_Click" + ex.Message, "");
                Console.WriteLine("Exception in excel_file_btn_Click, ", ex.Message);
            }
        }

        private void download_path_button_Click(object sender, System.EventArgs e)
        {
            folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
            {
                string folder_path = folderBrowserDialog.SelectedPath;
                download_directory_input.Text = folder_path;
            }
        }

        private void manual_cross_ref_CheckedChanged(object sender, System.EventArgs e)
        {
            if (manual_cross_ref.Checked)
            {
                toggleExcelfileButton(false);
                toggleCrossRef(true);
                toggleThreadField(false);
            }
            else
            {
                toggleBrowseButton(true);
                toggleExcelfileButton(true);
                toggleCrossRef(false);
                toggleThreadField(true);
            }
        }

        private void threadInput_TextChanged(object sender, EventArgs e)
        {

            double parsedValue;

            if (!double.TryParse(threadInput.Text, out parsedValue))
            {
                threadInput.Text = "";
            } else
            {
                try
                {
                    if (int.Parse(threadInput.Text) > 20)
                    {
                        MessageBox.Show("Max Value is 20");
                        threadInput.Text = "1";
                    }
                    else if (int.Parse(threadInput.Text) < 1)
                    {
                        MessageBox.Show("Min Value is 1");
                        threadInput.Text = "1";
                    }
                } catch (Exception ex)
                {
                    MessageBox.Show("Invalid Value Provided");
                    threadInput.Text = "1";
                }

            }
        }

        private void cross_Ref_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                start_download_button.PerformClick();
            }
        }

        private void threadInput_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                start_download_button.PerformClick();
            }
        }

        private void stop_button_Click(object sender, EventArgs e)
        {
            try
            {

                //stop_function_process();
                //toggleBrowseButton(true);
                //toggleExcelfileButton(true);
                //start_download_btn(true);
            } catch (Exception ex)
            {
                Locate.WriteLogFile.writeLog("errors_logs.txt", "Exception in stop_button_Click" + ex.Message, "");
                Console.WriteLine("Exception in stop_button_Click, ", ex.Message);
            }
        }
    }
}
