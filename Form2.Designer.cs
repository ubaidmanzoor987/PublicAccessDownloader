using System.Windows.Forms;

namespace ClarkCountryCaseDownloader
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private OpenFileDialog openFileDialog1;
        private FolderBrowserDialog folderBrowserDialog;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.stop_button = new System.Windows.Forms.Button();
            this.downloadedFilesBox = new System.Windows.Forms.RichTextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.threadInput = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.manual_cross_ref = new System.Windows.Forms.CheckBox();
            this.excel_file_input = new System.Windows.Forms.TextBox();
            this.cross_Ref = new System.Windows.Forms.TextBox();
            this.excel_file_btn = new System.Windows.Forms.Button();
            this.download_directory_input = new System.Windows.Forms.TextBox();
            this.download_path_button = new System.Windows.Forms.Button();
            this.start_download_button = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.activityBar = new System.Windows.Forms.RichTextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.stop_button);
            this.panel1.Controls.Add(this.downloadedFilesBox);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.threadInput);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.manual_cross_ref);
            this.panel1.Controls.Add(this.excel_file_input);
            this.panel1.Controls.Add(this.cross_Ref);
            this.panel1.Controls.Add(this.excel_file_btn);
            this.panel1.Controls.Add(this.download_directory_input);
            this.panel1.Controls.Add(this.download_path_button);
            this.panel1.Controls.Add(this.start_download_button);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.activityBar);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(12, 8);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(750, 677);
            this.panel1.TabIndex = 0;
            // 
            // stop_button
            // 
            this.stop_button.Location = new System.Drawing.Point(513, 289);
            this.stop_button.Name = "stop_button";
            this.stop_button.Size = new System.Drawing.Size(175, 38);
            this.stop_button.TabIndex = 31;
            this.stop_button.Text = "Stop";
            this.stop_button.UseVisualStyleBackColor = true;
            this.stop_button.Visible = false;
            this.stop_button.Click += new System.EventHandler(this.stop_button_Click);
            // 
            // downloadedFilesBox
            // 
            this.downloadedFilesBox.BackColor = System.Drawing.SystemColors.HighlightText;
            this.downloadedFilesBox.Location = new System.Drawing.Point(384, 452);
            this.downloadedFilesBox.Name = "downloadedFilesBox";
            this.downloadedFilesBox.ReadOnly = true;
            this.downloadedFilesBox.Size = new System.Drawing.Size(304, 178);
            this.downloadedFilesBox.TabIndex = 29;
            this.downloadedFilesBox.Text = "";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label7.Location = new System.Drawing.Point(380, 413);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(194, 20);
            this.label7.TabIndex = 30;
            this.label7.Text = "List of Downloaded Files";
            // 
            // threadInput
            // 
            this.threadInput.BackColor = System.Drawing.SystemColors.HighlightText;
            this.threadInput.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.threadInput.Location = new System.Drawing.Point(207, 289);
            this.threadInput.Multiline = true;
            this.threadInput.Name = "threadInput";
            this.threadInput.Size = new System.Drawing.Size(295, 38);
            this.threadInput.TabIndex = 7;
            this.threadInput.TextChanged += new System.EventHandler(this.threadInput_TextChanged);
            this.threadInput.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cross_Ref_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 289);
            this.label2.Name = "label2";
            this.label2.Padding = new System.Windows.Forms.Padding(5);
            this.label2.Size = new System.Drawing.Size(133, 26);
            this.label2.TabIndex = 28;
            this.label2.Text = "Number of Threads";
            // 
            // manual_cross_ref
            // 
            this.manual_cross_ref.AutoSize = true;
            this.manual_cross_ref.Location = new System.Drawing.Point(517, 88);
            this.manual_cross_ref.Margin = new System.Windows.Forms.Padding(5);
            this.manual_cross_ref.Name = "manual_cross_ref";
            this.manual_cross_ref.Size = new System.Drawing.Size(163, 20);
            this.manual_cross_ref.TabIndex = 2;
            this.manual_cross_ref.Text = "Use Manual Cross Ref";
            this.manual_cross_ref.UseVisualStyleBackColor = true;
            this.manual_cross_ref.CheckedChanged += new System.EventHandler(this.manual_cross_ref_CheckedChanged);
            // 
            // excel_file_input
            // 
            this.excel_file_input.BackColor = System.Drawing.SystemColors.HighlightText;
            this.excel_file_input.Enabled = false;
            this.excel_file_input.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.excel_file_input.Location = new System.Drawing.Point(207, 149);
            this.excel_file_input.Multiline = true;
            this.excel_file_input.Name = "excel_file_input";
            this.excel_file_input.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.excel_file_input.Size = new System.Drawing.Size(295, 38);
            this.excel_file_input.TabIndex = 3;
            // 
            // cross_Ref
            // 
            this.cross_Ref.BackColor = System.Drawing.SystemColors.HighlightText;
            this.cross_Ref.Enabled = false;
            this.cross_Ref.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cross_Ref.Location = new System.Drawing.Point(207, 88);
            this.cross_Ref.Multiline = true;
            this.cross_Ref.Name = "cross_Ref";
            this.cross_Ref.Size = new System.Drawing.Size(295, 37);
            this.cross_Ref.TabIndex = 1;
            this.cross_Ref.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cross_Ref_KeyPress);
            // 
            // excel_file_btn
            // 
            this.excel_file_btn.Location = new System.Drawing.Point(517, 149);
            this.excel_file_btn.Name = "excel_file_btn";
            this.excel_file_btn.Size = new System.Drawing.Size(175, 33);
            this.excel_file_btn.TabIndex = 4;
            this.excel_file_btn.Text = "Browse";
            this.excel_file_btn.UseVisualStyleBackColor = true;
            this.excel_file_btn.Click += new System.EventHandler(this.excel_file_btn_Click);
            // 
            // download_directory_input
            // 
            this.download_directory_input.BackColor = System.Drawing.SystemColors.HighlightText;
            this.download_directory_input.Enabled = false;
            this.download_directory_input.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.download_directory_input.Location = new System.Drawing.Point(207, 218);
            this.download_directory_input.Multiline = true;
            this.download_directory_input.Name = "download_directory_input";
            this.download_directory_input.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.download_directory_input.Size = new System.Drawing.Size(295, 38);
            this.download_directory_input.TabIndex = 5;
            // 
            // download_path_button
            // 
            this.download_path_button.Location = new System.Drawing.Point(517, 218);
            this.download_path_button.Name = "download_path_button";
            this.download_path_button.Size = new System.Drawing.Size(175, 33);
            this.download_path_button.TabIndex = 6;
            this.download_path_button.Text = "Browse";
            this.download_path_button.UseVisualStyleBackColor = true;
            this.download_path_button.Click += new System.EventHandler(this.download_path_button_Click);
            // 
            // start_download_button
            // 
            this.start_download_button.Location = new System.Drawing.Point(207, 356);
            this.start_download_button.Name = "start_download_button";
            this.start_download_button.Size = new System.Drawing.Size(295, 43);
            this.start_download_button.TabIndex = 8;
            this.start_download_button.Text = "Start Download";
            this.start_download_button.UseVisualStyleBackColor = true;
            this.start_download_button.Click += new System.EventHandler(this.start_download_button_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(23, 218);
            this.label5.Name = "label5";
            this.label5.Padding = new System.Windows.Forms.Padding(5);
            this.label5.Size = new System.Drawing.Size(135, 26);
            this.label5.TabIndex = 27;
            this.label5.Text = "Download Directory";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(23, 149);
            this.label4.Name = "label4";
            this.label4.Padding = new System.Windows.Forms.Padding(5);
            this.label4.Size = new System.Drawing.Size(123, 26);
            this.label4.TabIndex = 24;
            this.label4.Text = "Browse Excel File";
            // 
            // activityBar
            // 
            this.activityBar.BackColor = System.Drawing.SystemColors.HighlightText;
            this.activityBar.Location = new System.Drawing.Point(26, 452);
            this.activityBar.Name = "activityBar";
            this.activityBar.ReadOnly = true;
            this.activityBar.Size = new System.Drawing.Size(344, 178);
            this.activityBar.TabIndex = 9;
            this.activityBar.Text = "";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label6.Location = new System.Drawing.Point(22, 413);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(105, 20);
            this.label6.TabIndex = 17;
            this.label6.Text = "Activity Logs";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 88);
            this.label3.Name = "label3";
            this.label3.Padding = new System.Windows.Forms.Padding(5);
            this.label3.Size = new System.Drawing.Size(161, 26);
            this.label3.TabIndex = 13;
            this.label3.Text = "Enter Cross Ref Number";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(221, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(350, 32);
            this.label1.TabIndex = 10;
            this.label1.Text = "Public Access Downloader";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // Form2
            // 
            this.ClientSize = new System.Drawing.Size(761, 685);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private WindowsFormsSynchronizationContext mUiContext;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button start_download_button;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox cross_Ref;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox activityBar;
        private System.Windows.Forms.RichTextBox downloadedFilesBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox excel_file_input;
        private System.Windows.Forms.Button excel_file_btn;
        private System.Windows.Forms.Button download_path_button;
        private System.Windows.Forms.TextBox download_directory_input;
        private System.Windows.Forms.Label label5;
        private CheckBox manual_cross_ref;
        private Label label2;
        private TextBox threadInput;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private Label label7;
        private Button stop_button;
    }
}
