namespace WordAppGUI
{
    partial class FrmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.btnLoadFile = new System.Windows.Forms.Button();
            this.flowLayoutPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnOlustur = new System.Windows.Forms.Button();
            this.cmbFileName = new System.Windows.Forms.ComboBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.label2 = new System.Windows.Forms.Label();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.chkPdf = new System.Windows.Forms.CheckBox();
            this.btnWordOpen = new System.Windows.Forms.Button();
            this.rbExcel = new System.Windows.Forms.RadioButton();
            this.rbWord = new System.Windows.Forms.RadioButton();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.cmbFileName2 = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.BackColor = System.Drawing.Color.Green;
            this.btnLoadFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLoadFile.ForeColor = System.Drawing.Color.White;
            this.btnLoadFile.Location = new System.Drawing.Point(9, 18);
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.Size = new System.Drawing.Size(87, 23);
            this.btnLoadFile.TabIndex = 0;
            this.btnLoadFile.Text = "Veri Dosyası";
            this.btnLoadFile.UseVisualStyleBackColor = false;
            this.btnLoadFile.Click += new System.EventHandler(this.btnLoadFile_Click);
            // 
            // flowLayoutPanel
            // 
            this.flowLayoutPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.flowLayoutPanel.AutoScroll = true;
            this.flowLayoutPanel.BackColor = System.Drawing.Color.SeaShell;
            this.flowLayoutPanel.Location = new System.Drawing.Point(15, 121);
            this.flowLayoutPanel.Name = "flowLayoutPanel";
            this.flowLayoutPanel.Size = new System.Drawing.Size(301, 347);
            this.flowLayoutPanel.TabIndex = 1;
            // 
            // btnOlustur
            // 
            this.btnOlustur.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOlustur.BackColor = System.Drawing.Color.DarkRed;
            this.btnOlustur.Enabled = false;
            this.btnOlustur.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOlustur.ForeColor = System.Drawing.Color.White;
            this.btnOlustur.Location = new System.Drawing.Point(622, 70);
            this.btnOlustur.Name = "btnOlustur";
            this.btnOlustur.Size = new System.Drawing.Size(104, 23);
            this.btnOlustur.TabIndex = 2;
            this.btnOlustur.Text = "Oluştur";
            this.btnOlustur.UseVisualStyleBackColor = false;
            this.btnOlustur.Click += new System.EventHandler(this.btnOlustur_Click);
            // 
            // cmbFileName
            // 
            this.cmbFileName.BackColor = System.Drawing.Color.White;
            this.cmbFileName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFileName.FormattingEnabled = true;
            this.cmbFileName.Location = new System.Drawing.Point(6, 20);
            this.cmbFileName.Name = "cmbFileName";
            this.cmbFileName.Size = new System.Drawing.Size(111, 21);
            this.cmbFileName.TabIndex = 3;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(12, 71);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(604, 22);
            this.progressBar.TabIndex = 4;
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerReportsProgress = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 105);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Alanlar";
            // 
            // richTextBox
            // 
            this.richTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBox.BackColor = System.Drawing.Color.MintCream;
            this.richTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox.Location = new System.Drawing.Point(322, 121);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.ReadOnly = true;
            this.richTextBox.Size = new System.Drawing.Size(404, 347);
            this.richTextBox.TabIndex = 6;
            this.richTextBox.Text = "";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(319, 105);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(25, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Log";
            // 
            // chkPdf
            // 
            this.chkPdf.AutoSize = true;
            this.chkPdf.Location = new System.Drawing.Point(10, 22);
            this.chkPdf.Name = "chkPdf";
            this.chkPdf.Size = new System.Drawing.Size(88, 17);
            this.chkPdf.TabIndex = 7;
            this.chkPdf.Text = "PDF Oluşsun";
            this.chkPdf.UseVisualStyleBackColor = true;
            // 
            // btnWordOpen
            // 
            this.btnWordOpen.BackColor = System.Drawing.Color.RoyalBlue;
            this.btnWordOpen.Enabled = false;
            this.btnWordOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnWordOpen.ForeColor = System.Drawing.Color.White;
            this.btnWordOpen.Location = new System.Drawing.Point(102, 18);
            this.btnWordOpen.Name = "btnWordOpen";
            this.btnWordOpen.Size = new System.Drawing.Size(87, 23);
            this.btnWordOpen.TabIndex = 8;
            this.btnWordOpen.Text = "Şablon";
            this.btnWordOpen.UseVisualStyleBackColor = false;
            this.btnWordOpen.Click += new System.EventHandler(this.btnWordOpen_Click);
            // 
            // rbExcel
            // 
            this.rbExcel.AutoSize = true;
            this.rbExcel.Location = new System.Drawing.Point(22, 20);
            this.rbExcel.Name = "rbExcel";
            this.rbExcel.Size = new System.Drawing.Size(51, 17);
            this.rbExcel.TabIndex = 9;
            this.rbExcel.Text = "Excel";
            this.rbExcel.UseVisualStyleBackColor = true;
            // 
            // rbWord
            // 
            this.rbWord.AutoSize = true;
            this.rbWord.Checked = true;
            this.rbWord.Location = new System.Drawing.Point(79, 20);
            this.rbWord.Name = "rbWord";
            this.rbWord.Size = new System.Drawing.Size(51, 17);
            this.rbWord.TabIndex = 9;
            this.rbWord.TabStop = true;
            this.rbWord.Text = "Word";
            this.rbWord.UseVisualStyleBackColor = true;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline;
            this.linkLabel1.LinkColor = System.Drawing.Color.MediumSlateBlue;
            this.linkLabel1.Location = new System.Drawing.Point(686, 98);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(39, 13);
            this.linkLabel1.TabIndex = 10;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Yardım";
            this.linkLabel1.Click += new System.EventHandler(this.linkLabel1_Click);
            // 
            // cmbFileName2
            // 
            this.cmbFileName2.BackColor = System.Drawing.Color.White;
            this.cmbFileName2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFileName2.FormattingEnabled = true;
            this.cmbFileName2.Location = new System.Drawing.Point(123, 20);
            this.cmbFileName2.Name = "cmbFileName2";
            this.cmbFileName2.Size = new System.Drawing.Size(111, 21);
            this.cmbFileName2.TabIndex = 3;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rbExcel);
            this.groupBox1.Controls.Add(this.rbWord);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(150, 53);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Şablon Dosyası";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnLoadFile);
            this.groupBox2.Controls.Add(this.btnWordOpen);
            this.groupBox2.Location = new System.Drawing.Point(168, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 53);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Dosya Seçimi";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.cmbFileName);
            this.groupBox3.Controls.Add(this.cmbFileName2);
            this.groupBox3.Location = new System.Drawing.Point(374, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(242, 53);
            this.groupBox3.TabIndex = 13;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Dosya İsmi";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.chkPdf);
            this.groupBox4.Location = new System.Drawing.Point(622, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(104, 53);
            this.groupBox4.TabIndex = 14;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Seçenekler";
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(738, 478);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.richTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnOlustur);
            this.Controls.Add(this.flowLayoutPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Word Excel Maker";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLoadFile;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel;
        private System.Windows.Forms.Button btnOlustur;
        private System.Windows.Forms.ComboBox cmbFileName;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkPdf;
        private System.Windows.Forms.Button btnWordOpen;
        private System.Windows.Forms.RadioButton rbExcel;
        private System.Windows.Forms.RadioButton rbWord;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.ComboBox cmbFileName2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
    }
}

