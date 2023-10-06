using ExcelDataReader;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WordAppGUI.UserControls;
using Word = Microsoft.Office.Interop.Word;

namespace WordAppGUI
{
    public partial class FrmMain : Form
    {
        private string jsonFileContent = String.Empty;
        private string ExcelFilePath = "";
        private string WordFilePath = "";
        private Dictionary<string, string> WordParams = new Dictionary<string, string>();
        private string[] typeArray = null;
        private string OutputFolder = String.Empty;
        private System.Data.DataTable dt = new System.Data.DataTable();
        private string[] titles;
        public FrmMain()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }
        private void InitRefresh()
        {
            dt.Clear();
            dt.Columns.Clear();
            flowLayoutPanel.Controls.Clear();
            cmbFileName.Items.Clear();
            richTextBox.Clear();
        }
        private void ExcelOperation()
        {
            InitRefresh();

            using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);

                dt = dataSet.Tables[0];

                titles = new string[dt.Columns.Count];
                typeArray = new string[dt.Columns.Count];

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    titles[i] = dt.Columns[i].ColumnName;
                    cmbFileName.Items.Add(dt.Columns[i].ColumnName);

                    MyTextEdit myText = new MyTextEdit();
                    myText.txtKey.Text = dt.Columns[i].ColumnName;
                    myText.txtValue.Text = dt.Columns[i].ColumnName;
                    myText.cmbTip.SelectedIndex = 0;
                    flowLayoutPanel.Controls.Add(myText);

                }

                cmbFileName.SelectedIndex = 0;
            }
        }
        private void LoadFile(AppEnums fileType)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();

                ofd.Filter = fileType == AppEnums.Excel ? @"Excel Dosyası|*.xlsx" : @"Word Dosyası|*.dotx";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    if (fileType == AppEnums.Excel)
                    {
                        ExcelFilePath = ofd.FileName;
                        ExcelOperation();
                        btnWordOpen.Enabled = true;
                    }
                    else if (fileType == AppEnums.Word)
                    {
                        WordFilePath = ofd.FileName;
                        btnOlustur.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Hata!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void FillWordParams()
        {
            WordParams.Clear();
            int i = 0;
            foreach (var control in flowLayoutPanel.Controls)
            {
                var cntrl = control as MyTextEdit;

                if (string.IsNullOrEmpty(cntrl.txtValue.Text)) continue;

                WordParams.Add(cntrl.txtKey.Text, cntrl.txtValue.Text);
                typeArray[i] = cntrl.cmbTip.Text;
                i++;
            }
        }
        private void OpenOutputFolder()
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    OutputFolder = fbd.SelectedPath;
                }
                else
                {
                    return;
                }
            }
        }
        private void WordOperation()
        {
            btnOlustur.Enabled = false;
            btnLoadFile.Enabled = false;

            string fileIndex = cmbFileName.Text;

            string templatePath = WordFilePath;

            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = templatePath;

            string regexPattern = @"[\\/:*?<>|]";
            int xx = 0;

            foreach (System.Data.DataRow data in dt.Rows)
            {
                int percentage = (xx + 1) * 100 / dt.Rows.Count;
                xx++;

                string fileName = String.Empty;

                Word.Application wordApp = new Word.Application();
                Document wordDoc = new Document();
                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                int y = 0;

                foreach (Field myMergeField in wordDoc.Fields)
                {
                    Range rngFieldCode = myMergeField.Code;
                    String fieldText = rngFieldCode.Text;

                    Int32 endMerge = fieldText.IndexOf("\\");
                    Int32 fieldNameLength = fieldText.Length - endMerge;
                    String fieldName = fieldText.Substring(11, endMerge - 11);
                    fieldName = fieldName.Trim();

                    string deger = WordParams.FirstOrDefault(x => x.Value == fieldName).Key;

                    if (deger != null)
                    {

                        int indis = Array.IndexOf(titles, deger);
                        string tip = typeArray[indis];

                        string jData = data[indis].ToString();

                        if (deger == fileIndex)
                        {
                            fileName = Regex.Replace(jData, regexPattern, "");
                        }

                        //if (tip == "Ondalık")
                        //{
                        //    jData = Convert.ToDouble(jData).ToString("N2");
                        //}

                        myMergeField.Select();
                        wordApp.Selection.TypeText(jData);
                    }
                }

                wordDoc.SaveAs($@"{OutputFolder}\{fileName}.docx");
                richTextBox.AppendText($" {fileName}.docx\n");

                if (chkPdf.Checked)
                {
                    wordDoc.ExportAsFixedFormat($@"{OutputFolder}\{fileName}.pdf", WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                      WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true,
                     WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, ref oMissing);
                    richTextBox.AppendText($" {fileName}.pdf\n");
                }

                wordApp.Application.Quit();
                
                backgroundWorker.ReportProgress(percentage);
            }
        }
        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            LoadFile(AppEnums.Excel);
        }
        private void btnWordOpen_Click(object sender, EventArgs e)
        {
            LoadFile(AppEnums.Word);
        }
        private void btnOlustur_Click(object sender, EventArgs e)
        {
            FillWordParams();

            if (WordParams.Count == 0)
            {
                MessageBox.Show("En az bir alan doldurmalısınız", "Hata!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (OutputFolder == String.Empty)
            {
                OpenOutputFolder();
            }

            backgroundWorker.RunWorkerAsync();
        }
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                WordOperation();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            btnLoadFile.Enabled = true;
            MessageBox.Show("Dosyalar oluşturuldu", "Oluşturma İşlemi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            btnOlustur.Enabled = true;
        }
        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }
        #region UnusedMethods
        private void RunOperation()
        {
            jsonFileContent = File.ReadAllText($@"D:\Users\bim03\Documents\Repos\WordApp\WordApp\bin\Debug\data\data.json");

            List<JObject> jsonDataList = JsonConvert.DeserializeObject<List<JObject>>(jsonFileContent);

            JObject firstObject = jsonDataList[0];
            string[] keys = firstObject.Properties().Select(p => p.Name).ToArray();

            foreach (var key in keys)
            {
                MyTextEdit myText = new MyTextEdit();
                myText.txtKey.Text = key;
                flowLayoutPanel.Controls.Add(myText);
                cmbFileName.Items.Add(key);
            }

            cmbFileName.SelectedIndex = 0;

        }
        #endregion
    }
}
