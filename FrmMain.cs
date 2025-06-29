﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAppGUI.UserControls;
using Excel = Microsoft.Office.Interop.Excel;
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
        private void FindAndReplace(Word.Application wordApp, object findText, object replaceText)
        {
            wordApp.Selection.Find.Execute(ref findText, true, true, false, false, false, true, false, 1,
                ref replaceText, 2, false, false, false, false);
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

                cmbFileName2.Items.Add("-");

                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    titles[i] = dt.Columns[i].ColumnName;
                    cmbFileName.Items.Add(dt.Columns[i].ColumnName);
                    cmbFileName2.Items.Add(dt.Columns[i].ColumnName);

                    MyTextEdit myText = new MyTextEdit();
                    myText.txtKey.Text = dt.Columns[i].ColumnName;
                    myText.txtValue.Text = "@" + dt.Columns[i].ColumnName;
                    myText.cmbTip.SelectedIndex = 0;
                    flowLayoutPanel.Controls.Add(myText);

                }

                cmbFileName.SelectedIndex = 0;
                cmbFileName2.SelectedIndex = 0;
            }
        }
        private void LoadFile(AppEnums fileType)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();

                ofd.Filter = fileType == AppEnums.Excel ? @"Excel Dosyası|*.xlsx" : rbWord.Checked ? @"Word Dosyası|*.docx" : @"Excel Dosyası|*.xlsx"; //dotx

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
            var i = 0;
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
        private void ExcelWriteOperation()
        {
            var fileIndex = cmbFileName.Text;
            var file2Index = cmbFileName2.Text;
            var dosyaYolu = WordFilePath;
            var regexPattern = @"[\\/:*?<>|]";
            Excel.Application excelApp = new Excel.Application();

            try
            {
                for (var xx = 0; xx < dt.Rows.Count; xx++)
                {
                    DataRow data = dt.Rows[xx];
                    var fileName = String.Empty;
                    Workbook workbook = excelApp.Workbooks.Open(dosyaYolu);
                    Worksheet worksheet = (Worksheet)workbook.Sheets[1];

                    xx++;

                    foreach (var wp in WordParams)
                    {
                        var indis = Array.IndexOf(titles, wp.Key);
                        var value = data[indis].ToString();

                        worksheet.Range[wp.Value].Value = value;

                        var jData = data[indis].ToString();

                        var deger = WordParams.FirstOrDefault(x => x.Key == wp.Key).Key;

                        if (deger == fileIndex)
                        {
                            fileName = Regex.Replace(jData, regexPattern, "");
                        }
                        else if (deger == file2Index)
                        {
                            var fn = Regex.Replace(jData, regexPattern, "");

                            fileName = $"{fileName} - {fn}";
                        }

                    }

                    workbook.SaveAs($@"{OutputFolder}\{fileName}.xlsx");
                    if (chkPdf.Checked)
                    {
                        CreatePdfFile(workbook, fileName, null, AppEnums.Excel);
                    }
                    workbook.Close(true);
                    Marshal.ReleaseComObject(workbook);

                    var percentage = (xx + 1) * 100 / dt.Rows.Count;
                    backgroundWorker.ReportProgress(percentage);
                }
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }        
        private void WordOperation()
        {
            Word.Application wordApp;
            btnOlustur.Enabled = false;
            btnLoadFile.Enabled = false;

            var fileIndex = cmbFileName.Text;
            var file2Index = cmbFileName2.Text;

            var templatePath = WordFilePath;

            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = templatePath;

            var regexPattern = @"[\\/:*?<>|]";
            var xx = 0;

            foreach (DataRow data in dt.Rows)
            {
                var percentage = (xx + 1) * 100 / dt.Rows.Count;
                xx++;

                var fileName = String.Empty;

                wordApp = new Word.Application();
                Document wordDoc = new Document();
                wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                foreach (var wp in WordParams)
                {
                    var indis = Array.IndexOf(titles, wp.Key);
                    var value = data[indis].ToString();

                    var pattern = @"\d{2}\.\d{2}\.\d{4}";

                    Match match = Regex.Match(value, pattern);

                    if (match.Success) 
                    {
                        value = match.Value;
                    }

                    FindAndReplace(wordApp, wp.Value, value);

                    var jData = data[indis].ToString();

                    var deger = WordParams.FirstOrDefault(x => x.Key == wp.Key).Key;

                    if (deger == fileIndex)
                    {
                        fileName = Regex.Replace(jData, regexPattern, "");
                    }
                    else if (deger == file2Index)
                    {
                        var fn = Regex.Replace(jData, regexPattern, "");

                        fileName = $"{fileName} - {fn}";
                    }
                }


                wordDoc.SaveAs($@"{OutputFolder}\{fileName}.docx");
                AppendTextSafe($" {fileName}.docx\n");

                if (chkPdf.Checked)
                {
                    CreatePdfFile(wordDoc, fileName, oMissing, AppEnums.Word);
                    AppendTextSafe($" {fileName}.pdf\n");
                }

                wordApp.Application.Quit();
                
                backgroundWorker.ReportProgress(percentage);
            }
        }
        private void CreatePdfFile(object document, string fileName, Object oMissing, AppEnums type)
        {
            if (type == AppEnums.Word)
            {
                ((Document)document).ExportAsFixedFormat($@"{OutputFolder}\{fileName}.pdf", WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                         WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true,
                        WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, ref oMissing);
            }
            else if (type == AppEnums.Excel)
            {
                ((Workbook)document).ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, $@"{OutputFolder}\{fileName}.pdf");
            }
               
        }
        private void AppendTextSafe(string text)
        {
            if (richTextBox.InvokeRequired)
            {
                richTextBox.Invoke(new System.Action(() => richTextBox.AppendText(text)));
            }
            else
            {
                richTextBox.AppendText(text);
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
                if(rbExcel.Checked)
                    ExcelWriteOperation();
                else if(rbWord.Checked)
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
        private void ExcelWriteOperation_X()
        {
            string fileIndex = cmbFileName.Text;
            string file2Index = cmbFileName2.Text;
            string dosyaYolu = WordFilePath;
            string regexPattern = @"[\\/:*?<>|]";
            int xx = 0;
            foreach (DataRow data in dt.Rows)
            {
                string fileName = String.Empty;

                Excel.Application excelApp = new Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(dosyaYolu);
                Worksheet worksheet = (Worksheet)workbook.Sheets[1];

                int percentage = (xx + 1) * 100 / dt.Rows.Count;
                xx++;

                foreach (var wp in WordParams)
                {
                    int indis = Array.IndexOf(titles, wp.Key);
                    string value = data[indis].ToString();

                    worksheet.Range[wp.Value].Value = value;

                    string jData = data[indis].ToString();

                    string deger = WordParams.FirstOrDefault(x => x.Key == wp.Key).Key;

                    if (deger == fileIndex)
                    {
                        fileName = Regex.Replace(jData, regexPattern, "");
                    }
                    else if (deger == file2Index)
                    {
                        string fn = Regex.Replace(jData, regexPattern, "");

                        fileName = $"{fileName} - {fn}";
                    }

                }
                workbook.SaveAs($@"{OutputFolder}\{fileName}.xlsx");

                if (chkPdf.Checked)
                {
                    workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, $@"{OutputFolder}\{fileName}.pdf");
                }

                workbook.Close(true);
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                backgroundWorker.ReportProgress(percentage);
            }

        }
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

        private void linkLabel1_Click(object sender, EventArgs e)
        {
            FrmYardim yardim = new FrmYardim();
            yardim.ShowDialog();
        }
    }
}
