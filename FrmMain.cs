using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using WordAppGUI.UserControls;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WordAppGUI
{
    public partial class FrmMain : Form
    {
        private string ExcelFilePath = "";
        private string WordFilePath = "";
        private readonly Dictionary<string, string> WordParams;
        private string[] typeArray = null;
        private string OutputFolder = string.Empty;
        private System.Data.DataTable dt = new System.Data.DataTable();
        private string[] titles;
        public FrmMain()
        {
            InitializeComponent();
            WordParams = new Dictionary<string, string>();
            CheckForIllegalCrossThreadCalls = false;
            backgroundWorker.WorkerSupportsCancellation = true;
        }
        private void InitRefresh()
        {
            dt.Clear();
            dt.Columns.Clear();
            flowLayoutPanel.Controls.Clear();
            cmbFileName.Items.Clear();
            richTextBox.Clear();
            btnOlustur.Enabled = false;
            btnDurdur.Enabled = false;
            chkPdf.Checked = false;
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
                var reader = ExcelReaderFactory.CreateReader(stream);

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

                    var myText = new MyTextEdit();
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
                var ofd = new OpenFileDialog();

                var filter = "";

                switch (fileType)
                {
                    case AppEnums.Excel:
                        filter = MyConstants.EXCEL_FILTER;
                        break;

                    case AppEnums.Word:
                        filter = MyConstants.WORD_FILTER;
                        break;

                    default:
                        filter = "Tüm Dosyalar|*.*";
                        break;
                }


                ofd.Filter = filter;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    if (!File.Exists(ofd.FileName))
                    {
                        MyMessages.ErrorMessage("Seçilen dosya bulunamadı");
                        return;
                    }

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
            catch (UnauthorizedAccessException)
            {
                MyMessages.ErrorMessage("Dosyaya erişim izni yok.", "Yetki Hatası");
            }
            catch (FileNotFoundException)
            {
                MyMessages.ErrorMessage("Dosya bulunamadı.", "Dosya Hatası");
            }
            catch (Exception ex)
            {
                MyMessages.ErrorMessage(ex.Message);
            }
        }
        private void FillWordParams()
        {
            WordParams.Clear();
            var i = 0;
            foreach (var control in flowLayoutPanel.Controls)
            {
                var cntrl = control as MyTextEdit;

                if (string.IsNullOrEmpty(cntrl.txtValue.Text))
                {
                    continue;
                }

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
        private void ExcelWriteOperation(DoWorkEventArgs e)
        {
            var fileIndex = cmbFileName.Text;
            var file2Index = cmbFileName2.Text;
            var dosyaYolu = ExcelFilePath;
            var regexPattern = MyConstants.FILENAME_INVALID_CHARS;
            Excel.Application excelApp = null;

            try
            {
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    if (backgroundWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }
                    Workbook workbook = null;
                    Worksheet worksheet = null;
                    try
                    {
                        var data = dt.Rows[i];
                        var fileName = string.Empty;
                        workbook = excelApp.Workbooks.Open(dosyaYolu);
                        worksheet = (Worksheet)workbook.Sheets[1];

                        foreach (var wp in WordParams)
                        {
                            var indis = Array.IndexOf(titles, wp.Key);
                            var rawValue = data[indis];
                            string value;

                            if (rawValue is DateTime dateTimeValue)
                            {
                                value = dateTimeValue.ToString(MyConstants.DATE_FORMAT);
                            }
                            else
                            {
                                value = rawValue.ToString();
                            }

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

                        workbook.SaveAs($@"{OutputFolder}\\{fileName}.xlsx");
                        if (chkPdf.Checked)
                        {
                            CreatePdfFile(workbook, fileName, null, AppEnums.Excel);
                        }

                        var percentage = (i + 1) * 100 / dt.Rows.Count;
                        backgroundWorker.ReportProgress(percentage);
                    }
                    catch (Exception ex)
                    {
                        AppendTextSafe($"Hata (Satır {i + 1}): {ex.Message}\n");
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        if (worksheet != null)
                        {
                            Marshal.ReleaseComObject(worksheet);
                        }
                    }
                }
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        private void SetUIEnabled(bool enabled)
        {
            btnOlustur.Enabled = enabled;
            btnLoadFile.Enabled = enabled;
        }
        private void WordOperation(DoWorkEventArgs e)
        {
            Word.Application wordApp = null;
            try
            {
                SetUIEnabled(false);

                var fileIndex = cmbFileName.Text;
                var file2Index = cmbFileName2.Text;
                var templatePath = WordFilePath;

                object oMissing = System.Reflection.Missing.Value;
                object oTemplatePath = templatePath;

                var regexPattern = MyConstants.FILENAME_INVALID_CHARS;

                wordApp = new Word.Application
                {
                    Visible = false
                };

                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    if (backgroundWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }

                    Document wordDoc = null;
                    try
                    {
                        var data = dt.Rows[i];
                        var percentage = (i + 1) * 100 / dt.Rows.Count;
                        var fileName = string.Empty;

                        wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                        foreach (var wp in WordParams)
                        {
                            var indis = Array.IndexOf(titles, wp.Key);
                            var value = data[indis].ToString();

                            var pattern = MyConstants.DATE_PATTERN;
                            var match = Regex.Match(value, pattern);

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

                        backgroundWorker.ReportProgress(percentage);
                    }
                    catch (Exception ex)
                    {
                        AppendTextSafe($"Hata (Satır {i + 1}): {ex.Message}\n");
                    }
                    finally
                    {
                        if (wordDoc != null)
                        {
                            wordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                            Marshal.ReleaseComObject(wordDoc);
                        }
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }
        private void CreatePdfFile(object document, string fileName, object oMissing, AppEnums type)
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
        private void BtnLoadFile_Click(object sender, EventArgs e)
        {
            LoadFile(AppEnums.Excel);
        }
        private void BtnWordOpen_Click(object sender, EventArgs e)
        {
            LoadFile(AppEnums.Word);
        }
        private void BtnOlustur_Click(object sender, EventArgs e)
        {
            FillWordParams();

            if (WordParams.Count == 0)
            {
                MyMessages.WarningMessage("En az bir alan doldurmalısınız", "Hata");
                return;
            }

            if (OutputFolder == string.Empty)
            {
                OpenOutputFolder();
            }

            btnOlustur.Enabled = false;
            btnDurdur.Enabled = true;


            backgroundWorker.RunWorkerAsync();
        }
        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            btnDurdur.Enabled = true;
            try
            {
                if (rbExcel.Checked)
                {
                    ExcelWriteOperation(e);
                }
                else if (rbWord.Checked)
                {
                    WordOperation(e);
                }
            }
            catch (Exception exception)
            {
                MyMessages.ErrorMessage(exception.Message);
            }
        }
        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }
        private void LinkLabel1_Click(object sender, EventArgs e)
        {
            var yardim = new FrmYardim();
            yardim.ShowDialog();
        }
        private void BtnDurdur_Click(object sender, EventArgs e)
        {
            if (backgroundWorker.IsBusy)
            {
                backgroundWorker.CancelAsync();
            }
        }
        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MyMessages.WarningMessage("İşlem kullanıcı tarafından iptal edildi", "İptal");
            }
            else if (e.Error != null)
            {
                MyMessages.ErrorMessage("Bir hata oluştu: " + e.Error.Message);
            }
            else
            {
                MyMessages.InformationMessage("Dosyalar başarıyla oluşturuldu", "İşlem Tamamlandı");
            }

            btnOlustur.Enabled = true;
            btnDurdur.Enabled = false;
            btnLoadFile.Enabled = true;
        }
        private void FrmMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.N)
            {
                InitRefresh();
            }
        }
    }
}
