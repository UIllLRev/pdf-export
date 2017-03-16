using System;
using System.Windows.Forms;

using Microsoft.Office.Interop.Word;

namespace PDFArticleExport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".docx";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                // TODO: make more robust
                DocxToPdf(dlg.FileName, dlg.FileName.Replace(".docx", ".pdf"));
            }
        }
            private void DocxToPdf(String sourcePath, String destPath)
        {

            //Change the path of the .docx file and filename to your file name.

            object paramSourceDocPath = sourcePath;
            object paramMissing = Type.Missing;
            toolStripStatusLabel1.Text = "Starting Word...";
            var wordApplication = new Microsoft.Office.Interop.Word.Application();
            Document wordDocument = null;

            //Change the path of the .pdf file and filename to your file name.

            string paramExportFilePath = destPath;
            WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            WdExportOptimizeFor paramExportOptimizeFor =
                WdExportOptimizeFor.wdExportOptimizeForPrint;
            WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            WdExportCreateBookmarks paramCreateBookmarks =
                WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            try
            {
                toolStripStatusLabel1.Text = "Opening Word document...";
                // Open the source document.
                wordDocument = wordApplication.Documents.Open(
                    ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing);

                // Export it in the specified format.
                if (wordDocument != null)
                {
                    toolStripStatusLabel1.Text = "Hiding headers...";

                    wordDocument.Styles["_header"].Font.Color = WdColor.wdColorWhite;

                    foreach (HeaderFooter hf in wordDocument.Sections[1].Headers)
                    {
                        if (hf.IsHeader)
                        {
                            foreach (Range r in hf.Range.Sentences)
                            {
                                if (r.Text.IndexOf("Do Not Delete") >= 0)
                                {
                                    r.Select();
                                    wordApplication.Selection.ClearCharacterDirectFormatting();
                                }
                            }
                        }
                    }

                    toolStripStatusLabel1.Text = "Exporting to PDF...";
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                        paramExportFormat, paramOpenAfterExport,
                        paramExportOptimizeFor, paramExportRange, paramStartPage,
                        paramEndPage, paramExportItem, paramIncludeDocProps,
                        paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                        paramBitmapMissingFonts, paramUseISO19005_1,
                        ref paramMissing);

                    toolStripStatusLabel1.Text = "Done!";
                }
            }
            catch (Exception ex)
            {
                // Respond to the error
                System.Windows.Forms.MessageBox.Show(ex.Message);
                toolStripStatusLabel1.Text = "An error occurred.";
            }
            finally
            {
                // Close and release the Document object.
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing,
                        ref paramMissing);
                    wordDocument = null;
                }

                // Quit Word and release the ApplicationClass object.
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing,
                        ref paramMissing);
                    wordApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
