using iText.Forms;
using iText.Forms.Fields;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using iText.Layout.Element;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMHelper.PDF
{
    public static class PDFHelper
    {
        public static bool MergeAll(string strFilePathOutput, List<string> strFilePathInput, out string strErr)
        {
            strErr = string.Empty;
            bool blSuccess = true;
            FileInfo fileInfo = null;
            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(strFilePathOutput));
            try
            {
                pdfDoc.InitializeOutlines();
                foreach (string strInput in strFilePathInput)
                {
                    fileInfo = new FileInfo(strInput);
                    PdfReader pdfReader = new PdfReader(strInput);
                    pdfReader.SetUnethicalReading(true);
                    PdfDocument readerDoc = new PdfDocument(pdfReader);
                    readerDoc.CopyPagesTo(1, readerDoc.GetNumberOfPages(), pdfDoc, new PdfPageFormCopier());
                    readerDoc.Close();
                }
                pdfDoc.Close();
            }
            catch (Exception ex)
            {
                if (pdfDoc.HasOutlines())
                    pdfDoc.Close();
                if (File.Exists(strFilePathOutput))
                    File.Delete(strFilePathOutput);
                blSuccess = false;
                strErr = string.Format("Error - {0}{1}", ex.Message, (fileInfo == null) ? string.Empty : string.Format(" File - {0}", fileInfo.Name));
            }

            return blSuccess;
        }

        public static bool CanRead(string filePath, out string strErr)
        {
            strErr = string.Empty;
            bool blSuccess = true;
            FileInfo fileInfo = null;
            PdfReader pdfReader = null;
            PdfDocument readerDoc = null;
            try
            {

                fileInfo = new FileInfo(filePath);
                pdfReader = new PdfReader(filePath);
                pdfReader.SetUnethicalReading(true);
                readerDoc = new PdfDocument(pdfReader);
                readerDoc.Close();
            }
            catch (Exception ex)
            {
                blSuccess = false;
                strErr = string.Format("Error - {0}{1}", ex.Message, (fileInfo == null) ? string.Empty : string.Format(" File - {0}", fileInfo.Name));
            }
            finally
            {
                if (readerDoc != null)
                    readerDoc.Close();

                if (pdfReader != null)
                    pdfReader.Close();

            }

            return blSuccess;
        }

        public static bool ChangePageSize(string strFilePathOutput, string strFilePathInput, out string strErr)
        {
            bool blnIsSuccess = true;
            strErr = null;

            try
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(strFilePathInput), new PdfWriter(strFilePathOutput));
                float margin = 1.429F;
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    PdfPage page = pdfDoc.GetPage(i);
                    // change page size
                    Rectangle mediaBox = page.GetMediaBox();
                    Rectangle newMediaBox = new Rectangle(
                            mediaBox.GetLeft() - margin, mediaBox.GetBottom() - margin,
                            mediaBox.GetWidth() + margin * 2, mediaBox.GetHeight() + margin * 2);
                    page.SetMediaBox(newMediaBox);
                    // add border
                    PdfCanvas over = new PdfCanvas(page);
                    over.SetFillColor(new DeviceGray());
                    over.Rectangle(mediaBox.GetLeft(), mediaBox.GetBottom(),
                            mediaBox.GetWidth(), mediaBox.GetHeight());
                    over.Stroke();
                    // change rotation of the even pages
                    if (i % 2 == 0)
                    {
                        page.SetRotation(180);
                    }
                }
                pdfDoc.Close();
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strErr = ex.Message;
            }

            return blnIsSuccess;
        }

        public static bool ExtractTextPDFForm(string strPDFPath, out Dictionary<string, string> dicFormFieldNmValPair, out string strExceptionMessage)
        {
            bool blnIsSuccess = false;
            PdfReader pdfRead = null;
            PdfDocument pdfDoc = null;

            dicFormFieldNmValPair = null;
            strExceptionMessage = null;

            try
            {
                ValidatePDFPath(strPDFPath);
                pdfRead = new PdfReader(strPDFPath);
                pdfDoc = new PdfDocument(pdfRead);

                PdfAcroForm pdfAcroForm = PdfAcroForm.GetAcroForm(pdfDoc, false);
                IDictionary<string, PdfFormField> dicPDFFormVal = pdfAcroForm.GetFormFields();

                if (dicPDFFormVal != null && dicPDFFormVal.Count > 0)
                {
                    dicFormFieldNmValPair = new Dictionary<string, string>();

                    foreach (string strFieldName in dicPDFFormVal.Keys)
                    {
                        if (!string.IsNullOrWhiteSpace(strFieldName))
                            dicFormFieldNmValPair.Add(strFieldName, dicPDFFormVal[strFieldName].GetValueAsString());
                    }
                }

                blnIsSuccess = true;
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }
            finally
            {
                ClosePDF(pdfRead, pdfDoc);
            }

            return blnIsSuccess;
        }

        public static bool ExtractTextPDFDigital(string strPDFPath, out Dictionary<int, string> dicPageNoTextPair, out string strExceptionMessage)
        {
            bool blnIsSuccess = false;
            PdfReader pdfRead = null;
            PdfDocument pdfDoc = null;

            strExceptionMessage = null;
            dicPageNoTextPair = null;

            try
            {
                ValidatePDFPath(strPDFPath);
                pdfRead = new PdfReader(strPDFPath);
                pdfDoc = new PdfDocument(pdfRead);

                int iPDFNoOfPage = pdfDoc.GetNumberOfPages();
                dicPageNoTextPair = new Dictionary<int, string>();

                for (int iPageIndex = 1; iPageIndex <= iPDFNoOfPage; iPageIndex++)
                {
                    PdfPage pdfPage = pdfDoc.GetPage(iPageIndex);

                    string strPageText = PdfTextExtractor.GetTextFromPage(pdfPage, new LocationTextExtractionStrategy());
                    string strPageTextEncoded = !string.IsNullOrWhiteSpace(strPageText) ? Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(strPageText))) : string.Empty;

                    dicPageNoTextPair.Add(iPageIndex, strPageTextEncoded);
                }

                pdfDoc.Close();
                pdfRead.Close();

                blnIsSuccess = true;
            }
            
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }
            finally
            {
                ClosePDF(pdfRead, pdfDoc);
            }

            return blnIsSuccess;
        }

        public static bool ExtractUrlFromPDF(string strPDFPath, out Dictionary<int, List<string>> dicPageNoUrlPair, out string strExceptionMessage)
        {
            bool blnIsSuccess = false;
            PdfReader pdfRead = null;
            PdfDocument pdfDoc = null;

            dicPageNoUrlPair = null;
            strExceptionMessage = null;

            try
            {
                ValidatePDFPath(strPDFPath);
                pdfRead = new PdfReader(strPDFPath);
                pdfDoc = new PdfDocument(pdfRead);

                int iPDFNoOfPage = pdfDoc.GetNumberOfPages();
                dicPageNoUrlPair = new Dictionary<int, List<string>>();

                for (int iPageIndex = 1; iPageIndex <= iPDFNoOfPage; iPageIndex++)
                {
                    List<string> UrlList = null;
                    PdfPage pdfPage = pdfDoc.GetPage(iPageIndex);
                    IList<PdfAnnotation> pdfAnnotList = pdfPage.GetAnnotations();
                    
                    if (pdfAnnotList != null && pdfAnnotList.Count > 0)
                    {
                        foreach (var annot in pdfAnnotList)
                        {
                            if (annot != null && annot.GetSubtype().Equals(PdfName.Link))
                            {
                                PdfDictionary annotAction = ((PdfLinkAnnotation)annot).GetAction();

                                if (annotAction != null && (annotAction.Get(PdfName.S).Equals(PdfName.URI) || annotAction.Get(PdfName.S).Equals(PdfName.GoToR)))
                                {
                                    PdfString pdfUrl = annotAction.GetAsString(PdfName.URI);
                                    if (pdfUrl != null)
                                    {
                                        string strPdfUrl = pdfUrl.GetValue();
                                        if (!string.IsNullOrWhiteSpace(strPdfUrl))
                                        {
                                            UrlList = UrlList ?? new List<string>();
                                            UrlList.Add(strPdfUrl);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    dicPageNoUrlPair.Add(iPageIndex, UrlList);
                }

                pdfDoc.Close();
                pdfRead.Close();

                blnIsSuccess = true;
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }
            finally
            {
                ClosePDF(pdfRead, pdfDoc);
            }

            return blnIsSuccess;
        }

        public static bool ConvertImageToPDF(string strPDFOutputPath, List<string> lstImgFilePath, out string strExceptionMessage)
        {
            bool blnIsSuccess = true;
            PdfDocument pdfDoc = null;
            Document doc = null;
            strExceptionMessage = null;

            try
            {
                #region Pre-checking
                if (!(lstImgFilePath != null && lstImgFilePath.Count > 0))
                    throw new Exception("No image path to convert to PDF");
                if (string.IsNullOrWhiteSpace(strPDFOutputPath))
                    throw new Exception("There is no PDF output path specified");
                #endregion

                pdfDoc = new PdfDocument(new PdfWriter(strPDFOutputPath));
                doc = new Document(pdfDoc);
                pdfDoc.InitializeOutlines();
                int iPageNo = pdfDoc.GetNumberOfPages();

                foreach (string strImgFilePath in lstImgFilePath)
                {
                    if (iPageNo < 1)
                        iPageNo = 1;
                    if (!File.Exists(strImgFilePath))
                        continue;

                    Image img = new Image(ImageDataFactory.Create(strImgFilePath));
                    img.ScaleToFit(PageSize.Default.GetWidth(), PageSize.Default.GetHeight());
                    img.SetFixedPosition(iPageNo, PageSize.Default.GetWidth() - img.GetImageScaledWidth(), PageSize.Default.GetHeight() - img.GetImageScaledHeight());

                    doc.Add(img);
                    
                    iPageNo++;
                }
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }
            finally
            {
                if (pdfDoc != null && pdfDoc.HasOutlines())
                    pdfDoc.Close();
            }

            return blnIsSuccess;
        }

        #region Helper
        private static void ClosePDF(PdfReader pdfReader, PdfDocument pdfDoc)
        {
            try
            {
                if (pdfDoc != null)
                    pdfDoc.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void ValidatePDFPath(string strPDFFilePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(strPDFFilePath))
                    throw new Exception("PDF path is null or empty");
                if (!File.Exists(strPDFFilePath))
                    throw new Exception($"PDF file {strPDFFilePath} is not exists");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
