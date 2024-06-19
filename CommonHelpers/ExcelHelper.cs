using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Diagnostics;

namespace TMHelper.Common
{
    public class ExcelHelper
    {
        #region Properties
        public const string CONST_INCORRECT_EXCEL_FILE_PASSWORD_ERR_MSG = "The password you supplied is not correct";

        public string strFilePath { get; set; }
        public string strFileNm { get; set; }
        public List<string> lstColumnNm { get; set; }
        public System.Data.DataTable dtSource { get; set; }

        #region Microsoft.Office.Interop.Excel
        private object missing = Missing.Value;
        private Microsoft.Office.Interop.Excel.Application ExcelApp;
        private Microsoft.Office.Interop.Excel.Workbook ExcelBook;
        private Microsoft.Office.Interop.Excel.Worksheet ExcelSheet;
        private Microsoft.Office.Interop.Excel.Range range1 = null;
        #endregion

        #endregion

        #region Constructor
        public ExcelHelper()
        {

        }
        #endregion

        #region Accessible Method
        public bool DataTableToExcel(bool blAppend, out string strErr, bool blColumnNm)
        {
            bool blSuccess = false;
            strErr = string.Empty;
            int intStartRow = 1;
            int intEndRow = 1;
            int intStartColumn = 1;
            int intEndColumn = 1;

            try
            {
                if (this.dtSource == null)
                    throw new Exception("Source data table cannot be null. ");

                if (string.IsNullOrWhiteSpace(this.strFilePath))
                    throw new Exception("Output file path cannot be null or empty. ");

                if (string.IsNullOrWhiteSpace(this.strFileNm))
                    throw new Exception("Output file name cannot be null or empty. ");

                if (!(this.strFileNm.ToUpper().EndsWith(".XLS") || this.strFileNm.ToUpper().EndsWith(".XLSX")))
                    throw new Exception(string.Format("Output file format {0} is not applicable. ", this.strFileNm));

                string strFilePathFull = Path.Combine(this.strFilePath, this.strFileNm); //use Path.Combine instead to eliminate the consideration of including \ or not
                    //string.Format("{0}{1}", this.strFilePath, this.strFileNm);

                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Visible = false;
                if (blAppend && File.Exists(strFilePathFull))
                {
                    ExcelBook = ExcelApp.Workbooks.Open(strFilePathFull, Type.Missing, false);
                }
                else
                {
                    ExcelBook = ExcelApp.Workbooks.Add(missing);
                }

                ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Sheets.get_Item(1);

                // Get row number to start insert record
                if (string.IsNullOrWhiteSpace(ExcelSheet.Cells[1, 1].Value))
                {
                    intStartRow = 1;
                }
                else
                {
                    intStartRow = ((Range)ExcelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing)).Row;
                }

                intEndRow = intStartRow + dtSource.Rows.Count;
                intEndColumn = dtSource.Columns.Count;

                range1 = ExcelSheet.Range[ExcelSheet.Cells[intStartRow, intStartColumn], ExcelSheet.Cells[intEndRow, intEndColumn]];
                //Header                
                if (intStartRow == 1)
                {
                    for (int i = 0; i < dtSource.Columns.Count; i++)
                    {
                        range1.Cells[intStartRow, i + 1] = dtSource.Columns[i].ColumnName;
                    }
                }

                //Datas
                //for (int i = 0; i < dtSource.Rows.Count; i++)
                //{
                //    for (int j = 0; j < dtSource.Columns.Count; j++)
                //    {
                //        range1.Cells[i + 2, j + 1] = dtSource.Rows[i][j];
                //    }
                //}

                var data = new object[dtSource.Rows.Count, dtSource.Columns.Count];
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                        data[i, j] = dtSource.Rows[i][j];
                    }
                }

                var startCell = (Range)ExcelSheet.Cells[intStartRow + 1, intStartColumn];
                var endCell = (Range)ExcelSheet.Cells[intEndRow, intEndColumn];
                var writeRange = ExcelSheet.Range[startCell, endCell];
                writeRange.NumberFormat = "@";
                writeRange.Value2 = data;

                if (intStartRow == 1)
                {
                    ExcelBook.SaveAs(strFilePathFull, GetFileFormat((strFilePathFull.ToUpper().EndsWith(".XLSX") ? ExcelFileFormat.XLSX : ExcelFileFormat.XLS)),
                            Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlExclusive, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                }
                else
                {
                    ExcelBook.Save();
                }

                ExcelBook.Close();
                ExcelApp.Quit();
                blSuccess = true;
            }
            catch (Exception)
            {

                throw;
            }

            return blSuccess;
        }

        //overload method - with release memory after used, directory checking and colored header
        public bool DataTableToExcel(System.Data.DataTable dtSource, string strExcelOutputFolderPath, string strExcelFileNm, bool blAppend, out string strErr, bool blnIsChangeSheetNm = false, System.Data.DataTable dtblHeaderData = null)
        {
            bool blnIsSuccess = false;
            
            int intStartRow = 1;
            int intEndRow = 1;
            int intStartColumn = 1;
            int intEndColumn = 1;

            strErr = string.Empty;

            try
            {
                if (dtSource == null)
                    throw new Exception("Source data table cannot be null. ");

                if (string.IsNullOrWhiteSpace(strExcelOutputFolderPath))
                    throw new Exception("Output folder path cannot be null or empty. ");

                if (string.IsNullOrWhiteSpace(strExcelFileNm))
                    throw new Exception("Output file name cannot be null or empty. ");

                if (!Directory.Exists(strExcelOutputFolderPath))
                    throw new Exception("Output folder path not exists");

                if (!(strExcelFileNm.ToUpper().EndsWith(".XLS") || strExcelFileNm.ToUpper().EndsWith(".XLSX")))
                    throw new Exception(string.Format("Output file format {0} is not applicable. ", strExcelFileNm));

                string strFilePathFull = Path.Combine(strExcelOutputFolderPath, strExcelFileNm);
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Visible = false;
                if (blAppend && File.Exists(strFilePathFull))
                {
                    ExcelBook = ExcelApp.Workbooks.Open(strFilePathFull, Type.Missing, false);
                }
                else
                {
                    ExcelBook = ExcelApp.Workbooks.Add(missing);
                }

                ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Sheets.get_Item(1);

                // Get row number to start insert record
                if (string.IsNullOrWhiteSpace(ExcelSheet.Cells[1, 1].Value))
                {
                    intStartRow = 1;
                }
                else
                {
                    intStartRow = ((Range)ExcelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing)).Row;
                }

                if (dtblHeaderData != null && dtblHeaderData.Rows != null && dtblHeaderData.Rows.Count > 0)
                {
                    int iHeaderRow = 1;
                    foreach (DataRow drHeader in dtblHeaderData.Rows)
                    {
                        int iHeaderColumn = 1;
                        foreach (DataColumn dcHeader in dtblHeaderData.Columns)
                        {
                            ExcelSheet.Cells[iHeaderRow, iHeaderColumn] = FormatCellValueAsString(drHeader[dcHeader.ColumnName].ToString());
                            ExcelSheet.Cells[iHeaderRow, iHeaderColumn].Font.Bold = true;
                            iHeaderColumn++;
                        }

                        iHeaderRow++;
                    }

                    intStartRow = ((Range)ExcelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing)).Row + 1;
                }

                intEndRow = intStartRow + dtSource.Rows.Count;
                intEndColumn = dtSource.Columns.Count;

                range1 = ExcelSheet.Range[ExcelSheet.Cells[intStartRow, intStartColumn], ExcelSheet.Cells[intEndRow, intEndColumn]];

                //Header                
                if (intStartRow == 1 || (dtblHeaderData != null && dtblHeaderData.Rows != null && dtblHeaderData.Rows.Count + 1 == intStartRow))
                {
                    for (int i = 0; i < dtSource.Columns.Count; i++)
                    {
                        ExcelSheet.Cells[intStartRow, i + 1] = string.Format("'{0}", dtSource.Columns[i].ColumnName);
                        ExcelSheet.Cells[intStartRow, i + 1].Rows.Interior.Color = XlRgbColor.rgbLightGrey;
                    }
                }

                var data = new object[dtSource.Rows.Count, dtSource.Columns.Count];
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                        data[i, j] = string.Format("'{0}", dtSource.Rows[i][j]);
                    }
                }

                var startCell = (Range)ExcelSheet.Cells[intStartRow + 1, intStartColumn];
                var endCell = (Range)ExcelSheet.Cells[intEndRow, intEndColumn];
                var writeRange = ExcelSheet.Range[startCell, endCell];
                //writeRange.NumberFormat = "@";
                writeRange.Value2 = data;
                writeRange.Columns.AutoFit();

                if (blnIsChangeSheetNm)
                    ExcelSheet.Name = dtSource.TableName;

                if (intStartRow == 1 || (dtblHeaderData != null && dtblHeaderData.Rows != null && dtblHeaderData.Rows.Count + 1 == intStartRow))
                {
                    ExcelBook.SaveAs(strFilePathFull, GetFileFormat((strFilePathFull.ToUpper().EndsWith(".XLSX") ? ExcelFileFormat.XLSX : ExcelFileFormat.XLS)),
                            Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlExclusive, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                }
                else
                {
                    ExcelBook.Save();
                }

                ExcelBook.Close();
                blnIsSuccess = true;
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strErr = ex.Message;
            }
            finally
            {
                ReleaseExcelObject();
            }

            return blnIsSuccess;
        }

        public bool DataTableToExcel(System.Data.DataTable dtSource, string strExcelOutputFolderPath, string strExcelFileNm, string strWorksheetName, bool blAppend, bool clearContent, out string strErr, bool noChange = false)
        {
            bool blnIsSuccess = false;

            int intStartRow = 1;
            int intEndRow = 1;
            int intStartColumn = 1;
            int intEndColumn = 1;

            strErr = string.Empty;

            try
            {
                if (dtSource == null)
                    throw new Exception("Source data table cannot be null. ");

                if (string.IsNullOrWhiteSpace(strExcelOutputFolderPath))
                    throw new Exception("Output folder path cannot be null or empty. ");

                if (string.IsNullOrWhiteSpace(strExcelFileNm))
                    throw new Exception("Output file name cannot be null or empty. ");

                if (!Directory.Exists(strExcelOutputFolderPath))
                    throw new Exception("Output folder path not exists");

                if (!(strExcelFileNm.ToUpper().EndsWith(".XLS") || strExcelFileNm.ToUpper().EndsWith(".XLSX")))
                    throw new Exception(string.Format("Output file format {0} is not applicable. ", strExcelFileNm));

                string strFilePathFull = Path.Combine(strExcelOutputFolderPath, strExcelFileNm);
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Visible = false;
                ExcelApp.DisplayAlerts = !noChange;
                if (blAppend && File.Exists(strFilePathFull))
                {
                    ExcelBook = ExcelApp.Workbooks.Open(strFilePathFull, Type.Missing, false);
                    for(int i=1; i<=ExcelBook.Worksheets.Count; i++)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet tempExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Sheets.get_Item(i);
                        if (tempExcelSheet.Name.Equals(strWorksheetName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Sheets.get_Item(i);
                            break;
                        }
                    }
                    
                    if (ExcelSheet == null)
                        throw new Exception(string.Format("Cannot find worksheet {0}", strWorksheetName));
                }
                else
                {
                    ExcelBook = ExcelApp.Workbooks.Add(missing);
                    ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Sheets.get_Item(1);
                    ExcelSheet.Name = strWorksheetName;
                }

                if (clearContent)
                {
                    Microsoft.Office.Interop.Excel.Range rng = (Range)ExcelSheet.Cells[((Range)ExcelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing)).Row, 
                        ((Range)ExcelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing)).Column];
                    rng.Delete();
                }

                // Get row number to start insert record
                if (string.IsNullOrWhiteSpace(ExcelSheet.Cells[1, 1].Value))
                {
                    intStartRow = 1;
                }
                else
                {
                    intStartRow = ((Range)ExcelSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing)).Row;
                }

                intEndRow = intStartRow + dtSource.Rows.Count;
                intEndColumn = dtSource.Columns.Count;

                range1 = ExcelSheet.Range[ExcelSheet.Cells[intStartRow, intStartColumn], ExcelSheet.Cells[intEndRow, intEndColumn]];

                //Header                
                if (intStartRow == 1)
                {
                    for (int i = 0; i < dtSource.Columns.Count; i++)
                    {
                        range1.Cells[intStartRow, i + 1] = string.Format("'{0}", dtSource.Columns[i].ColumnName);
                        range1.Cells[intStartRow, i + 1].Rows.Interior.Color = XlRgbColor.rgbLightGrey;
                    }
                }

                var data = new object[dtSource.Rows.Count, dtSource.Columns.Count];
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                        data[i, j] = string.Format("'{0}", dtSource.Rows[i][j]);
                    }
                }

                var startCell = (Range)ExcelSheet.Cells[intStartRow + 1, intStartColumn];
                var endCell = (Range)ExcelSheet.Cells[intEndRow, intEndColumn];
                var writeRange = ExcelSheet.Range[startCell, endCell];
                //writeRange.NumberFormat = "@";
                writeRange.Value2 = data;
                writeRange.Columns.AutoFit();

                if (intStartRow == 1)
                {
                    ExcelBook.SaveAs(strFilePathFull, GetFileFormat((strFilePathFull.ToUpper().EndsWith(".XLSX") ? ExcelFileFormat.XLSX : ExcelFileFormat.XLS)),
                            Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlExclusive,
                            noChange ? Type.Missing : XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                }
                else
                {
                    ExcelBook.Save();
                }

                ExcelBook.Close();
                blnIsSuccess = true;
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strErr = ex.Message;
            }
            finally
            {
                ReleaseExcelObject();
            }

            return blnIsSuccess;
        }

        public bool DataSetToExcel(System.Data.DataSet dsSource, string strExcelOutputFolderPath, string strExcelFileNm, out string strErr, List<string> HyperlinkColNmList = null, bool noChange = false, bool blIsFormatValueAsString = true)
        {
            bool blnIsSuccess = false;

            int intStartRow = 1;
            int intEndRow = 1;
            int intStartColumn = 1;
            int intEndColumn = 1;

            Dictionary<string, int> dicHyperlinkColIndex = null;

            strErr = string.Empty;

            try
            {
                #region Pre-checking
                if (dsSource == null)
                    throw new Exception("Source dataset cannot be null. ");

                if (dsSource.Tables == null)
                    throw new Exception("Source dataset tables cannot be null");

                if (dsSource.Tables.Count == 0)
                    throw new Exception("There is no table exists in source dataset");

                if (string.IsNullOrWhiteSpace(strExcelOutputFolderPath))
                    throw new Exception("Output folder path cannot be null or empty. ");

                if (string.IsNullOrWhiteSpace(strExcelFileNm))
                    throw new Exception("Output file name cannot be null or empty. ");

                if (!Directory.Exists(strExcelOutputFolderPath))
                    throw new Exception("Output folder path not exists");

                if (!(strExcelFileNm.ToUpper().EndsWith(".XLS") || strExcelFileNm.ToUpper().EndsWith(".XLSX")))
                    throw new Exception(string.Format("Output file format {0} is not applicable. ", strExcelFileNm));
                #endregion

                #region Initialize variables
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Visible = false;
                ExcelApp.DisplayAlerts = !noChange;
                ExcelBook = ExcelApp.Workbooks.Add(missing);

                int iSheetIndex = 1;
                string strFilePathFull = Path.Combine(strExcelOutputFolderPath, strExcelFileNm);
                #endregion

                foreach (System.Data.DataTable dtblSource in dsSource.Tables)
                {
                    if (iSheetIndex == 1)
                        ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Sheets.get_Item(iSheetIndex);
                    else
                        ExcelSheet = ExcelBook.Sheets.Add(After: ExcelBook.Sheets[ExcelBook.Sheets.Count]);

                    // Get row number to start insert record
                    intStartRow = 1;
                    intStartColumn = 1;

                    intEndRow = intStartRow + dtblSource.Rows.Count;
                    intEndColumn = dtblSource.Columns.Count == 0 ? 1 : dtblSource.Columns.Count;

                    range1 = ExcelSheet.Range[ExcelSheet.Cells[intStartRow, intStartColumn], ExcelSheet.Cells[intEndRow, intEndColumn]];

                    #region Write header value
                    dicHyperlinkColIndex = null; //Reset for every sheet
                    for (int i = 0; i < dtblSource.Columns.Count; i++)
                    {
                        range1.Cells[intStartRow, i + 1] = string.Format("{0}", dtblSource.Columns[i].ColumnName);
                        range1.Cells[intStartRow, i + 1].Rows.Interior.Color = XlRgbColor.rgbLightGrey;

                        if (HyperlinkColNmList != null
                            && HyperlinkColNmList.Count > 0
                            && HyperlinkColNmList.Contains(dtblSource.Columns[i].ColumnName))
                        {
                            dicHyperlinkColIndex = dicHyperlinkColIndex ?? new Dictionary<string, int>();
                            if (!dicHyperlinkColIndex.ContainsKey(dtblSource.Columns[i].ColumnName))
                                dicHyperlinkColIndex.Add(dtblSource.Columns[i].ColumnName, i + 1);
                        }
                    }
                    #endregion

                    #region Write content value
                    var data = new object[dtblSource.Rows.Count, dtblSource.Columns.Count];
                    for (int i = 0; i < dtblSource.Rows.Count; i++)
                    {
                        for (int j = 0; j < dtblSource.Columns.Count; j++)
                        {
                            data[i, j] = string.Format(blIsFormatValueAsString ? "'{0}" : "{0}", dtblSource.Rows[i][j]);
                        }
                    }

                    var startCell = (Range)ExcelSheet.Cells[intStartRow + 1, intStartColumn];
                    var endCell = (Range)ExcelSheet.Cells[intEndRow, intEndColumn];
                    var writeRange = ExcelSheet.Range[startCell, endCell];
                    writeRange.Value2 = data;
                    writeRange.Columns.AutoFit();
                    #endregion

                    #region Start adding hyperlink if have
                    if (dicHyperlinkColIndex != null && dicHyperlinkColIndex.Count > 0)
                    {
                        int iTotalDataRow = ExcelSheet.UsedRange.Rows.Count;
                        for (int iRow = intStartRow + 1; iRow <= iTotalDataRow; iRow++)
                        {
                            foreach (var keyValuePair in dicHyperlinkColIndex)
                            {
                                ExcelSheet.Hyperlinks.Add(ExcelSheet.Cells[iRow, keyValuePair.Value], ExcelSheet.Cells[iRow, keyValuePair.Value].Value);
                            }
                        }
                    }
                    #endregion

                    ExcelSheet.Name = dtblSource.TableName;
                    iSheetIndex++;
                }

                ExcelBook.SaveAs(strFilePathFull, GetFileFormat((strFilePathFull.ToUpper().EndsWith(".XLSX") ? ExcelFileFormat.XLSX : ExcelFileFormat.XLS)),
                            Type.Missing, Type.Missing, false, false, noChange ? XlSaveAsAccessMode.xlShared : XlSaveAsAccessMode.xlExclusive,
                            noChange ? Type.Missing : XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                ExcelBook.Close();
                blnIsSuccess = true;
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strErr = ex.Message;
            }
            finally
            {
                ReleaseExcelObject();
            }

            return blnIsSuccess;
        }

        public bool ExcelToDataSet(string strExcelFilePath, out DataSet dsResult, out string strExceptionMessage)
        {
            return ExcelToDataSet(strExcelFilePath, 1, out dsResult, out strExceptionMessage, true);
        }

        public bool ExcelToDataSet(string strExcelFilePath, int iHeaderRowIndex, out DataSet dsResult, out string strExceptionMessage)
        {
            return ExcelToDataSet(strExcelFilePath, iHeaderRowIndex, out dsResult, out strExceptionMessage, true);
        }

        public bool ExcelToDataSet(string strExcelFilePath, int iHeaderRowIndex, out DataSet dsResult, out string strExceptionMessage, bool blUseHeaderRow = true)
        {
            bool blnIsSuccess = true;
            FileStream fStream = null;
            IExcelDataReader excelReader = null;

            dsResult = null;
            strExceptionMessage = string.Empty;

            try
            {
                if (!File.Exists(strExcelFilePath))
                    throw new Exception(string.Format("There is no excel file exists in the file path. {0}", strExcelFilePath));

                fStream = File.Open(strExcelFilePath, FileMode.Open, FileAccess.Read);

                if (strExcelFilePath.ToLower().EndsWith(".xls"))
                    excelReader = ExcelReaderFactory.CreateBinaryReader(fStream);
                else if (strExcelFilePath.ToLower().EndsWith(".xlsx"))
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(fStream);
                else if (strExcelFilePath.ToLower().EndsWith(".csv"))
                    excelReader = ExcelReaderFactory.CreateCsvReader(fStream);

                dsResult = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = blUseHeaderRow,
                        ReadHeaderRow = (rowReader) =>
                        {
                            for (int iIndex = 0; iIndex < iHeaderRowIndex - 1; iIndex++)
                            {
                                if (rowReader == null)
                                    break;

                                rowReader.Read();
                            }
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                blnIsSuccess = false;
                strExceptionMessage = ex.Message;
            }
            finally
            {
                if (excelReader != null) { excelReader.Close(); }
                if (fStream != null) { fStream.Dispose(); }
            }

            return blnIsSuccess;
        }

        public bool GetFilteredRowCount(Range rngFiltered, int intHeaderRow, out int Output, out string strErr)
        {
            bool blSuccess = false;
            Output = 0;
            strErr = string.Empty;

            try
            {
                foreach (Range area in rngFiltered.Areas)
                {
                    foreach (Range row in area.Rows)
                    {
                        if (intHeaderRow != row.Row)
                        {
                            Output++;
                        }
                    }
                }
                blSuccess = true;
            }
            catch (Exception ex)
            {
                blSuccess = false;
                strErr = ex.Message;
            }

            return blSuccess;
        }

        public bool UpdateFilterColumn(Range rngFiltered, int intHeaderRow,int intColumnToUpdate, string strValue, out string strErr)
        {
            bool blSuccess = false;
            strErr = string.Empty;

            try
            {
                foreach (Range area in rngFiltered.Areas)
                {
                    foreach (Range row in area.Rows)
                    {
                        if (intHeaderRow != row.Row)
                        {
                            rngFiltered.Worksheet.Cells[row.Row, intColumnToUpdate] = strValue;
                        }
                    }
                }
                blSuccess = true;
            }
            catch (Exception ex)
            {
                blSuccess = false;
                strErr = ex.Message;
            }

            return blSuccess;
        }

        public void ReleaseExcelObject()
        {
            if (this.ExcelSheet != null) Marshal.ReleaseComObject(this.ExcelSheet);
            if (this.ExcelBook != null) Marshal.ReleaseComObject(this.ExcelBook);
            if (this.ExcelApp != null) { ExcelApp.Quit(); Marshal.ReleaseComObject(this.ExcelApp); }
        }

        public string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        public void CloseOpenExcel(string strExportFileName)
        {
            int intCounter = 0;
            while (intCounter <= 2)
            {
                bool blnIsWindowFound = ProcessesAndWindowsHelper.WaitForWindowToExists(strExportFileName, 5);

                if (blnIsWindowFound)
                {
                    System.Threading.Thread.Sleep(1000);

                    try
                    {
                        Microsoft.Office.Interop.Excel.Application app =
                            (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
                    }
                    catch
                    {
                        // Excel is not running.
                    }
                }

                try
                {
                    var processes = from p in Process.GetProcessesByName("EXCEL")
                                    select p;

                    foreach (var process in processes)
                    {
                        process.Kill();
                    }
                }
                catch
                {
                    // Excel is not running.
                }

                intCounter++;
            }
        }
        
        public int GetSheetTotalRow(Microsoft.Office.Interop.Excel.Worksheet targetExcelWorkSheet)
        {
            int iSheetTotalRow = 1;

            try
            {
                iSheetTotalRow = targetExcelWorkSheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, 
                                                                 XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                                                                 false, Missing.Value, Missing.Value).Row;
            }
            catch (Exception ex)
            {
                throw ex; //Exception to be handle by caller
            }

            return iSheetTotalRow;
        }

        public int GetSheetTotalColumn(Microsoft.Office.Interop.Excel.Worksheet targetExcelWorkSheet)
        {
            int iSheetTotalCol = 1;

            try
            {
                iSheetTotalCol = targetExcelWorkSheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value,
                                                                 XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious,
                                                                 false, Missing.Value, Missing.Value).Column;
            }
            catch (Exception ex)
            {
                throw ex; //Exception to be handle by caller
            }

            return iSheetTotalCol;
        }

        public int? GetRowIndexOfValue(Range targetExcelRange, string strTextToFind)
        {
            int? iRowIndex = null;

            try
            {
                Range rngFoundCell = targetExcelRange.Find(What:strTextToFind);
                if (rngFoundCell != null)
                    iRowIndex = rngFoundCell.Row;
                else
                    iRowIndex = null;
            }
            catch (Exception ex)
            {
                iRowIndex = null;
            }

            return iRowIndex;
        }

        public int? GetColIndexOfValue(Range targetExcelRange, string strTextToFind)
        {
            int? iColIndex = null;

            try
            {
                Range rngFoundCell = targetExcelRange.Find(What: strTextToFind);
                if (rngFoundCell != null)
                    iColIndex = rngFoundCell.Column;
                else
                    iColIndex = null;
            }
            catch (Exception ex)
            {
                iColIndex = null;
            }

            return iColIndex;
        }

        public bool CheckIsSheetExists(Microsoft.Office.Interop.Excel.Workbook targetExcelBook, string strSheetNmToCheck)
        {
            bool blnIsSheetExists = false;

            try
            {
                #region Pre-Checking
                if (targetExcelBook == null)
                    return false;
                if (string.IsNullOrWhiteSpace(strSheetNmToCheck))
                    return false;
                #endregion

                #region Check Sheet Exists
                blnIsSheetExists = targetExcelBook
                                   .Worksheets
                                   .OfType<Microsoft.Office.Interop.Excel.Worksheet>()
                                   .Any(ws => !string.IsNullOrWhiteSpace(ws.Name)
                                              && ws.Name.Trim().Equals(strSheetNmToCheck.Trim()));
                #endregion
            }
            catch (Exception ex)
            {
                blnIsSheetExists = false;
            }

            return blnIsSheetExists;
        }

        public string FormatCellValueAsString(string strInput)
        {
            if (!string.IsNullOrWhiteSpace(strInput))
                return string.Format("'{0}", strInput);
            else
                return strInput;
        }

        public void ReleaseExcelObject(Microsoft.Office.Interop.Excel.Application targetExcelApp, Microsoft.Office.Interop.Excel.Workbook targetExcelBook, Microsoft.Office.Interop.Excel.Worksheet targetExcelWorkSheet)
        {
            if (targetExcelWorkSheet != null) Marshal.ReleaseComObject(targetExcelWorkSheet);
            if (targetExcelBook != null) Marshal.ReleaseComObject(targetExcelBook);
            if (targetExcelApp != null) { targetExcelApp.Quit(); Marshal.ReleaseComObject(targetExcelApp); }
        }

        public bool SetWorkbookConnection(_Workbook wb, string connName, string cmdText, out string errMsg)
        {
            bool success = true;
            errMsg = string.Empty;

            try
            {
                foreach (WorkbookConnection connection in wb.Connections)
                {
                    if (connection.Name.Equals(connName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        WorksheetDataConnection wsConn = connection.WorksheetDataConnection;
                        wsConn.CommandText = cmdText;
                    }
                }
            }
            catch (Exception ex)
            {
                success = false;
                errMsg = string.Format("Name: {0}, Command: {1}, Error: {2}", connName, cmdText, ex.Message);
            }

            return success;
        }

        public static Workbook OpenProtectedExcelFile(Application excelApplication, string strExcelFilePath, string strExcelFilePassword = "")
        {
            Workbook excelWorkbook = null;

            try
            {
                string strExcelFileExt = Path.GetExtension(strExcelFilePath);
                if (!(!string.IsNullOrWhiteSpace(strExcelFilePath) && new List<string> { ".xls", ".xlsx", ".xlsm" }.Contains(strExcelFileExt.Trim())))
                    throw new Exception($"Excel file extension {strExcelFileExt} is not valid");

                try
                {
                    excelWorkbook = excelApplication.Workbooks.Open(strExcelFilePath, Type.Missing, false, Password:strExcelFilePassword);
                }
                catch (Exception exp)
                {
                    if (!string.IsNullOrWhiteSpace(exp.Message) && !exp.Message.Trim().StartsWith(CONST_INCORRECT_EXCEL_FILE_PASSWORD_ERR_MSG))
                    {
                        //Force to edit protected view excel file
                        ProtectedViewWindow procViewWindow = excelApplication.ProtectedViewWindows.Open(strExcelFilePath, Password: strExcelFilePassword);
                        procViewWindow.Activate();
                        excelWorkbook = procViewWindow.Edit();
                    }
                    else
                    {
                        throw exp;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return excelWorkbook;
        }
        
        public static Worksheet GetFirstVisibleWorkSheet(Microsoft.Office.Interop.Excel.Workbook targetExcelBook)
        {
            Worksheet wsFirst = null;

            try
            {
                if (targetExcelBook == null) { throw new Exception("Target excel workbook is null"); }

                foreach (Worksheet ws in targetExcelBook.Worksheets)
                {
                    if (ws.Visible == XlSheetVisibility.xlSheetVisible)
                    {
                        wsFirst = ws;
                        break;
                    }
                }

                if (wsFirst == null) { throw new Exception("No visible worksheet available"); }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return wsFirst;
        }

        public int GetColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        private string GetColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }
        #endregion

        #region Helpers
        private XlFileFormat GetFileFormat(ExcelFileFormat outFormat)
        {
            XlFileFormat fmt = XlFileFormat.xlWorkbookDefault;

            switch (outFormat)
            {
                case ExcelFileFormat.XLSX:
                    fmt = XlFileFormat.xlOpenXMLWorkbook;
                    break;
                case ExcelFileFormat.XLS:
                    fmt = XlFileFormat.xlWorkbookNormal;
                    break;
                default:
                    fmt = XlFileFormat.xlWorkbookDefault;
                    break;
            }

            return fmt;
        }
        #endregion

        #region enum
        public enum ExcelFileFormat
        {
            XLS,
            XLSX
        }
        #endregion
    }
}