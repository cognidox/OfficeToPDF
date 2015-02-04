/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010
 *  Copyright (C) 2011-2014 Cognidox Ltd
 *  http://www.cognidox.com/opensource/
 *
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 *
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Excel files
    /// </summary>
    class ExcelConverter: Converter
    {
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            Boolean running = (Boolean)options["noquit"];
            Microsoft.Office.Interop.Excel.Application app = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;

            String tmpFile = null;
            object oMissing = System.Reflection.Missing.Value;
            Boolean nowrite = (Boolean)options["readonly"];
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                app.DisplayAlerts = false;
                app.AskToUpdateLinks = false;
                app.AlertBeforeOverwriting = false;
                app.EnableLargeOperationAlert = false;
                app.Interactive = false;
                app.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                if ((Boolean)options["hidden"])
                {
                    // Try and at least minimise it
                    app.WindowState = XlWindowState.xlMinimized;
                    app.Visible = false;
                }

                String readPassword = "";
                if (!String.IsNullOrEmpty((String)options["password"]))
                {
                    readPassword = (String)options["password"];
                }
                Object oReadPass = (Object)readPassword;

                String writePassword = "";
                if (!String.IsNullOrEmpty((String)options["writepassword"]))
                {
                    writePassword = (String)options["writepassword"];
                }
                Object oWritePass = (Object)writePassword;

                // Check for password protection and no password
                if (Converter.IsPasswordProtected(inputFile) && String.IsNullOrEmpty(readPassword))
                {
                    Console.WriteLine("Unable to open password protected file");
                    return (int)ExitCode.PasswordFailure;
                }

                workbooks = app.Workbooks;
                workbook = workbooks.Open(inputFile, true, nowrite, oMissing, oReadPass, oWritePass, true, oMissing, oMissing, oMissing, oMissing, oMissing, false, oMissing, oMissing);

                // Unable to open workbook
                if (workbook == null)
                {
                    return (int)ExitCode.FileOpenFailure;
                }

                // Try and avoid xls files raising a dialog
                tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".xls";
                XlFileFormat fmt = XlFileFormat.xlOpenXMLWorkbook;
                XlFixedFormatQuality quality = XlFixedFormatQuality.xlQualityStandard;
                if (workbook.HasVBProject)
                {
                    fmt = XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                    tmpFile += "m";
                }
                else
                {
                    tmpFile += "x";
                }

                // Large excel files may simply not print reliably - if the excel_max_rows
                // configuration option is set, then we must close up and forget about 
                // converting the file. However, if a print area is set in one of the worksheets
                // in the document, then assume the author knew what they were doing and
                // use the print area.
                var max_rows = (int)options[@"excel_max_rows"];
                if (max_rows > 0)
                {
                    // Loop through all the worksheets in the workbook looking to any
                    // that have too many rows
                    var worksheets = workbook.Worksheets;
                    var row_count_check_ok = true;
                    var found_rows = 0;
                    var found_worksheet = "";
                    foreach (var ws in worksheets)
                    {
                        // Check for a print area
                        var page_setup = ((Microsoft.Office.Interop.Excel.Worksheet)ws).PageSetup;
                        var print_area = page_setup.PrintArea;
                        Converter.releaseCOMObject(page_setup);
                        if (string.IsNullOrEmpty(print_area))
                        {
                            // There is no print area, check that the row count is <= to the
                            // excel_max_rows value. Note that we can't just take the range last
                            // row, as this may return a huge value, rather find the last non-blank
                            // row.
                            var range = ((Microsoft.Office.Interop.Excel.Worksheet)ws).UsedRange;
                            var cells = range.Cells;
                            var cellSearch = cells.Find("*", oMissing, oMissing, oMissing, oMissing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, oMissing, oMissing);
                            var row_count = cellSearch.Row;
                            found_worksheet = ((Microsoft.Office.Interop.Excel.Worksheet)ws).Name;
                            Converter.releaseCOMObject(cellSearch);
                            Converter.releaseCOMObject(cells);
                            Converter.releaseCOMObject(range);
                            Converter.releaseCOMObject(ws);

                            if (row_count > max_rows)
                            {
                                // Too many rows on this worksheet - mark the workbook as unprintable
                                row_count_check_ok = false;
                                found_rows = row_count;
                                break;
                            }
                        }
                    }
                    Converter.releaseCOMObject(worksheets);
                    if (!row_count_check_ok)
                    {
                        throw new Exception(String.Format("Too many rows to process ({0}) on worksheet {1}", found_rows, found_worksheet));
                    }
                }

                // Remember - Never use 2 dots with COM objects!
                // Using more than one dot leaves wrapper objects left over
                var wbWin = workbook.Windows;
                var appWin = app.Windows;
                if (wbWin.Count > 0)
                {
                    wbWin[1].Visible = (Boolean)options["hidden"] ? false : true;
                    Converter.releaseCOMObject(wbWin);
                }
                if (appWin.Count > 0)
                {
                    appWin[1].Visible = (Boolean)options["hidden"] ? false : true;
                    Converter.releaseCOMObject(appWin);
                }
                Boolean includeProps = !(Boolean)options["excludeprops"];

                workbook.SaveAs(tmpFile, fmt, Type.Missing, Type.Missing, Type.Missing, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing);
                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                    outputFile, quality, includeProps, false, Type.Missing, Type.Missing, false, Type.Missing);
                return (int)ExitCode.Success;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }

                if (!running)
                {
                    if (workbooks != null)
                    {
                        workbooks.Close();
                    }

                    if (app != null)
                    {
                        ((Microsoft.Office.Interop.Excel._Application)app).Quit();
                    }
                }

                // Clean all the COM leftovers
                Converter.releaseCOMObject(workbook);
                Converter.releaseCOMObject(workbooks);
                Converter.releaseCOMObject(app);

                if (tmpFile != null)
                {
                    System.IO.File.Delete(tmpFile);
                }
            }
        }
    }
}
