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
using Microsoft.Office.Interop.Excel;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Excel files
    /// </summary>
    class ExcelConverter: Converter
    {
        public static new Boolean Convert(String inputFile, String outputFile, Hashtable options)
        {
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
                app.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                if ((Boolean)options["hidden"])
                {
                    // Try and at least minimise it
                    app.WindowState = XlWindowState.xlMinimized;
                    app.Visible = false;
                }
                workbooks = app.Workbooks;
                workbook = workbooks.Open(inputFile, true, nowrite, oMissing, oMissing, oMissing, true, oMissing, oMissing, oMissing, oMissing, oMissing, false, oMissing, oMissing);

                // Unable to open workbook
                if (workbook == null)
                {
                    return false;
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

                // Remember - Never use 2 dots with COM objects!
                // Using more than one dot leaves wrapper objects left over
                var wbWin = workbook.Windows;
                var appWin = app.Windows;
                if (wbWin.Count > 0)
                {
                    wbWin[1].Visible = (Boolean)options["hidden"] ? false : true;
                }
                if (appWin.Count > 0)
                {
                    appWin[1].Visible = (Boolean)options["hidden"] ? false : true;
                }
                workbook.SaveAs(tmpFile, fmt, Type.Missing, Type.Missing, Type.Missing, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing);
                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                    outputFile, quality, Type.Missing, false, Type.Missing, Type.Missing, false, Type.Missing);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }

                if (workbooks != null)
                {
                    workbooks.Close();
                }

                if (app != null)
                {
                    ((Microsoft.Office.Interop.Excel._Application)app).Quit();
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
