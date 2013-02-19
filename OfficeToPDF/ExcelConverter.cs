/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010
 *  Copyright (C) 2011  Cognidox Ltd
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
            MessageFilter.Register();
            Microsoft.Office.Interop.Excel.Application app = null;
            String tmpFile = null;
            object oMissing = System.Reflection.Missing.Value;
            Boolean nowrite = (Boolean)options["readonly"];
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                if ((Boolean)options["hidden"])
                {
                    // Try and at least minimise it
                    app.WindowState = XlWindowState.xlMinimized;
                    app.Visible = false;
                }
                Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
                Microsoft.Office.Interop.Excel.Workbook workbook = null;
                workbooks = app.Workbooks;
                workbook = workbooks.Open(inputFile, true, nowrite, oMissing, oMissing, oMissing, true, oMissing, oMissing, oMissing, oMissing, oMissing, false, oMissing, oMissing);

                // Try and avoid xls files raising a dialog
                tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".xls";
                XlFileFormat fmt = XlFileFormat.xlOpenXMLWorkbook;
                if (workbook.HasVBProject)
                {
                    fmt = XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                    tmpFile += "m";
                }
                else
                {
                    tmpFile += "x";
                }
                workbook.Windows[1].Visible = (Boolean)options["hidden"] ? false : true;
                app.Windows[1].Visible = (Boolean)options["hidden"] ? false : true;
                workbook.SaveAs(tmpFile, fmt, Type.Missing, Type.Missing, Type.Missing, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing);
                workbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                outputFile, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false, Type.Missing, Type.Missing, false, Type.Missing);
                workbooks.Close();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            finally
            {
                if (tmpFile != null)
                {
                    System.IO.File.Delete(tmpFile);
                }
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
                 MessageFilter.Revoke();
            }
        }
    }
}
