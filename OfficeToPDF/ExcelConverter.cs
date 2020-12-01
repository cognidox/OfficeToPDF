/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010/2013/2016
 *  Copyright (C) 2011-2018 Cognidox Ltd
 *  https://www.cognidox.com/opensource/
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
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Excel files
    /// </summary>
    class ExcelConverter: Converter
    {
        // These are the properties we will extract from the first worksheet of any template document
        private static string[] templateProperties = 
        {
            "BlackAndWhite", "BottomMargin", "CenterFooter", "CenterHeader",
            "CenterHorizontally", "CenterVertically", "DifferentFirstPageHeaderFooter",
            "Draft", "FirstPageNumber", "FitToPagesTall", "FitToPagesWide",
            "FooterMargin", "HeaderMargin", "LeftFooter", "LeftHeader",
            "LeftMargin", "OddAndEvenPagesHeaderFooter", "Order", "Orientation", "PaperSize", "PrintArea",
            "PrintGridlines", "PrintHeadings", "PrintNotes", "PrintTitleColumns", "PrintTitleRows",
            "RightFooter", "RightHeader", "RightMargin",
            "ScaleWithDocHeaderFooter", "TopMargin", "Zoom"
        };

        // Main conversion routine
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            Boolean running = (Boolean)options["noquit"];
            Microsoft.Office.Interop.Excel.Application app = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            System.Object activeSheet = null;
            Window activeWindow = null;
            Windows wbWin = null;
            Hashtable templatePageSetup = new Hashtable();

            String tmpFile = null;
            object oMissing = System.Reflection.Missing.Value;
            Boolean nowrite = (Boolean)options["readonly"];
            try
            {
                // Excel can be very slow to start up, so try to get the COM
                // object a few times
                int tries = 10;
                app = new Microsoft.Office.Interop.Excel.Application();
                while (tries > 0)
                {
                    try
                    {
                        // Try to set a property on the object
                        app.ScreenUpdating = false;
                    }
                    catch (COMException)
                    {
                        // Decrement the number of tries and have a bit of a snooze
                        tries--;
                        Thread.Sleep(500);
                        continue;
                    }
                    // Looks ok, so bail out of the loop
                    break;
                }
                if (tries == 0)
                {
                    ReleaseCOMObject(app);
                    return (int)ExitCode.ApplicationError;
                }

                app.Visible = true;
                app.DisplayAlerts = false;
                app.AskToUpdateLinks = false;
                app.AlertBeforeOverwriting = false;
                app.EnableLargeOperationAlert = false;
                app.Interactive = false;
                app.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;

                var onlyActiveSheet = (Boolean)options["excel_active_sheet"];
                Boolean activeSheetOnMaxRows = (Boolean)options["excel_active_sheet_on_max_rows"];
                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean skipRecalculation = (Boolean)options["excel_no_recalculate"];
                Boolean showHeadings = (Boolean)options["excel_show_headings"];
                Boolean showFormulas = (Boolean)options["excel_show_formulas"];
                Boolean isHidden = (Boolean)options["hidden"];
                Boolean screenQuality = (Boolean)options["screen"];
                Boolean updateLinks = !(Boolean)options["excel_no_link_update"];
                int maxRows = (int)options[@"excel_max_rows"];
                int worksheetNum = (int)options["excel_worksheet"];
                int sheetForConversionIdx = 0;
                activeWindow = app.ActiveWindow;
                Sheets allSheets = null;
                XlFileFormat fmt = XlFileFormat.xlOpenXMLWorkbook;
                XlFixedFormatQuality quality = XlFixedFormatQuality.xlQualityStandard;
                if (isHidden)
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

                app.EnableEvents = (bool)options["excel_auto_macros"];
                workbooks = app.Workbooks;
                // If we have no write password and we're attempting to open for writing, we might be
                // caught out by an unexpected write password
                if (writePassword == "" && !nowrite)
                {
                    oWritePass = (Object)"FAKEPASSWORD";
                    try
                    {
                        workbook = workbooks.Open(inputFile, updateLinks, nowrite, oMissing, oReadPass, oWritePass, true, oMissing, oMissing, oMissing, oMissing, oMissing, false, oMissing, oMissing);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Attempt to open it in read-only mode
                        workbook = workbooks.Open(Filename: inputFile, UpdateLinks: updateLinks, ReadOnly: true, Password: oReadPass, WriteResPassword: oWritePass, IgnoreReadOnlyRecommended: true,
                            AddToMru: false, CorruptLoad: GetXlRepairFile());
                    }
                }
                else
                {
                    workbook = workbooks.Open(Filename: inputFile, UpdateLinks: updateLinks, ReadOnly: nowrite, Password: oReadPass, WriteResPassword: oWritePass,
                        IgnoreReadOnlyRecommended: true, AddToMru: false, CorruptLoad: GetXlRepairFile());
                }

                // Add in a delay to let Excel sort itself out
                AddCOMDelay(options);

                // Unable to open workbook
                if (workbook == null)
                {
                    return (int)ExitCode.FileOpenFailure;
                }

                if (app.EnableEvents)
                {
                    workbook.RunAutoMacros(XlRunAutoMacro.xlAutoOpen);
                }
                
                // Get any template options
                SetPageOptionsFromTemplate(app, workbooks, options, ref templatePageSetup);

                // Get the sheets
                allSheets = workbook.Sheets;

                // Try and avoid xls files raising a dialog
                var temporaryStorageDir = Path.GetTempFileName();
                File.Delete(temporaryStorageDir);
                Directory.CreateDirectory(temporaryStorageDir);
                // We will save as xlsb (binary format) since this doesn't raise some errors when processing
                tmpFile = Path.Combine(temporaryStorageDir, Path.GetFileNameWithoutExtension(inputFile) + ".xlsb");
                fmt = XlFileFormat.xlExcel12;

                // Set up the print quality
                if (screenQuality)
                {
                    quality = XlFixedFormatQuality.xlQualityMinimum;
                }

                // If a worksheet has been specified, try and use just the one
                if (worksheetNum > 0)
                {
                    // Force us just to use the active sheet
                    onlyActiveSheet = true;
                    try
                    {
                        if (worksheetNum > allSheets.Count)
                        {
                            // Sheet count is too big
                            return (int)ExitCode.WorksheetNotFound;
                        }
                        if (allSheets[worksheetNum] is _Worksheet)
                        {
                            ((_Worksheet)allSheets[worksheetNum]).Activate();
                            sheetForConversionIdx = ((_Worksheet)allSheets[worksheetNum]).Index;
                        }
                        else if (allSheets[worksheetNum] is _Chart)
                        {
                            ((_Chart)allSheets[worksheetNum]).Activate();
                            sheetForConversionIdx = ((_Chart)allSheets[worksheetNum]).Index;
                        }

                    }
                    catch (Exception)
                    {
                        return (int)ExitCode.WorksheetNotFound;
                    }
                }
                
                if (showFormulas)
                {
                    // Determine whether to show formulas
                    try
                    {
                        activeWindow.DisplayFormulas = true;
                    }
                    catch (Exception) { }
                }

                // Keep the windows hidden
                if (isHidden)
                {
                    wbWin = workbook.Windows;
                    if (null != wbWin)
                    {
                        if (wbWin.Count > 0)
                        {
                            wbWin[1].Visible = false;
                        }
                    }
                    if (null != activeWindow)
                    {
                        activeWindow.Visible = false;
                    }
                }

                // Keep track of the active sheet
                int activeSheetIdx = 1;
                if (workbook.ActiveSheet != null)
                {
                    activeSheet = workbook.ActiveSheet;
                    if (activeSheet is _Worksheet)
                    {
                        activeSheetIdx = ((Worksheet)activeSheet).Index;
                    }
                    else if (activeSheet is _Chart)
                    {
                        activeSheetIdx = ((Microsoft.Office.Interop.Excel.Chart)activeSheet).Index;
                    }
                }

                // Large excel files may simply not print reliably - if the excel_max_rows
                // configuration option is set, then we must close up and forget about 
                // converting the file. However, if a print area is set in one of the worksheets
                // in the document, then assume the author knew what they were doing and
                // use the print area.

                // We may need to loop through all the worksheets in the document
                // depending on the options given. If there are maximum row restrictions
                // or formulas are being shown, then we need to loop through all the
                // worksheets
                if (maxRows > 0 || showFormulas || showHeadings)
                {
                    var row_count_check_ok = true;
                    var found_rows = 0;
                    var found_worksheet = "";
                    bool[] rowCountOK = new bool[allSheets.Count + 1];

                    // Loop through all the sheets (worksheets and charts)
                    for (int wsIdx = 1; wsIdx <= allSheets.Count; wsIdx++)
                    {
                        var ws = allSheets.Item[wsIdx];
                        rowCountOK[wsIdx] = true;

                        // Skip anything that is not the active sheet
                        if (onlyActiveSheet)
                        {
                            // Have to be careful to treat _Worksheet and _Chart items differently
                            try
                            {
                                // Get the index of the active sheet
                                if (wsIdx != activeSheetIdx)
                                {
                                    // If we are not the active sheet, then skip to the next
                                    ReleaseCOMObject(ws);
                                    continue;
                                }
                            }
                            catch (Exception)
                            {
                                if (ws != null)
                                {
                                    ReleaseCOMObject(ws);
                                }
                                continue;
                            }
                            sheetForConversionIdx = wsIdx;
                        }

                        if (showHeadings && ws is _Worksheet)
                        {
                            PageSetup pageSetup = null;
                            try
                            {
                                pageSetup = ((Worksheet)ws).PageSetup;
                                pageSetup.PrintHeadings = true;
                                
                            }
                            catch (Exception) { }
                            finally
                            {
                                ReleaseCOMObject(pageSetup);
                            }
                        }

                        // If showing formulas, make things auto-fit
                        if (showFormulas && ws is _Worksheet)
                        {
                            Range cols = null;
                            try
                            {
                                ((_Worksheet)ws).Activate();
                                activeWindow.DisplayFormulas = true;
                                cols = ((Worksheet)ws).Columns;
                                cols.AutoFit();
                            }
                            catch (Exception) { }
                            finally
                            {
                                ReleaseCOMObject(cols);
                            }
                        }

                        // If there is a maximum row count, make sure we check each worksheet
                        if (maxRows > 0 && ws is _Worksheet)
                        {
                            // Check for a print area
                            var pageSetup = ((Worksheet)ws).PageSetup;
                            var printArea = pageSetup.PrintArea;
                            ReleaseCOMObject(pageSetup);
                            if (string.IsNullOrEmpty(printArea))
                            {
                                // There is no print area, check that the row count is <= to the
                                // excel_max_rows value. Note that we can't just take the range last
                                // row, as this may return a huge value, rather find the last non-blank
                                // row.
                                var row_count = 0;
                                var range = ((Worksheet)ws).UsedRange;
                                if (range != null)
                                {
                                    var rows = range.Rows;
                                    if (rows != null && rows.Count > maxRows)
                                    {
                                        var cells = range.Cells;
                                        if (cells != null)
                                        {
                                            var cellSearch = cells.Find("*", oMissing, oMissing, oMissing, oMissing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, oMissing, oMissing);
                                            // Make sure we actually get some results, since the worksheet may be totally blank
                                            if (cellSearch != null)
                                            {
                                                row_count = cellSearch.Row;
                                                found_worksheet = ((Worksheet)ws).Name;
                                            }
                                            ReleaseCOMObject(cellSearch);
                                        }
                                        ReleaseCOMObject(cells);
                                    }
                                    ReleaseCOMObject(rows);
                                }
                                ReleaseCOMObject(range);

                                if (row_count > maxRows)
                                {
                                    // Too many rows on this worksheet - mark the workbook as unprintable
                                    row_count_check_ok = false;
                                    rowCountOK[wsIdx] = false;
                                    found_rows = row_count;
                                    Converter.ReleaseCOMObject(ws);
                                    if (activeSheetOnMaxRows)
                                    {
                                        // Keep checking
                                        continue;
                                    }
                                    break;
                                }
                            }
                        } // End of row check
                        Converter.ReleaseCOMObject(ws);
                    }

                    // Make sure we are not converting a document with too many rows
                    if (row_count_check_ok == false)
                    {
                        // We may want to try and convert the active sheet if it has not been included in the
                        // sheets with too many rows
                        bool bailOut = true;
                        if (activeSheetOnMaxRows && !onlyActiveSheet)
                        {
                            if (rowCountOK[activeSheetIdx])
                            {
                                bailOut = false;
                                sheetForConversionIdx = activeSheetIdx;
                                onlyActiveSheet = true;
                            }
                        }
                        if (bailOut)
                        {
                            throw new Exception(String.Format("Too many rows to process ({0}) on worksheet {1}", found_rows, found_worksheet));
                        }
                    }
                }

                // Allow for re-calculation to be skipped
                if (skipRecalculation)
                {
                    app.Calculation = XlCalculation.xlCalculationManual;
                    app.CalculateBeforeSave = false;
                }

                workbook.SaveAs(tmpFile, fmt, Type.Missing, Type.Missing, Type.Missing, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing);

                if (onlyActiveSheet)
                {
                    // Set up a delegate function for times we want to print
                    PrintDocument printFunc = delegate (string destination, string printer)
                    {
                        ((Worksheet)activeSheet).PrintOut(ActivePrinter: printer, PrintToFile: true, PrToFileName: destination);
                    };

                    if (sheetForConversionIdx > 0)
                    {
                        activeSheet = allSheets.Item[sheetForConversionIdx];
                    }
                    if (activeSheet is _Worksheet)
                    {
                        var wps = ((_Worksheet)activeSheet).PageSetup;
                        SetPageSetupProperties(templatePageSetup, wps);
                        if (String.IsNullOrEmpty((string)options["printer"]))
                        {
                            try
                            {
                                ((Worksheet)activeSheet).ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                                outputFile, quality, includeProps, false, Type.Missing, Type.Missing, false, Type.Missing);
                            }
                            catch (Exception)
                            {
                                if (!String.IsNullOrEmpty((string)options["fallback_printer"]))
                                {
                                    PrintToGhostscript((string)options["fallback_printer"], outputFile, printFunc);
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                        else
                        {
                            PrintToGhostscript((string)options["printer"], outputFile, printFunc);
                        }
                        ReleaseCOMObject(wps);
                    }
                    else if (activeSheet is _Chart)
                    {
                        var wps = ((_Chart)activeSheet).PageSetup;
                        SetPageSetupProperties(templatePageSetup, wps);
                        ((Microsoft.Office.Interop.Excel.Chart)activeSheet).ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                            outputFile, quality, includeProps, false, Type.Missing, Type.Missing, false, Type.Missing);
                        ReleaseCOMObject(wps);
                    }
                    else
                    {
                        return (int)ExitCode.UnknownError;
                    }
                    AddCOMDelay(options);
                }
                else
                {
                    PrintDocument printFunc = delegate (string destination, string printer)
                    {
                        workbook.PrintOutEx(ActivePrinter: printer, PrintToFile: true, PrToFileName: destination);
                    };
                    if (HasTemplateOption(options))
                    {
                        // Set up the template page setup options on all the worksheets
                        // in the workbook
                        var worksheets = workbook.Worksheets;
                        for (int wsIdx = 1; wsIdx <= worksheets.Count; wsIdx++)
                        {
                            var ws = worksheets[wsIdx];
                            var wps = (ws is _Worksheet) ? ((_Worksheet)ws).PageSetup : ((_Chart)ws).PageSetup;
                            SetPageSetupProperties(templatePageSetup, wps);
                            ReleaseCOMObject(wps);
                            ReleaseCOMObject(ws);
                        }
                        ReleaseCOMObject(worksheets);
                    }
                    if (String.IsNullOrEmpty((string)options["printer"]))
                    {
                        try
                        {
                            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF,
                            outputFile, quality, includeProps, false, Type.Missing, Type.Missing, false, Type.Missing);
                        }
                        catch (Exception)
                        {
                            if (!String.IsNullOrEmpty((string)options["fallback_printer"]))
                            {
                                PrintToGhostscript((string)options["fallback_printer"], outputFile, printFunc);
                            }
                            else
                            {
                                throw;
                            }
                        }
                    } else {
                        PrintToGhostscript((string)options["printer"], outputFile, printFunc);
                    }
                }

                ReleaseCOMObject(allSheets);
                ReleaseCOMObject(fmt);
                ReleaseCOMObject(quality);

                return (int)ExitCode.Success;
            }
            catch (COMException ce)
            {
                if ((uint)ce.ErrorCode == 0x800A03EC)
                {
                    return (int)ExitCode.EmptyWorksheet;
                }
                else
                {
                    Console.WriteLine(ce.Message);
                    return (int)ExitCode.UnknownError;
                }
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
                    ReleaseCOMObject(activeSheet);
                    ReleaseCOMObject(activeWindow);
                    ReleaseCOMObject(wbWin);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    // Excel sometimes needs a bit of a delay before we close in order to
                    // let things get cleaned up
                    workbook.Saved = true;
                    CloseExcelWorkbook(workbook);
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
                ReleaseCOMObject(workbook);
                ReleaseCOMObject(workbooks);
                ReleaseCOMObject(app);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (tmpFile != null && File.Exists(tmpFile))
                {
                    System.IO.File.Delete(tmpFile);
                    // Remove the temporary path to the temp file
                    Directory.Delete(Path.GetDirectoryName(tmpFile));
                }
            }
        }

        private static XlCorruptLoad GetXlRepairFile()
        {
            return XlCorruptLoad.xlRepairFile;
        }

        // Return true if there is a valid template option
        protected static bool HasTemplateOption(Hashtable options)
        {
            if (String.IsNullOrEmpty((string)options["template"]) ||
                !File.Exists((string)options["template"]) ||
                !System.Text.RegularExpressions.Regex.IsMatch((string)options["template"], @"^.*\.xl[st][mx]?$"))
            {
                return false;
            }
            return true;
        }

        // Read the first worksheet from a template document and extract and store
        // the page settings for later use
        protected static void SetPageOptionsFromTemplate(Application app, Workbooks workbooks, Hashtable options, ref Hashtable templatePageSetup)
        {
            if (!HasTemplateOption(options))
            {
                return;
            }

            try
            {
                var template = workbooks.Open((string)options["template"]);
                AddCOMDelay(options);
                if (template != null)
                {
                    // Run macros from template if the /excel_template_macros option is given
                    if ((bool)options["excel_template_macros"])
                    {
                        var eventsEnabled = app.EnableEvents;
                        app.EnableEvents = true;
                        template.RunAutoMacros(XlRunAutoMacro.xlAutoOpen);
                        app.EnableEvents = eventsEnabled;
                    }

                    var templateSheets = template.Worksheets;
                    if (templateSheets != null)
                    {
                        // Copy the page setup details from the first sheet or chart in the template
                        if (templateSheets.Count > 0)
                        {
                            PageSetup tps = null;
                            var firstItem = templateSheets[1];
                            if (firstItem is _Worksheet)
                            {
                                tps = ((_Worksheet)firstItem).PageSetup;
                            }
                            else if (firstItem is _Chart)
                            {
                                tps = ((_Chart)firstItem).PageSetup;
                            }
                            var tpsType = tps.GetType();
                            for (int i = 0; i < templateProperties.Length; i++)
                            {
                                var prop = tpsType.InvokeMember(templateProperties[i], System.Reflection.BindingFlags.GetProperty, null, tps, null);
                                if (prop != null)
                                {
                                    templatePageSetup[templateProperties[i]] = prop;
                                }
                            }
                            Converter.ReleaseCOMObject(firstItem);
                        }
                        ReleaseCOMObject(templateSheets);
                    }
                    CloseExcelWorkbook(template);
                }
                ReleaseCOMObject(template);
            }
            finally
            {
            }
        }

        // Add in the required millisecond delay
        private static void AddCOMDelay(Hashtable options)
        {
            if ((int)options["excel_delay"] > 0)
            {
                Thread.Sleep((int)options["excel_delay"]);
            }
        }

        // Load stored worksheet properties into the page setup
        protected static void SetPageSetupProperties(Hashtable tps, PageSetup wps)
        {
            if (tps == null || tps.Count == 0)
            {
                return;
            }

            var wpsType = wps.GetType();
            for (int i = 0; i < templateProperties.Length; i++)
            {
                object[] value = { tps[templateProperties[i]] };
                try
                {
                    wpsType.InvokeMember(templateProperties[i], System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.SetProperty, Type.DefaultBinder, wps, value);
                }
                catch(Exception)
                {
                    Console.WriteLine("Unable to set property {0}", templateProperties[i]);
                }
            }
        }

        private static bool CloseExcelWorkbook(Workbook workbook)
        {
            int tries = 20;
            while (tries-- > 0)
            {
                try
                {
                    workbook.Close();
                    return true;
                }
                catch (COMException)
                {
                    Thread.Sleep(500);
                }
            }
            return false;
        }
    }
}
