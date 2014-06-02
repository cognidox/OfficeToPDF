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
using MSCore = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Powerpoint files
    /// </summary>
    class PowerpointConverter : Converter
    {
        public static new Boolean Convert(String inputFile, String outputFile, Hashtable options)
        {
            try
            {
                Microsoft.Office.Interop.PowerPoint.Application app = null;
                try
                {
                    MSCore.MsoTriState nowrite = (Boolean)options["readonly"] ? MSCore.MsoTriState.msoTrue : MSCore.MsoTriState.msoFalse;
                    app = new Microsoft.Office.Interop.PowerPoint.Application();
                    bool pdfa = (Boolean)options["pdfa"] ? true : false;
                    if ((Boolean)options["hidden"])
                    {
                        // Can't really hide the window, so at least minimise it
                        app.WindowState = PpWindowState.ppWindowMinimized;
                    }
                    PpFixedFormatIntent quality = PpFixedFormatIntent.ppFixedFormatIntentScreen;
                    if ((Boolean)options["print"])
                    {
                        quality = PpFixedFormatIntent.ppFixedFormatIntentPrint;
                    }
                    app.FeatureInstall = MSCore.MsoFeatureInstall.msoFeatureInstallNone;
                    app.DisplayDocumentInformationPanel = false;
                    app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
                    app.Visible = MSCore.MsoTriState.msoTrue;
                    app.AutomationSecurity = MSCore.MsoAutomationSecurity.msoAutomationSecurityLow;
                    var presentations = app.Presentations;
                    presentations.Open2007(inputFile, nowrite, MSCore.MsoTriState.msoTrue, MSCore.MsoTriState.msoTrue, MSCore.MsoTriState.msoTrue);
                    var activePresentation = app.ActivePresentation;
                    activePresentation.ExportAsFixedFormat(outputFile, PpFixedFormatType.ppFixedFormatTypePDF, quality, MSCore.MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType.ppPrintOutputSlides, MSCore.MsoTriState.msoFalse, null, PpPrintRangeType.ppPrintAll, "", false, true, true, true, pdfa, Type.Missing);
                    activePresentation.Saved = MSCore.MsoTriState.msoTrue;
                    activePresentation.Close();

                    Converter.releaseCOMObject(activePresentation);
                    Converter.releaseCOMObject(presentations);
                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return false;
                }
                finally
                {
                    if (app != null)
                    {
                        app.Quit();
                    }
                    Converter.releaseCOMObject(app);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }
    }
}
