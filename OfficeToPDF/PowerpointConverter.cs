/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010
 *  Copyright (C) 2011-2015 Cognidox Ltd
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
using System.Reflection;
using System.Runtime.InteropServices;
using MSCore = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Powerpoint files
    /// </summary>
    class PowerpointConverter : Converter
    {
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            // Check for password protection
            if (Converter.IsPasswordProtected(inputFile))
            {
                Console.WriteLine("Unable to open password protected file");
                return (int)ExitCode.PasswordFailure;
            }

            Boolean running = (Boolean)options["noquit"];
            try
            {
                Microsoft.Office.Interop.PowerPoint.Application app = null;
                Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = null;
                Microsoft.Office.Interop.PowerPoint.Presentations presentations = null;
                try
                {
                    try
                    {
                        app = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch (System.Exception)
                    {
                        app = new Microsoft.Office.Interop.PowerPoint.Application();
                        running = false;
                    }
                    MSCore.MsoTriState nowrite = (Boolean)options["readonly"] ? MSCore.MsoTriState.msoTrue : MSCore.MsoTriState.msoFalse;
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
                    if ((Boolean)options["screen"])
                    {
                        quality = PpFixedFormatIntent.ppFixedFormatIntentScreen;
                    }
                    Boolean includeProps = !(Boolean)options["excludeprops"];
                    Boolean includeTags = !(Boolean)options["excludetags"];
                    app.FeatureInstall = MSCore.MsoFeatureInstall.msoFeatureInstallNone;
                    app.DisplayDocumentInformationPanel = false;
                    app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
                    app.Visible = MSCore.MsoTriState.msoTrue;
                    app.AutomationSecurity = MSCore.MsoAutomationSecurity.msoAutomationSecurityLow;
                    presentations = app.Presentations;
                    activePresentation = presentations.Open2007(inputFile, nowrite, MSCore.MsoTriState.msoTrue, MSCore.MsoTriState.msoTrue, MSCore.MsoTriState.msoTrue);
                    activePresentation.Final = false;

                    // Sometimes, presentations can have restrictions on them that block
                    // access to the object model (e.g. fonts containing restrictions).
                    // If we attempt to access the object model and fail, then try a more
                    // sneaky method of getting the presentation - create an empty presentation
                    // and insert the slides from the original file.
                    var fonts = activePresentation.Fonts;
                    try
                    {
                        var fontCount = fonts.Count;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        Converter.releaseCOMObject(fonts);
                        // This presentation looked read-only
                        activePresentation.Close();
                        Converter.releaseCOMObject(activePresentation);
                        // Create a new blank presentation and insert slides from the original
                        activePresentation = presentations.Add(MSCore.MsoTriState.msoFalse);
                        // This is only a band-aid - backgrounds won't come through
                        activePresentation.Slides.InsertFromFile(inputFile, 0);
                    }
                    Converter.releaseCOMObject(fonts);

                    // See if there is a macro to run
                    if (!String.IsNullOrWhiteSpace((String)options["powerpoint_run_macro"]))
                    {
                        Console.WriteLine("Running macro " + (String)options["powerpoint_run_macro"]);
                        try
                        {
                            System.Threading.Thread.Sleep(2000);
                            using (new ChangeLocalHelper("en-us"))
                            {
                                app.Run((String)options["powerpoint_run_macro"], null);
                            }
                            //PowerpointConverter.RunMacro((Microsoft.Office.Interop.PowerPoint._Application)app, new Object[] { (String)options["powerpoint_run_macro"] });
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Failed to run macro");
                            Console.WriteLine(e.Message);
                        }
                    }

                    activePresentation.ExportAsFixedFormat(outputFile, PpFixedFormatType.ppFixedFormatTypePDF, quality, MSCore.MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType.ppPrintOutputSlides, MSCore.MsoTriState.msoFalse, null, PpPrintRangeType.ppPrintAll, "", includeProps, true, includeTags, true, pdfa, Type.Missing);
                    activePresentation.Saved = MSCore.MsoTriState.msoTrue;
                    activePresentation.Close();

                    return (int)ExitCode.Success;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return (int)ExitCode.UnknownError;
                }
                finally
                {
                    Converter.releaseCOMObject(activePresentation);
                    Converter.releaseCOMObject(presentations);

                    if (app != null && !running)
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
                return (int)ExitCode.UnknownError;
            }
        }

        private static void RunMacro(object oApp, object[] oRunArgs)
        {
            try
            {
                using (new ChangeLocalHelper("en-us")) {
                    oApp.GetType().InvokeMember("Run",
                        System.Reflection.BindingFlags.Default |
                        System.Reflection.BindingFlags.InvokeMethod,
                        null, oApp, oRunArgs);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Fail macro " + e.Message);
            }
        }
    }
    class ChangeLocalHelper : IDisposable
    {
        private string _localeName;
        private string _originalLocale;
        public ChangeLocalHelper(string localeName)
        {
            this._localeName = localeName;
            _originalLocale = System.Threading.Thread.CurrentThread.CurrentCulture.Name;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(this._localeName);
        }

        #region IDisposable Members
        public void Dispose()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(this._originalLocale);
        }
        #endregion
    }

}
