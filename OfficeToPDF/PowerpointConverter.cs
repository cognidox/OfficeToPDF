/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010/2013/2016
 *  Copyright (C) 2011-2016 Cognidox Ltd
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
using System.Threading;
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
            List<PDFBookmark> bookmarks = new List<PDFBookmark>();
            return Convert(inputFile, outputFile, options, ref bookmarks);
        }

        public static new int Convert(String inputFile, String outputFile, Hashtable options, ref List<PDFBookmark> bookmarks)
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
                        int tries = 10;
                        app = new Microsoft.Office.Interop.PowerPoint.Application();
                        running = false;
                        while (tries > 0)
                        {
                            try
                            {
                                // Try to set a property on the object
                                app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
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
                            Converter.releaseCOMObject(app);
                            return (int)ExitCode.ApplicationError;
                        }
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
                    activePresentation.ExportAsFixedFormat(outputFile, PpFixedFormatType.ppFixedFormatTypePDF, quality, MSCore.MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType.ppPrintOutputSlides, MSCore.MsoTriState.msoFalse, null, PpPrintRangeType.ppPrintAll, "", includeProps, true, includeTags, true, pdfa, Type.Missing);

                    // Determine if we need to make bookmarks
                    if ((bool)options["bookmarks"])
                    {
                        loadBookmarks(activePresentation, ref bookmarks);
                        
                    }
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

        // Loop through all the slides in the presentation creating bookmark items
        // for all the slides that are not hidden
        private static void loadBookmarks(Presentation activePresentation, ref List<PDFBookmark> bookmarks)
        {
            var slides = activePresentation.Slides;
            if (slides.Count > 0)
            {
                var page = 1;

                // Create a top-level bookmark
                var parentBookmark = new PDFBookmark();
                parentBookmark.title = activePresentation.Name;
                parentBookmark.page = 1;
                parentBookmark.children = new List<PDFBookmark>();

                // Loop through the slides, adding a ToC entry to the top-level bookmark
                foreach (var s in slides)
                {
                    // Look at the transition on the slide to determine if it is hidden
                    var trans = ((Slide)s).SlideShowTransition;
                    if (trans.Hidden == MSCore.MsoTriState.msoCTrue || trans.Hidden == MSCore.MsoTriState.msoTrue)
                    {
                        Converter.releaseCOMObject(trans);
                        Converter.releaseCOMObject(s);
                        continue;
                    }
                    Converter.releaseCOMObject(trans);

                    // Create a new bookmark and add the page
                    var bookmark = new PDFBookmark();
                    bookmark.page = page++;

                    // Work out a title - base this on the slide name and any title shape text
                    var slideName = ((Slide)s).Name;
                    var shapes = ((Slide)s).Shapes;

                    // See if there is a title in the slides shapes
                    if (shapes.HasTitle == MSCore.MsoTriState.msoTrue || shapes.HasTitle == MSCore.MsoTriState.msoCTrue)
                    {
                        var shapeTitle = shapes.Title;
                        if (shapeTitle != null && (shapeTitle.HasTextFrame == MSCore.MsoTriState.msoCTrue || shapeTitle.HasTextFrame == MSCore.MsoTriState.msoTrue))
                        {
                            var textframe = shapeTitle.TextFrame;
                            if (textframe != null && (textframe.HasText == MSCore.MsoTriState.msoTrue || textframe.HasText == MSCore.MsoTriState.msoCTrue))
                            {
                                var textrange = textframe.TextRange;
                                if (!String.IsNullOrWhiteSpace(textrange.TrimText().Text))
                                {
                                    slideName = textrange.TrimText().Text;
                                }
                                Converter.releaseCOMObject(textrange);
                            }
                            Converter.releaseCOMObject(textframe);
                        }
                        Converter.releaseCOMObject(shapeTitle);
                    }
                    Converter.releaseCOMObject(shapes);

                    bookmark.title = String.Format("Page {0} - {1}", bookmark.page, slideName);

                    // Put the bookmark into our parent bookmark children
                    parentBookmark.children.Add(bookmark);

                    // Clean up the references to the slide
                    Converter.releaseCOMObject(s);
                }
                // Add the top-level bookmark which will be passed back to the main program
                bookmarks.Add(parentBookmark);
            }
            Converter.releaseCOMObject(slides);
        }
    }
}
