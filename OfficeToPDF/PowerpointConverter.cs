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
using System.Collections.Generic;
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
            if (IsPasswordProtected(inputFile))
            {
                Console.WriteLine("Unable to open password protected file");
                return (int)ExitCode.PasswordFailure;
            }
            
            Boolean running = (Boolean)options["noquit"];
            try
            {
                Microsoft.Office.Interop.PowerPoint.Application app = null;
                Presentation activePresentation = null;
                Presentations presentations = null;
                try
                {
                    try
                    {
                        app = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch (System.Exception)
                    {
                        int tries = 10;
                        // Create the application
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
                            ReleaseCOMObject(app);
                            return (int)ExitCode.ApplicationError;
                        }
                    }
                    Boolean includeProps = !(Boolean)options["excludeprops"];
                    Boolean includeTags = !(Boolean)options["excludetags"];
                    PpPrintOutputType printType = PpPrintOutputType.ppPrintOutputSlides;
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
                    if (!String.IsNullOrWhiteSpace((String)options["powerpoint_output"]))
                    {
                        bool printIsValid = false;
                        printType = GetOutputType((String)options["powerpoint_output"], ref printIsValid);
                    }

                    // Powerpoint files can be protected by a write password, but there's no way
                    // of opening them in 
                    if (nowrite == MSCore.MsoTriState.msoFalse && IsReadOnlyEnforced(inputFile))
                    {
                        // Seems like PowerPoint interop will ignore the read-only option
                        // when it is opening a document with a write password and still pop
                        // up a password input dialog. To prevent this freezing, don't open
                        // the file
                        throw new Exception("Presentation has a write password - this prevents it being opened");
                    }
                    
                    app.FeatureInstall = MSCore.MsoFeatureInstall.msoFeatureInstallNone;
                    app.DisplayDocumentInformationPanel = false;
                    app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
                    app.Visible = MSCore.MsoTriState.msoTrue;
                    app.AutomationSecurity = MSCore.MsoAutomationSecurity.msoAutomationSecurityLow;
                    presentations = app.Presentations;
                    String filenameWithPasswords = inputFile;
                    if (!String.IsNullOrWhiteSpace((string)options["password"]) ||
                        !String.IsNullOrWhiteSpace((string)options["writepassword"]))
                    {
                        // seems we can use the passwords by appending them to the file name!
                        filenameWithPasswords = String.Format("{0}::{1}::{2}", inputFile, 
                            (String.IsNullOrEmpty((string)options["password"]) ? "" : (string)options["password"]),
                            (String.IsNullOrEmpty((string)options["writepassword"]) ? "" : (string)options["writepassword"]));
                        Console.WriteLine(filenameWithPasswords);
                    }
                    activePresentation = presentations.Open2007(FileName: filenameWithPasswords, ReadOnly: nowrite, Untitled: MSCore.MsoTriState.msoTrue, OpenAndRepair: MSCore.MsoTriState.msoTrue);
                    var changeLoop = 0;
                    while (changeLoop++ < 10)
                    {
                        // Try and wait for the presentation to become usable
                        try
                        {
                            activePresentation.Final = false;
                            break;
                        }
                        catch (Exception)
                        {
                            Thread.Sleep(500);
                        }
                    }
                    
                    // Sometimes, presentations can have restrictions on them that block
                    // access to the object model (e.g. fonts containing restrictions).
                    // If we attempt to access the object model and fail, then try a more
                    // sneaky method of getting the presentation - create an empty presentation
                    // and insert the slides from the original file.
                    Fonts fonts = null;
                    Boolean bitmapMissingFonts = !(Boolean)options["powerpoint_ref_fonts"];
                    try
                    {
                        fonts = activePresentation.Fonts;
                        var fontCount = fonts.Count;
                    }
                    catch (COMException)
                    {
                        ReleaseCOMObject(fonts);
                        // This presentation looked read-only
                        ClosePowerPointPresentation(activePresentation);
                        ReleaseCOMObject(activePresentation);
                        // Create a new blank presentation and insert slides from the original
                        activePresentation = presentations.Add(MSCore.MsoTriState.msoFalse);
                        // This is only a band-aid - backgrounds won't come through
                        activePresentation.Slides.InsertFromFile(inputFile, 0);
                    }
                    ReleaseCOMObject(fonts);
                    
                    // Set up a delegate function for times we want to print
                    PrintDocument printFunc = delegate (string destination, string printer)
                    {
                        PrintOptions activePrintOptions = activePresentation.PrintOptions;
                        activePrintOptions.PrintInBackground = MSCore.MsoTriState.msoFalse;
                        activePrintOptions.ActivePrinter = printer;
                        activePrintOptions.PrintInBackground = MSCore.MsoTriState.msoFalse;
                        activePresentation.PrintOut(PrintToFile: destination, Copies: 1);
                        ReleaseCOMObject(activePrintOptions);
                    };
                    
                    if (String.IsNullOrEmpty((string)options["printer"]))
                    {
                        try
                        {
                            activePresentation.ExportAsFixedFormat(outputFile, PpFixedFormatType.ppFixedFormatTypePDF, quality, MSCore.MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, printType, MSCore.MsoTriState.msoFalse, null, PpPrintRangeType.ppPrintAll, "", includeProps, true, includeTags, bitmapMissingFonts, pdfa, Type.Missing);
                        }
                        catch (Exception) {
                            if (!String.IsNullOrEmpty((string)options["fallback_printer"])) {
                                PrintToGhostscript((string)options["fallback_printer"], outputFile, printFunc);
                            } else {
                                throw;
                            }
                        }
                        finally
                        {
                            ReleaseCOMObject(printType);
                            ReleaseCOMObject(quality);
                        }
                    } else
                    {
                        // Print via a delegate
                        PrintToGhostscript((string)options["printer"], outputFile, printFunc);
                    }
                    ReleaseCOMObject(printType);
                    ReleaseCOMObject(quality);

                    // Determine if we need to make bookmarks
                    if ((bool)options["bookmarks"])
                    {
                        LoadBookmarks(activePresentation, ref bookmarks);

                    }
                    activePresentation.Saved = MSCore.MsoTriState.msoTrue;
                    ClosePowerPointPresentation(activePresentation);

                    return (int)ExitCode.Success;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return (int)ExitCode.UnknownError;
                }
                finally
                {
                    ReleaseCOMObject(activePresentation);
                    ReleaseCOMObject(presentations);

                    if (app != null && !running)
                    {
                        app.Quit();
                    }
                    ReleaseCOMObject(app);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (COMException e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
        }

        // Try and close PowerPoint presentation, giving time for Office to get
        // itself in order
        private static bool ClosePowerPointPresentation(Presentation presentation)
        {
            int tries = 20;
            while (tries-- > 0)
            {
                try
                {
                    presentation.Close();
                    return true;
                }
                catch (Exception)
                {
                    Thread.Sleep(500);
                }
            }
            return false;
        }

        // Loop through all the slides in the presentation creating bookmark items
        // for all the slides that are not hidden
        private static void LoadBookmarks(Presentation activePresentation, ref List<PDFBookmark> bookmarks)
        {
            var slides = activePresentation.Slides;
            if (slides.Count > 0)
            {
                var page = 1;

                // Create a top-level bookmark
                var parentBookmark = new PDFBookmark
                {
                    title = activePresentation.Name,
                    page = 1,
                    children = new List<PDFBookmark>()
                };

                // Loop through the slides, adding a ToC entry to the top-level bookmark
                for (int sldIdx = 1; sldIdx <= slides.Count; sldIdx++)
                {
                    var s = slides[sldIdx];
                    // Look at the transition on the slide to determine if it is hidden
                    var trans = ((Slide)s).SlideShowTransition;
                    if (trans.Hidden == MSCore.MsoTriState.msoCTrue || trans.Hidden == MSCore.MsoTriState.msoTrue)
                    {
                        ReleaseCOMObject(trans);
                        ReleaseCOMObject(s);
                        continue;
                    }
                    ReleaseCOMObject(trans);

                    // Create a new bookmark and add the page
                    var bookmark = new PDFBookmark
                    {
                        page = page++
                    };

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
                                ReleaseCOMObject(textrange);
                            }
                            ReleaseCOMObject(textframe);
                        }
                        ReleaseCOMObject(shapeTitle);
                    }
                    ReleaseCOMObject(shapes);

                    bookmark.title = String.Format("Page {0} - {1}", bookmark.page, slideName);

                    // Put the bookmark into our parent bookmark children
                    parentBookmark.children.Add(bookmark);

                    // Clean up the references to the slide
                    ReleaseCOMObject(s);
                }
                // Add the top-level bookmark which will be passed back to the main program
                bookmarks.Add(parentBookmark);
            }
            ReleaseCOMObject(slides);
        }

        
        // Return the PpPrintOutputType for a given option string
        public static PpPrintOutputType GetOutputType(string printType, ref bool valid)
        {
            valid = true;
            switch (printType)
            {
                case "handout":
                case "handouts":
                case "handout1":
                    return PpPrintOutputType.ppPrintOutputOneSlideHandouts;
                case "handout2":
                case "handouts2":
                    return PpPrintOutputType.ppPrintOutputTwoSlideHandouts;
                case "handout3":
                case "handouts3":
                    return PpPrintOutputType.ppPrintOutputThreeSlideHandouts;
                case "handout4":
                case "handouts4":
                    return PpPrintOutputType.ppPrintOutputFourSlideHandouts;
                case "handout6":
                case "handouts6":
                    return PpPrintOutputType.ppPrintOutputSixSlideHandouts;
                case "handout9":
                case "handouts9":
                    return PpPrintOutputType.ppPrintOutputNineSlideHandouts;
                case "notes":
                    return PpPrintOutputType.ppPrintOutputNotesPages;
                case "slides":
                    return PpPrintOutputType.ppPrintOutputSlides;
                case "outline":
                    return PpPrintOutputType.ppPrintOutputOutline;
                case "build_slides":
                case "buildslides":
                case "build-slides":
                    return PpPrintOutputType.ppPrintOutputBuildSlides;
                default:
                    valid = false;
                    break;
            }
            return PpPrintOutputType.ppPrintOutputSlides;
        }
    }
}
