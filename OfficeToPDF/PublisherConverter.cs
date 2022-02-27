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
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Publisher;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Microsoft Publisher files
    /// </summary>
    class PublisherConverter: Converter, IConverter
    {
        int IConverter.Convert(String inputFile, String outputFile, ArgParser options, ref List<PDFBookmark> bookmarks)
        {
            if (options.verbose)
            {
                Console.WriteLine("Converting with Publisher converter");
            }
            return Convert(inputFile, outputFile, options, ref bookmarks);
        }

        public static ExitCode StartPublisher(ref Boolean running, ref Application publisher)
        {
            try
            {
                publisher = (Microsoft.Office.Interop.Publisher.Application)Marshal.GetActiveObject("Publisher.Application");
            }
            catch (System.Exception)
            {
                publisher = new Microsoft.Office.Interop.Publisher.Application();
                running = false;
            }
            return ExitCode.Success;
        }

        static int Convert(String inputFile, String outputFile, ArgParser options, ref List<PDFBookmark> bookmarks)
        {
            Boolean running = options.noquit;
            Microsoft.Office.Interop.Publisher.Application publisher = null;
            String tmpFile = null;
            IWatchdog watchdog = new NullWatchdog();
            try
            {
                ExitCode result = StartPublisher(ref running, ref publisher);
                if (result != ExitCode.Success)
                    return (int)result;

                watchdog = WatchdogFactory.CreateStarted(publisher, options.timeout);

                Boolean nowrite = options.@readonly;
                bool pdfa = options.pdfa;
                if (options.hidden)
                {
                    var activeWin = publisher.ActiveWindow;
                    activeWin.Visible = false;
                    ReleaseCOMObject(activeWin);
                }
                publisher.Open(inputFile, nowrite, false, PbSaveOptions.pbDoNotSaveChanges);
                PbFixedFormatIntent quality = PbFixedFormatIntent.pbIntentStandard;
                if (options.print)
                {
                    quality = PbFixedFormatIntent.pbIntentPrinting;
                }
                if (options.screen)
                {
                    quality = PbFixedFormatIntent.pbIntentMinimum;
                }
                Boolean includeProps = !options.excludeprops;
                Boolean includeTags = !options.excludetags;

                // Try and avoid dialogs about versions
                tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".pub";
                var activeDocument = publisher.ActiveDocument;
                activeDocument.SaveAs(tmpFile, PbFileFormat.pbFilePublication, false);
                activeDocument.ExportAsFixedFormat(PbFixedFormatType.pbFixedFormatTypePDF, outputFile, quality, includeProps, -1, -1, -1, -1, -1, -1, -1, true, PbPrintStyle.pbPrintStyleDefault, includeTags, true, pdfa);

                // Determine if we need to make bookmarks
                if (options.bookmarks)
                {
                    LoadBookmarks(activeDocument, ref bookmarks, options);
                }
                activeDocument.Close();

                ReleaseCOMObject(activeDocument);
                return (int)ExitCode.Success;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            finally
            {
                watchdog.Stop();

                if (tmpFile != null)
                {
                    System.IO.File.Delete(tmpFile);
                }
                if (publisher != null)
                {
                    ClosePublisherApplication(publisher);
                }
                ReleaseCOMObject(publisher);
            }
        }

        internal static void ClosePublisherApplication(Application publisher)
        {
            ((Microsoft.Office.Interop.Publisher._Application)publisher).Quit();
        }

        // Loop through all the pages in the document creating bookmark items for them
        private static void LoadBookmarks(Document activeDocument, ref List<PDFBookmark> bookmarks, ArgParser options)
        {
            var pages = activeDocument.Pages;
            if (pages.Count > 0)
            {
                // Create a top-level bookmark
                var parentBookmark = new PDFBookmark();
                parentBookmark.title = options.original_basename;
                parentBookmark.page = 1;
                parentBookmark.children = new List<PDFBookmark>();

                foreach (var p in pages)
                {
                    var bookmark = new PDFBookmark();
                    bookmark.page = ((Page)p).PageIndex;
                    bookmark.title = ((Page)p).Name;
                    parentBookmark.children.Add(bookmark);
                    ReleaseCOMObject(p);

                }
                ReleaseCOMObject(pages);
                bookmarks.Add(parentBookmark);
            }
        }
    }
}
