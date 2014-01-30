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
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Word;


namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Word files
    /// </summary>
    class WordConverter: Converter
    {
        /// <summary>
        /// Convert a Word file to PDF
        /// </summary>
        /// <param name="inputFile">Full path of the input Word file</param>
        /// <param name="outputFile">Full path of the output PDF</param>
        /// <returns></returns>
        public static new Boolean Convert(String inputFile, String outputFile, Hashtable options)
        {
            Microsoft.Office.Interop.Word.Application word = null;
            object oMissing = System.Reflection.Missing.Value;
            try
            {
                word = new Microsoft.Office.Interop.Word.Application();
                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                word.DisplayRecentFiles = false;
                word.DisplayDocumentInformationPanel = false;
                word.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                Object filename = (Object)inputFile;
                Boolean visible = !(Boolean)options["hidden"];
                Boolean nowrite = (Boolean)options["readonly"];
                WdExportOptimizeFor quality = WdExportOptimizeFor.wdExportOptimizeForOnScreen;
                if ((Boolean)options["print"])
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                }
                WdExportCreateBookmarks bookmarks = (Boolean)options["bookmarks"] ? 
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks : 
                    WdExportCreateBookmarks.wdExportCreateNoBookmarks;
                Document doc = word.Documents.Open(ref filename, ref oMissing,
                        nowrite, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, visible,
                        true, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();
                if ((Boolean)options["hidden"])
                {
                    word.ActiveWindow.Visible = false;
                    word.ActiveWindow.WindowState = WdWindowState.wdWindowStateMinimize;
                }

                doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, false, 
                    quality, WdExportRange.wdExportAllDocument, 
                    1, 1, WdExportItem.wdExportDocumentContent, false, true, bookmarks);

                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;

                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            finally
            {
                if (word != null)
                {
                    ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                    word = null;
                }
            }
        }
    }
}
