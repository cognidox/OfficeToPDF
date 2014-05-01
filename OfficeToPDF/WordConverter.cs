/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010/2013
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
            Microsoft.Office.Interop.Word.Template tmpl;
            try
            {
                tmpl = null;
                word = new Microsoft.Office.Interop.Word.Application();
                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                word.DisplayRecentFiles = false;
                word.DisplayDocumentInformationPanel = false;
                word.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                Object filename = (Object)inputFile;
                Boolean visible = !(Boolean)options["hidden"];
                Boolean nowrite = (Boolean)options["readonly"];
                bool pdfa = (Boolean)options["pdfa"] ? true : false;
                WdExportOptimizeFor quality = WdExportOptimizeFor.wdExportOptimizeForOnScreen;
                if ((Boolean)options["print"])
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                }
                WdExportCreateBookmarks bookmarks = (Boolean)options["bookmarks"] ? 
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks : 
                    WdExportCreateBookmarks.wdExportCreateNoBookmarks;
                var documents = word.Documents;
                Document doc = documents.OpenNoRepairDialog(ref filename, ref oMissing,
                        nowrite, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, visible,
                        true, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();
                if ((Boolean)options["hidden"])
                {
                    var activeWin = word.ActiveWindow;
                    activeWin.Visible = false;
                    activeWin.WindowState = WdWindowState.wdWindowStateMinimize;
                    Converter.releaseCOMObject(activeWin);
                }

                // Check if we have a template file to apply to this document
                // The template must be a file and must end in .dot, .dotx or .dotm
                if (!String.IsNullOrEmpty((String)options["template"]))
                {
                    string template = (string)options["template"];
                    if (File.Exists(template) && System.Text.RegularExpressions.Regex.IsMatch(template, @"^.*\.dot[mx]?$"))
                    {
                        doc.set_AttachedTemplate(template);
                        doc.UpdateStyles();
                        tmpl = doc.get_AttachedTemplate();
                    }
                    else
                    {
                        Console.WriteLine("Invalid template '{0}'", template);
                    }
                }

                doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, false, 
                    quality, WdExportRange.wdExportAllDocument, 
                    1, 1, WdExportItem.wdExportDocumentContent, false, true, bookmarks, true, true, pdfa);

                if (tmpl != null)
                {
                    tmpl.Saved = true;
                }

                object saveChanges = nowrite ? WdSaveOptions.wdDoNotSaveChanges : WdSaveOptions.wdSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                Converter.releaseCOMObject(documents);
                Converter.releaseCOMObject(doc);
                Converter.releaseCOMObject(tmpl);

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
                }
                Converter.releaseCOMObject(word);
            }
        }
    }
}
