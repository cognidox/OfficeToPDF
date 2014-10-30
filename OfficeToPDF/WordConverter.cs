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
using System.Runtime.InteropServices;
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
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            Boolean running = (Boolean)options["noquit"];
            Microsoft.Office.Interop.Word.Application word = null;
            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Template tmpl;
            try
            {
                tmpl = null;
                try
                {
                    word = (Microsoft.Office.Interop.Word.Application) Marshal.GetActiveObject("Word.Application");
                }
                catch (System.Exception)
                {
                    word = new Microsoft.Office.Interop.Word.Application();
                    running = false;
                }
                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                word.DisplayRecentFiles = false;
                word.DisplayDocumentInformationPanel = false;
                word.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                var wdOptions = word.Options;
                wdOptions.UpdateFieldsAtPrint = false;
                wdOptions.UpdateLinksAtPrint = false;
                wdOptions.WarnBeforeSavingPrintingSendingMarkup = false;
                wdOptions.BackgroundSave = true;
                wdOptions.SavePropertiesPrompt = false;
                wdOptions.DoNotPromptForConvert = true;
                wdOptions.PromptUpdateStyle = false;
                wdOptions.ConfirmConversions = false;
                Object filename = (Object)inputFile;
                Boolean visible = !(Boolean)options["hidden"];
                Boolean nowrite = (Boolean)options["readonly"];
                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean includeTags = !(Boolean)options["excludetags"];
                bool pdfa = (Boolean)options["pdfa"] ? true : false;
                WdExportOptimizeFor quality = WdExportOptimizeFor.wdExportOptimizeForOnScreen;
                if ((Boolean)options["print"])
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                }
                WdExportCreateBookmarks bookmarks = (Boolean)options["bookmarks"] ? 
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks : 
                    WdExportCreateBookmarks.wdExportCreateNoBookmarks;
                WdExportItem showMarkup = WdExportItem.wdExportDocumentContent;
                if ((Boolean)options["markup"])
                {
                    showMarkup = WdExportItem.wdExportDocumentWithMarkup;
                }
                var documents = word.Documents;

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

                Document doc = null;
                try
                {
                    doc = documents.OpenNoRepairDialog(ref filename, ref oMissing,
                        nowrite, ref oMissing, ref oReadPass, ref oMissing, ref oMissing,
                        ref oWritePass, ref oMissing, ref oMissing, ref oMissing, visible,
                        true, ref oMissing, ref oMissing, ref oMissing);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    Console.WriteLine("Unable to open file");
                }
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

                // Update some of the field types in the document so the printed
                // PDF looks correct. Skips some field types (such as ASK) that would
                // create dialogs
                var fields = doc.Fields;
                for (int i = 1; i <= fields.Count; i++)
                {
                    switch (fields[i].Type)
                    {
                        case WdFieldType.wdFieldAuthor:
                        case WdFieldType.wdFieldAutoText:
                        case WdFieldType.wdFieldComments:
                        case WdFieldType.wdFieldCreateDate:
                        case WdFieldType.wdFieldDate:
                        case WdFieldType.wdFieldDocProperty:
                        case WdFieldType.wdFieldDocVariable:
                        case WdFieldType.wdFieldEditTime:
                        case WdFieldType.wdFieldFileName:
                        case WdFieldType.wdFieldFileSize:
                        case WdFieldType.wdFieldFootnoteRef:
                        case WdFieldType.wdFieldGreetingLine:
                        case WdFieldType.wdFieldIndex:
                        case WdFieldType.wdFieldInfo:
                        case WdFieldType.wdFieldKeyWord:
                        case WdFieldType.wdFieldLastSavedBy:
                        case WdFieldType.wdFieldNoteRef:
                        case WdFieldType.wdFieldNumChars:
                        case WdFieldType.wdFieldNumPages:
                        case WdFieldType.wdFieldNumWords:
                        case WdFieldType.wdFieldPage:
                        case WdFieldType.wdFieldPageRef:
                        case WdFieldType.wdFieldPrintDate:
                        case WdFieldType.wdFieldRef:
                        case WdFieldType.wdFieldRevisionNum:
                        case WdFieldType.wdFieldSaveDate:
                        case WdFieldType.wdFieldSection:
                        case WdFieldType.wdFieldSectionPages:
                        case WdFieldType.wdFieldSubject:
                        case WdFieldType.wdFieldTime:
                        case WdFieldType.wdFieldTitle:
                        case WdFieldType.wdFieldTOA:
                        case WdFieldType.wdFieldTOAEntry:
                        case WdFieldType.wdFieldTOC:
                        case WdFieldType.wdFieldTOCEntry:
                        case WdFieldType.wdFieldUserAddress:
                        case WdFieldType.wdFieldUserInitials:
                        case WdFieldType.wdFieldUserName:
                            fields[i].Update();
                            break;
                    }
                }
                doc.Saved = true;
                Converter.releaseCOMObject(fields);
                doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, false, 
                    quality, WdExportRange.wdExportAllDocument,
                    1, 1, showMarkup, includeProps, true, bookmarks, includeTags, true, pdfa);

                if (tmpl != null)
                {
                    tmpl.Saved = true;
                }

                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                if (nowrite)
                {
                    doc.Saved = true;
                }
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);

                Converter.releaseCOMObject(wdOptions);
                Converter.releaseCOMObject(documents);
                Converter.releaseCOMObject(doc);
                Converter.releaseCOMObject(tmpl);

                return (int)ExitCode.Success;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            finally
            {
                if (word != null && !running)
                {
                    ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                }
                Converter.releaseCOMObject(word);
            }
        }
    }
}
