/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010/2013
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
            String temporaryStorageDir = null;
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
                try
                {
                    wdOptions.UpdateFieldsAtPrint = false;
                    wdOptions.UpdateLinksAtPrint = false;
                    wdOptions.WarnBeforeSavingPrintingSendingMarkup = false;
                    wdOptions.BackgroundSave = true;
                    wdOptions.SavePropertiesPrompt = false;
                    wdOptions.DoNotPromptForConvert = true;
                    wdOptions.PromptUpdateStyle = false;
                    wdOptions.ConfirmConversions = false;
                    wdOptions.CheckGrammarAsYouType = false;
                    wdOptions.CheckGrammarWithSpelling = false;
                    wdOptions.CheckSpellingAsYouType = false;
                    wdOptions.DisplaySmartTagButtons = false;
                    wdOptions.EnableLivePreview = false;
                    wdOptions.ShowReadabilityStatistics = false;
                    wdOptions.SuggestSpellingCorrections = false;
                    wdOptions.AllowDragAndDrop = false;
                    wdOptions.EnableMisusedWordsDictionary = false;
                }
                catch (SystemException)
                {
                }
                Object filename = (Object)inputFile;
                Boolean visible = !(Boolean)options["hidden"];
                Boolean nowrite = (Boolean)options["readonly"];
                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean includeTags = !(Boolean)options["excludetags"];
                bool pdfa = (Boolean)options["pdfa"] ? true : false;
                WdExportOptimizeFor quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                if ((Boolean)options["print"])
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                }
                if ((Boolean)options["screen"])
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForOnScreen;
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
                var normalTemplate = word.NormalTemplate;

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
                    normalTemplate.Saved = true;
                    Console.WriteLine("Unable to open password protected file");
                    return (int)ExitCode.PasswordFailure;
                }

                Document doc = null;
                try
                {
                    if ((bool)options["merge"] && !String.IsNullOrEmpty((string)options["template"]) &&
                        File.Exists((string)options["template"]) &&
                        System.Text.RegularExpressions.Regex.IsMatch((string)options["template"], @"^.*\.dot[mx]?$"))
                    {
                        // Create a new document based on a template
                        doc = documents.Add((string)options["template"]);
                        Object rStart = 0;
                        Object rEnd = 0;
                        Range range = doc.Range(rStart, rEnd);
                        range.InsertFile(inputFile);
                        Converter.releaseCOMObject(range);
                        // Make sure we save the file with the original filename so 
                        // filename fields update correctly
                        temporaryStorageDir = Path.GetTempFileName();
                        File.Delete(temporaryStorageDir);
                        Directory.CreateDirectory(temporaryStorageDir);
                        doc.SaveAs(Path.Combine(temporaryStorageDir, Path.GetFileName(inputFile)));
                    }
                    else
                    {
                        // Open the source document
                        doc = documents.OpenNoRepairDialog(ref filename, ref oMissing,
                            nowrite, ref oMissing, ref oReadPass, ref oMissing, ref oMissing,
                            ref oWritePass, ref oMissing, ref oMissing, ref oMissing, visible,
                            true, ref oMissing, ref oMissing, ref oMissing);
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Console.WriteLine("Unable to open file");
                    return (int)ExitCode.FileOpenFailure;
                }
                doc.Activate();

                // Prevent "property not available" errors, see http://blogs.msmvps.com/wordmeister/2013/02/22/word2013bug-not-available-for-reading/
                var docWin = doc.ActiveWindow;
                var docWinView = docWin.View;
                docWinView.Type = WdViewType.wdPrintPreview;

                // Try to avoid Word thinking any changes are happening to the document
                doc.SpellingChecked = true;
                doc.GrammarChecked = true;

                // Changing these properties may be disallowed if the document is protected
                if (doc.ProtectionType == WdProtectionType.wdNoProtection)
                {
                    doc.TrackMoves = false;
                    doc.TrackRevisions = false;
                    doc.TrackFormatting = false;
                }
                normalTemplate.Saved = true;
                
                // Hide the document window if need be
                if ((Boolean)options["hidden"])
                {
                    var activeWin = word.ActiveWindow;
                    activeWin.Visible = false;
                    activeWin.WindowState = WdWindowState.wdWindowStateMinimize;
                    Converter.releaseCOMObject(activeWin);
                }

                // Check if we have a template file to apply to this document
                // The template must be a file and must end in .dot, .dotx or .dotm
                if (!String.IsNullOrEmpty((String)options["template"]) && !(bool)options["merge"])
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
                foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
                {
                    var sectionRange = section.Range;
                    var sectionFields = sectionRange.Fields;
                    foreach (Field sectionField in sectionFields)
                    {
                        WordConverter.updateField(sectionField, word, inputFile);
                    }

                    var headers = section.Headers;
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter header in headers)
                    {
                        var range = header.Range;
                        var rangeFields = range.Fields;
                        foreach (Field rangeField in rangeFields)
                        {
                            WordConverter.updateField(rangeField, word, inputFile);
                        }
                    }

                    var footers = section.Footers;
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter footer in footers)
                    {
                        var range = footer.Range;
                        var rangeFields = range.Fields;
                        foreach (Field rangeField in rangeFields)
                        {
                            WordConverter.updateField(rangeField, word, inputFile);
                        }
                    }
                }

                var docFields = doc.Fields;
                foreach (Field docField in docFields)
                {
                    WordConverter.updateField(docField, word, inputFile);
                }
                var storyRanges = doc.StoryRanges;
                foreach (Range range in storyRanges)
                {
                    var rangeFields = range.Fields;
                    foreach (Field field in rangeFields)
                    {
                        WordConverter.updateField(field, word, inputFile);
                    }
                }

                var pageSetup = doc.PageSetup;
                if ((float)options["word_header_dist"] >= 0)
                {
                    pageSetup.HeaderDistance = (float)options["word_header_dist"];
                }
                if ((float)options["word_footer_dist"] >= 0)
                {
                    pageSetup.FooterDistance = (float)options["word_footer_dist"];
                }

                normalTemplate.Saved = true;
                doc.Saved = true;
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
                normalTemplate.Saved = true;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);

                Converter.releaseCOMObject(pageSetup);
                Converter.releaseCOMObject(docWinView);
                Converter.releaseCOMObject(docWin);
                Converter.releaseCOMObject(wdOptions);
                Converter.releaseCOMObject(documents);
                Converter.releaseCOMObject(doc);
                Converter.releaseCOMObject(tmpl);
                Converter.releaseCOMObject(normalTemplate);

                return (int)ExitCode.Success;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            finally
            {
                if (temporaryStorageDir != null && Directory.Exists(temporaryStorageDir))
                {
                    if (File.Exists(Path.Combine(temporaryStorageDir, Path.GetFileName(inputFile))))
                    {
                        File.Delete(Path.Combine(temporaryStorageDir, Path.GetFileName(inputFile)));
                    }
                    Directory.Delete(temporaryStorageDir);
                }
                if (word != null && !running)
                {
                    ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                }
                Converter.releaseCOMObject(word);
            }
        }
        private static void updateField(Field field, Microsoft.Office.Interop.Word.Application word, String filename)
        {
            switch (field.Type)
            {
                case WdFieldType.wdFieldAuthor:
                case WdFieldType.wdFieldAutoText:
                case WdFieldType.wdFieldComments:
                case WdFieldType.wdFieldCreateDate:
                case WdFieldType.wdFieldDate:
                case WdFieldType.wdFieldDocProperty:
                case WdFieldType.wdFieldDocVariable:
                case WdFieldType.wdFieldEditTime:
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
                    field.Update();
                    break;
                case WdFieldType.wdFieldFileName:
                    // Handle the filename as a special situation, since it doesn't seem to
                    // update correctly (issue #13)
                    field.Select();
                    field.Delete();
                    Selection selection = word.Selection;
                    selection.TypeText(Path.GetFileName(filename));
                    Converter.releaseCOMObject(selection);
                    break;
            }
        }
    }
}
