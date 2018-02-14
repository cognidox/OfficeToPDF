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
using System.Linq;
using System.Text;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
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
            float wordVersion = 0;
            List<AppOption> wordOptionList = new List<AppOption>();
            try
            {
                tmpl = null;
                try
                {
                    word = (Microsoft.Office.Interop.Word.Application) Marshal.GetActiveObject("Word.Application");
                }
                catch (System.Exception)
                {
                     int tries = 10;
                     word = new Microsoft.Office.Interop.Word.Application();
                     running = false;
                     while (tries > 0)
                     {
                        try
                        {
                            // Try to set a property on the object
                            word.ScreenUpdating = false;
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
                        Converter.releaseCOMObject(word);
                        return (int)ExitCode.ApplicationError;
                    }
                }
                
                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                // Issue #48 - we should allow control over whether the history is lost
                if (!(Boolean)options["word_keep_history"])
                {
                    word.DisplayRecentFiles = false;
                }
                word.DisplayDocumentInformationPanel = false;
                word.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                wordVersion = (float)System.Convert.ToDecimal(word.Version, new CultureInfo("en-US"));
                var wdOptions = word.Options;
                try
                {
                    // Set the Word options in a way that allows us to reset the options when we finish
                    wordOptionList.Add(new AppOption("AlertIfNotDefault", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("AllowReadingMode", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("PrecisePositioning", true, ref wdOptions));
                    wordOptionList.Add(new AppOption("UpdateFieldsAtPrint", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("UpdateLinksAtPrint", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("WarnBeforeSavingPrintingSendingMarkup", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("BackgroundSave", true, ref wdOptions));
                    wordOptionList.Add(new AppOption("SavePropertiesPrompt", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("DoNotPromptForConvert", true, ref wdOptions));
                    wordOptionList.Add(new AppOption("PromptUpdateStyle", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("ConfirmConversions", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("CheckGrammarAsYouType", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("CheckGrammarWithSpelling", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("CheckSpellingAsYouType", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("DisplaySmartTagButtons", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("EnableLivePreview", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("ShowReadabilityStatistics", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("SuggestSpellingCorrections", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("AllowDragAndDrop", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("EnableMisusedWordsDictionary", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("ShowFormatError", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("StoreRSIDOnSave", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("SaveNormalPrompt", false, ref wdOptions));
                }
                catch (SystemException)
                {
                }
                
                Object filename = (Object)inputFile;
                Boolean hasSignatures = false;
                Boolean visible = !(Boolean)options["hidden"];
                Boolean nowrite = (Boolean)options["readonly"];
                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean includeTags = !(Boolean)options["excludetags"];
                Boolean bitmapMissingFonts = !(Boolean)options["word_ref_fonts"];
                Boolean autosave = options.ContainsKey("IsTempWord") && (Boolean)options["IsTempWord"];
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

                // Large Word files may simply not print reliably - if the word_max_pages
                // configuration option is set, then we must close up and forget about 
                // converting the file.
                var maxPages = (int)options[@"word_max_pages"];

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

                // Check if there are signatures in the document which 
                var signatures = doc.Signatures;
                if (signatures.Count > 0)
                {
                    signatures.ShowSignaturesPane = false;
                    nowrite = true;
                    autosave = false;
                    hasSignatures = true;
                    options["word_no_field_update"] = true;
                }
                Converter.releaseCOMObject(signatures);

                // Check if there are too many pages
                if (maxPages > 0)
                {
                    var pageCount = doc.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                    doc.Saved = true;
                    if (pageCount > maxPages)
                    {
                        throw new Exception(String.Format("Too many pages to process ({0}). More than {1}", pageCount, maxPages));
                    }
                }
                
                // Prevent "property not available" errors, see http://blogs.msmvps.com/wordmeister/2013/02/22/word2013bug-not-available-for-reading/
                var docWin = doc.ActiveWindow;
                var docWinView = docWin.View;
                if (wordVersion >= 15)
                {
                    docWinView.ReadingLayout = false;
                }

                // Sometimes the print view will not be available (e.g. for a blog post)
                // Try and switch view
                try
                {
                    docWinView.Type = WdViewType.wdPrintPreview;
                }
                catch(Exception){}

                // Hide comments
                try
                {
                    docWinView.RevisionsView = WdRevisionsView.wdRevisionsViewFinal;
                    docWinView.ShowRevisionsAndComments = false;
                }
                catch(SystemException){}

                // Try to avoid Word thinking any changes are happening to the document
                doc.SpellingChecked = true;
                doc.GrammarChecked = true;

                // Changing these properties may be disallowed if the document is protected
                // and is not signed
                if (doc.ProtectionType == WdProtectionType.wdNoProtection && !hasSignatures)
                {
                    if (autosave) { doc.Save(); doc.Saved = true; }
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

                // See if we have to update fields
                if (!(Boolean)options["word_no_field_update"])
                {
                    updateDocumentFields(doc, word, inputFile, options);
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

                try
                {
                    // Make sure we are not in a header footer view
                    docWinView.SeekView = WdSeekView.wdSeekPrimaryHeader;
                    docWinView.SeekView = WdSeekView.wdSeekPrimaryFooter;
                    docWinView.SeekView = WdSeekView.wdSeekMainDocument;
                }
                catch(Exception)
                {
                    // We might fail when switching views
                }

                normalTemplate.Saved = true;
                if (autosave)
                {
                    doc.Save();
                }
                doc.Saved = true;

                // If a document has signatures, Word can take a bit of time to
                // validate them and get itself ready to do the conversion. Here
                // we pause a bit to give Word time to get itself ready.
                if (hasSignatures)
                {
                    // Bit of a hack really
                    Thread.Sleep(2500);
                }
                doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, false,
                        quality, WdExportRange.wdExportAllDocument,
                        1, 1, showMarkup, includeProps, true, bookmarks, includeTags, bitmapMissingFonts, pdfa);
                
                if (tmpl != null)
                {
                    tmpl.Saved = true;
                }

                object saveChanges = autosave? WdSaveOptions.wdSaveChanges : WdSaveOptions.wdDoNotSaveChanges;
                if (nowrite)
                {
                    doc.Saved = true;
                }
                normalTemplate.Saved = true;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);

                // Reset options
                foreach (AppOption opt in wordOptionList)
                {
                    opt.resetValue(ref wdOptions);
                }

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
                    try
                    {
                        if (File.Exists(Path.Combine(temporaryStorageDir, Path.GetFileName(inputFile))))
                        {
                            File.Delete(Path.Combine(temporaryStorageDir, Path.GetFileName(inputFile)));
                        }
                        Directory.Delete(temporaryStorageDir);
                    }
                    catch (Exception) { }
                }
                if (word != null && !running)
                {
                    closeWordApplication(word);
                }
                Converter.releaseCOMObject(word);
            }
        }

        // Try and close Word, giving time for Office to get
        // itself in order
        private static bool closeWordApplication(Application word)
        {
            object oMissing = System.Reflection.Missing.Value;
            int tries = 20;
            while (tries-- > 0)
            {
                try
                {
                    ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                    return true;
                }
                catch (COMException)
                {
                    Thread.Sleep(500);
                }
            }
            return false;
        }
        // We want to be able to reset the options in Word so it doesn't affect subsequent
        // usage
        private class AppOption
        {
            public string name { get; set; }
            public Boolean value { get; set; }
            public Boolean originalValue { get; set; }
            public AppOption(string name, Boolean value, ref Options wdOptions)
            {
                try
                {
                    this.name = name;
                    this.value = value;
                    this.originalValue = (Boolean)wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, wdOptions, null);

                    if (this.originalValue != value)
                    {
                        wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] {value});
                    }
                }
                catch
                {
                    // We may be setting word options that are not available in the version of word
                    // being used, so just skip these errors
                }
            }

            // Allow the value on the options to be reset
            public void resetValue(ref Options wdOptions)
            {
                if (this.value != this.originalValue)
                {
                    wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { this.originalValue });
                }
            }
        }

        // Update all the fields in a document
        private static void updateDocumentFields(Document doc, Microsoft.Office.Interop.Word.Application word, String inputFile, Hashtable options)
        {
            // Update fields quickly if it is safe to do so. We have
            // to check for broken links as they may raise Word dialogs or leave broken content
            if ((Boolean)options["word_field_quick_update"] ||
                ((Boolean)options["word_field_quick_update_safe"] && !hasBrokenLinks(doc)))
            {
                var fields = doc.Fields;
                fields.Update();
                Converter.releaseCOMObject(fields);
                return;
            }

            try
            {
                // Update some of the field types in the document so the printed
                // PDF looks correct. Skips some field types (such as ASK) that would
                // create dialogs
                var docSections = doc.Sections;
                if (docSections.Count > 0)
                {
                    for (var dsi = 1; dsi <= docSections.Count; dsi++)
                    {
                        var section = docSections[dsi];
                        var sectionRange = section.Range;
                        var sectionFields = sectionRange.Fields;
                        var headers = section.Headers;
                        var footers = section.Footers;

                        if (sectionFields.Count > 0)
                        {
                            for (var si = 1; si <= sectionFields.Count; si++)
                            {
                                var sectionField = sectionFields[si];
                                WordConverter.updateField(sectionField, word, inputFile);
                                Converter.releaseCOMObject(sectionField);
                            }
                        }

                        updateHeaderFooterFields(headers, word, inputFile);
                        updateHeaderFooterFields(footers, word, inputFile);

                        Converter.releaseCOMObject(footers);
                        Converter.releaseCOMObject(headers);
                        Converter.releaseCOMObject(sectionFields);
                        Converter.releaseCOMObject(sectionRange);
                        Converter.releaseCOMObject(section);
                    }
                }
                Converter.releaseCOMObject(docSections);
            }
            catch (COMException)
            {
                // There can be odd errors when column widths are out of the page boundaries
                // See github issue #14
            }
            
            var docFields = doc.Fields;
            var storyRanges = doc.StoryRanges;

            if (docFields.Count > 0)
            {
                for (var fi = 1; fi <= docFields.Count; fi++)
                {
                    var docField = docFields[fi];
                    WordConverter.updateField(docField, word, inputFile);
                    Converter.releaseCOMObject(docField);
                }
            }

            foreach (Range range in storyRanges)
            {
                updateFieldsInRange(range, word, inputFile);
                Converter.releaseCOMObject(range);
            }
 
            Converter.releaseCOMObject(storyRanges);
            Converter.releaseCOMObject(docFields);
        }

        // update fields in a header or footer
        private static void updateHeaderFooterFields(HeadersFooters list, Microsoft.Office.Interop.Word.Application word, String filename)
        {
            foreach (Microsoft.Office.Interop.Word.HeaderFooter item in list)
            {
                if (item.Exists && !item.LinkToPrevious)
                {
                    var range = item.Range;
                    updateFieldsInRange(range, word, filename);
                    Converter.releaseCOMObject(range);
                }
                Converter.releaseCOMObject(item);
            }
        }

        // update all fields in a range
        private static void updateFieldsInRange(Range range, Microsoft.Office.Interop.Word.Application word, String filename)
        {
            var rangeFields = range.Fields;
            if (rangeFields.Count > 0)
            {
                for (var i = 1; i <= rangeFields.Count; i++)
                {
                    var field = rangeFields[i];
                    WordConverter.updateField(field, word, filename);
                    Converter.releaseCOMObject(field);
                }
            }
            Converter.releaseCOMObject(rangeFields);
        }

        // Update a specific field
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
                case WdFieldType.wdFieldIf:
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

        // Check if the document has any broken links from shapes and inline shapes.
        // We need to know this to determine if it is safe to perform
        // an update on all fields
        private static bool hasBrokenLinks(Document doc)
        {
            var hasBrokenLinks = false;
            var docShapes = doc.Shapes;
            hasBrokenLinks = hasBrokenLinksInShapeList<Shapes>(ref docShapes);
            if (!hasBrokenLinks)
            {
                // If there are no broken Shapes, then try the inline shapes list
                var inlineShapes = doc.InlineShapes;
                hasBrokenLinks = hasBrokenLinksInShapeList<InlineShapes>(ref inlineShapes);
                Converter.releaseCOMObject(inlineShapes);
            }
            Converter.releaseCOMObject(docShapes);
            return hasBrokenLinks;
        }

        // Loop through a list of shapes or inline shapes finding out if
        // any one has a broken reference
        private static bool hasBrokenLinksInShapeList<T>(ref T shapeList) 
            where T : IEnumerable
        {
            var hasBrokenLinks = false;
            var items = shapeList.GetEnumerator();
            while (items.MoveNext()) {
                var shapeThing = items.Current;
                var linkFormat = (typeof(T) == typeof(Shapes) ? ((Shape)shapeThing).LinkFormat : ((InlineShape)shapeThing).LinkFormat);
                if (linkFormat != null)
                {
                    // See if the linked file exists (if it is a local path and not a URL)
                    // Treat cid: references as also broken
                    var sourceName = linkFormat.SourceFullName;
                    var linkPath = sourceName.ToString();
                    if (linkPath.IndexOf("cid:") == 0 || 
                        (linkPath.ToLower().IndexOf("http://") != 0 &&
                         linkPath.ToLower().IndexOf("https://") != 0 && !File.Exists(linkPath)))
                    {
                        hasBrokenLinks = true;
                    }
                    Converter.releaseCOMObject(sourceName);
                }
                Converter.releaseCOMObject(linkFormat);
                Converter.releaseCOMObject(shapeThing);
                if (hasBrokenLinks)
                {
                    // If there are broken links, we can break out now since we
                    // don't care about anything else
                    break;
                }
            }
            Converter.releaseCOMObject(items);
            return hasBrokenLinks;
        }
    }
}
