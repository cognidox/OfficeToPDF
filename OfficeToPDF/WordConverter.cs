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
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Word files
    /// </summary>
    internal class WordConverter : Converter, IConverter
    {
        int IConverter.Convert(String inputFile, String outputFile, ArgParser options, ref List<PDFBookmark> bookmarks)
        {
            if (options.verbose)
            {
                Console.WriteLine("Converting with Word converter");
            }
            return Convert(inputFile, outputFile, options);
        }

        public static ExitCode StartWord(ref Boolean running, ref Application word)
        {
            try
            {
                word = (Microsoft.Office.Interop.Word.Application)Marshal.GetActiveObject("Word.Application");
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
                    ReleaseCOMObject(word);
                    return ExitCode.ApplicationError;
                }
            }
            return ExitCode.Success;
        }

        /// <summary>
        /// Convert a Word file to PDF
        /// </summary>
        /// <param name="inputFile">Full path of the input Word file</param>
        /// <param name="outputFile">Full path of the output PDF</param>
        /// <returns></returns>
        internal static int Convert(String inputFile, String outputFile, ArgParser options)
        {
            Boolean running = options.noquit;
            Application word = null;
            object oMissing = System.Reflection.Missing.Value;
            Template tmpl;
            String temporaryStorageDir = null;
            float wordVersion = 0;
            List<IAppOption> wordOptionList = new List<IAppOption>();
            IWatchdog watchdog = new NullWatchdog();
            try
            {
                ExitCode result = StartWord(ref running, ref word);
                if (result != ExitCode.Success)
                    return (int)result;

                watchdog = WatchdogFactory.CreateStarted(word, options.timeout);

                String filename = (String)inputFile;
                Boolean hasSignatures = WordConverter.HasDigitalSignatures(filename);
                Boolean fileIsCorrupt = WordConverter.IsFileCorrupt(filename);
                Boolean visible = !options.hidden;
                Boolean openAndRepair = !options.word_no_repair;
                Boolean nowrite = options.@readonly || fileIsCorrupt;
                Boolean includeProps = !options.excludeprops;
                Boolean includeTags = !options.excludetags;
                Boolean bitmapMissingFonts = !options.word_ref_fonts;
                Boolean isTempWord = options.IsTempWord; // Indication of converting an Outlook message
                
                String writePassword = "";
                String readPassword = "";
                int maxPages = 0;

                WdExportOptimizeFor quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                WdExportItem showMarkup = WdExportItem.wdExportDocumentContent;
                WdExportCreateBookmarks bookmarks = options.bookmarks ?
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks :
                    WdExportCreateBookmarks.wdExportCreateNoBookmarks;
                Options wdOptions = null;
                Documents documents = null;
                Template normalTemplate = null;
                
                tmpl = null;
                wdOptions = word.Options;
                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                // Issue #48 - we should allow control over whether the history is lost
                if (!options.word_keep_history)
                {
                    word.DisplayRecentFiles = false;
                }
                word.DisplayDocumentInformationPanel = false;
                word.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                wordVersion = (float)System.Convert.ToDecimal(word.Version, new CultureInfo("en-US"));
                
                // Set the Word options in a way that allows us to reset the options when we finish
                try
                {
                    wordOptionList.Add(AppOptionFactory.Create(nameof(wdOptions.AlertIfNotDefault), false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("AllowReadingMode", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("PrecisePositioning", true, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("UpdateFieldsAtPrint", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("UpdateLinksAtPrint", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("UpdateLinksAtOpen", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("UpdateFieldsWithTrackedChangesAtPrint", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("WarnBeforeSavingPrintingSendingMarkup", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("BackgroundSave", true, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("SavePropertiesPrompt", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("DoNotPromptForConvert", true, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("PromptUpdateStyle", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("ConfirmConversions", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("CheckGrammarAsYouType", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("CheckGrammarWithSpelling", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("CheckSpellingAsYouType", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("DisplaySmartTagButtons", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("EnableLivePreview", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("ShowReadabilityStatistics", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("SuggestSpellingCorrections", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("AllowDragAndDrop", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("EnableMisusedWordsDictionary", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("ShowFormatError", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("StoreRSIDOnSave", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("SaveNormalPrompt", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("AllowFastSave", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("BackgroundOpen", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("ShowMarkupOpenSave", false, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("SaveInterval", 0, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("PrintHiddenText", options.word_show_hidden, ref wdOptions));
                    wordOptionList.Add(AppOptionFactory.Create("MapPaperSize", !options.word_no_map_papersize, ref wdOptions));
                }
                catch (SystemException)
                {
                }

                // Set up the PDF output quality
                if (options.print)
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                }
                if (options.screen)
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForOnScreen;
                }

                if (options.markup)
                {
                    showMarkup = WdExportItem.wdExportDocumentWithMarkup;
                }

                if (!String.IsNullOrEmpty(options.password))
                {
                    readPassword = options.password;
                }

                if (!String.IsNullOrEmpty(options.writepassword))
                {
                    writePassword = options.writepassword;
                }

                // Large Word files may simply not print reliably - if the word_max_pages
                // configuration option is set, then we must close up and forget about 
                // converting the file.
                maxPages = options.word_max_pages;

                documents = word.Documents;
                normalTemplate = word.NormalTemplate;

                // Check for password protection and no password
                if (IsPasswordProtected(inputFile) && String.IsNullOrEmpty(readPassword))
                {
                    normalTemplate.Saved = true;
                    Console.WriteLine("Unable to open password protected file");
                    return (int)ExitCode.PasswordFailure;
                }

                // If we are opening a document with a write password and no read password, and
                // we are not in read only mode, we should follow the document properties and
                // enforce a read only open. If we do not, Word pops up a dialog
                if (!nowrite && String.IsNullOrEmpty(writePassword) && IsReadOnlyEnforced(inputFile))
                {
                    nowrite = true;
                }

                // Having signatures means we should open the document very carefully
                if (hasSignatures)
                {
                    nowrite = true;
                    openAndRepair = false;
                }
                
                Document doc = null;
                try
                {
                    if (options.merge && !String.IsNullOrEmpty(options.template) &&
                        File.Exists(options.template) &&
                        System.Text.RegularExpressions.Regex.IsMatch(options.template, @"^.*\.dot[mx]?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        // Create a new document based on a template
                        doc = documents.Add(options.template);
                        Object rStart = 0;
                        Object rEnd = 0;
                        Range range = doc.Range(rStart, rEnd);
                        range.InsertFile(inputFile);
                        ReleaseCOMObject(range);
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
                        doc = documents.OpenNoRepairDialog(FileName: filename, ReadOnly: nowrite, PasswordDocument: readPassword, WritePasswordDocument: writePassword, Visible: visible, OpenAndRepair: openAndRepair);
                    }
                }
                catch (COMException)
                {
                    Console.WriteLine("Unable to open file " + filename);
                    return (int)ExitCode.FileOpenFailure;
                }

                // Check if there are signatures in the document which changes how we do things
                if (hasSignatures)
                {
                    // Add in a delay to allow signatures to load
                    Thread.Sleep(500);
                }
                else
                {
                    Window docWin = null;
                    View docWinView = null;
                    doc.Activate();

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
                    docWin = doc.ActiveWindow;
                    docWinView = docWin.View;
                    if (wordVersion >= 15)
                    {
                        docWinView.ReadingLayout = false;
                    }

                    // Sometimes the print view will not be available (e.g. for a blog post)
                    // Try and switch view
                    try
                    {
                        // There is an issue here that if a document has fill-in fields, they'll
                        // prompt for input as the view type changes.
                        RemoveFillInFields(doc, false);
                        docWinView.Type = WdViewType.wdPrintPreview;
                    }
                    catch (Exception) { }

                    // Handle markup
                    try
                    {
                        if (options.word_show_all_markup)
                        {
                            options.word_show_comments = true;
                            options.word_show_revs_comments = true;
                            options.word_show_format_changes = true;
                            options.word_show_ink_annot = true;
                            options.word_show_ins_del = true;
                        }
                        if (options.word_show_comments ||
                            options.word_show_revs_comments ||
                            options.word_show_format_changes ||
                            options.word_show_ink_annot ||
                            options.word_show_ins_del ||
                            showMarkup == WdExportItem.wdExportDocumentWithMarkup)
                        {
                            docWinView.MarkupMode = options.word_markup_balloon ?
                                WdRevisionsMode.wdBalloonRevisions : WdRevisionsMode.wdInLineRevisions;
                        }
                        word.PrintPreview = false;
                        docWinView.RevisionsView = WdRevisionsView.wdRevisionsViewFinal;
                        docWinView.ShowRevisionsAndComments = options.word_show_revs_comments;
                        docWinView.ShowComments = options.word_show_comments;
                        docWinView.ShowFormatChanges = options.word_show_format_changes;
                        docWinView.ShowInkAnnotations = options.word_show_ink_annot;
                        docWinView.ShowInsertionsAndDeletions = options.word_show_ins_del;
                    }
                    catch (SystemException e) {
                        Console.WriteLine("Failed to set revision settings {0}", e.Message);
                    }

                    // Try to avoid Word thinking any changes are happening to the document
                    doc.SpellingChecked = true;
                    doc.GrammarChecked = true;

                    // Changing these properties may be disallowed if the document is protected
                    // and is not signed
                    if (doc.ProtectionType == WdProtectionType.wdNoProtection && !hasSignatures)
                    {
                        doc.TrackMoves = false;
                        doc.TrackRevisions = false;
                        doc.TrackFormatting = false;

                        if (options.word_fix_table_columns)
                        {
                            FixWordTableColumnWidths(doc);
                        }
                    }

                    normalTemplate.Saved = true;

                    // Hide the document window if need be
                    if (options.hidden)
                    {
                        word.Visible = false;
                        var activeWin = word.ActiveWindow;
                        activeWin.Visible = false;
                        activeWin.WindowState = WdWindowState.wdWindowStateMinimize;
                        ReleaseCOMObject(activeWin);
                    }

                    // Check if we have a template file to apply to this document
                    // The template must be a file and must end in .dot, .dotx or .dotm
                    if (!String.IsNullOrEmpty(options.template) && !options.merge)
                    {
                        string template = options.template;
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
                    if (! options.word_no_field_update)
                    {
                        UpdateDocumentFields(doc, word, inputFile, options);
                    }

                    var pageSetup = doc.PageSetup;
                    if (options.word_header_dist >= 0)
                    {
                        pageSetup.HeaderDistance = options.word_header_dist;
                    }
                    if (options.word_footer_dist >= 0)
                    {
                        pageSetup.FooterDistance = options.word_footer_dist;
                    }
                    ReleaseCOMObject(pageSetup);
                    try
                    {
                        // Make sure we are not in a header footer view
                        docWinView.SeekView = WdSeekView.wdSeekPrimaryHeader;
                        docWinView.SeekView = WdSeekView.wdSeekPrimaryFooter;
                        docWinView.SeekView = WdSeekView.wdSeekMainDocument;
                    }
                    catch (Exception)
                    {
                        // We might fail when switching views
                    }

                    normalTemplate.Saved = true;
                    doc.Saved = true;
                    ReleaseCOMObject(docWinView);
                    ReleaseCOMObject(docWin);
                }

                // Set up a delegate function if we're using a printer
                PrintDocument printFunc = delegate (string destination, string printerName)
                {
                    word.ActivePrinter = printerName;
                    doc.PrintOut(Background: false, OutputFileName: destination);
                };

                // Enable screen updating before exporting to ensure that Word
                // renders borders correctly
                word.ScreenUpdating = true;

                if (String.IsNullOrEmpty(options.printer)) {
                    // No printer given, so export
                    try
                    {
                        doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, false,
                        quality, WdExportRange.wdExportAllDocument,
                        1, 1, showMarkup, includeProps, true, bookmarks, includeTags, bitmapMissingFonts, options.pdfa);
                    } catch (Exception)
                    {
                        // Couldn't export, so see if there is a fallback printer
                        if (!String.IsNullOrEmpty(options.fallback_printer)) {
                            PrintToGhostscript(options.fallback_printer, outputFile, printFunc);
                        }
                        else
                        {
                            throw;
                        }
                    }
                } else
                {
                    PrintToGhostscript(options.printer, outputFile, printFunc);
                }

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

                // Reset options
                foreach (IAppOption opt in wordOptionList)
                {
                    opt.ResetValue(ref wdOptions);
                }

                ReleaseCOMObject(wdOptions);
                ReleaseCOMObject(documents);
                ReleaseCOMObject(doc);
                ReleaseCOMObject(tmpl);
                ReleaseCOMObject(normalTemplate);

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
                    CloseWordApplication(word);
                }
                ReleaseCOMObject(word);
            }
        }

        // Update tables
        // There may be an issue with some tables which are auto-generated where columns widths
        // in the body of the tables don't match the header columns
        // See github issue 27
        private static void FixWordTableColumnWidths(Document doc)
        {
            try
            {
                Tables tables = doc.Tables;
                for (var i = 1; i <= tables.Count; i++)
                {
                    Table t = tables[i];
                    Rows allRows = t.Rows;
                    if (allRows.Count > 1 && i > 1)
                    {
                        Row firstRow = allRows.First;
                        if (firstRow.HeadingFormat == 0)
                        {
                            Cells cells = firstRow.Cells;
                            if (cells.Count > 0)
                            {
                                Table previousTable = tables[i - 1];
                                Rows previousRows = previousTable.Rows;
                                if (previousRows.Count > 0)
                                {
                                    Row previousFirstRow = previousRows.First;
                                    if (previousFirstRow.HeadingFormat == -1)
                                    {
                                        Columns columns = t.Columns;
                                        Columns previousColumns = previousTable.Columns;
                                        // Only match columns widths if the tables match column count and width
                                        if (t.PreferredWidth == previousTable.PreferredWidth && columns.Count == previousColumns.Count)
                                        {
                                            for (var cx = 1; cx <= cells.Count; cx++)
                                            {
                                                Cell cell = cells[cx];
                                                if (cell.PreferredWidth == 0 && cell.PreferredWidthType == WdPreferredWidthType.wdPreferredWidthAuto)
                                                {
                                                    Column thisColumn = cell.Column;
                                                    Column previousColumn = previousTable.Columns[cell.ColumnIndex];
                                                    thisColumn.Width = previousColumn.Width;
                                                    ReleaseCOMObject(previousColumn);
                                                    ReleaseCOMObject(thisColumn);
                                                }
                                                ReleaseCOMObject(cell);
                                            }
                                        }
                                        ReleaseCOMObject(columns);
                                        ReleaseCOMObject(previousColumns);
                                        ReleaseCOMObject(cells);
                                    }
                                    ReleaseCOMObject(previousFirstRow);
                                }
                                ReleaseCOMObject(previousRows);
                                ReleaseCOMObject(previousTable);
                            }
                        }
                        ReleaseCOMObject(firstRow);
                    }
                    ReleaseCOMObject(allRows);
                    ReleaseCOMObject(t);
                }
                ReleaseCOMObject(tables);
            }
            catch (Exception) { }
        }

        // Try and close Word, giving time for Office to get
        // itself in order
        internal static bool CloseWordApplication(Microsoft.Office.Interop.Word.Application word)
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

        // Remove fill-in fields which can cause blocking dialogs
        private static void RemoveFillInFields(Microsoft.Office.Interop.Word.Document doc, bool updateField)
        {
            bool altered = false;
            var ranges = doc.StoryRanges;
            try
            {
                foreach (Range range in ranges) {
                    var fields = range.Fields;
                    try {
                        foreach (Field f in fields)
                        {
                            try
                            {
                                bool isBrokenLink = false;
                                // Break included files that may cause a dialog
                                if (f.Type == WdFieldType.wdFieldIncludePicture || f.Type == WdFieldType.wdFieldInclude)
                                {
                                    if (!File.Exists(f.LinkFormat.SourceFullName))
                                    {
                                        isBrokenLink = true;
                                    }
                                }
                                if (f.Type == WdFieldType.wdFieldFillIn || isBrokenLink)
                                {
                                    try
                                    {
                                        altered = true;
                                        f.Unlink();
                                    }
                                    catch (Exception)
                                    {
                                        f.Delete();
                                    }
                                }
                            }
                            finally
                            {
                                ReleaseCOMObject(f);
                            }
                        }
                        if (updateField)
                        {
                            fields.Update();
                        }
                    }
                    finally
                    {
                        ReleaseCOMObject(fields);
                    }
                    ReleaseCOMObject(range);
                }
            }
            finally
            {
                ReleaseCOMObject(ranges);
                if (altered)
                {
                    doc.Saved = true;
                }
            }
        }

        // Update all the fields in a document
        private static void UpdateDocumentFields(Microsoft.Office.Interop.Word.Document doc, Microsoft.Office.Interop.Word.Application word, String inputFile, ArgParser options)
        {
            // Update fields quickly if it is safe to do so. We have
            // to check for broken links as they may raise Word dialogs or leave broken content
            if (options.word_field_quick_update ||
                (options.word_field_quick_update_safe && !HasBrokenLinks(doc)))
            {
                RemoveFillInFields(doc, true);
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
                                UpdateField(sectionField, word, inputFile);
                                ReleaseCOMObject(sectionField);
                            }
                        }

                        UpdateHeaderFooterFields(headers, word, inputFile);
                        UpdateHeaderFooterFields(footers, word, inputFile);

                        ReleaseCOMObject(footers);
                        ReleaseCOMObject(headers);
                        ReleaseCOMObject(sectionFields);
                        ReleaseCOMObject(sectionRange);
                        ReleaseCOMObject(section);
                    }
                }
                ReleaseCOMObject(docSections);
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
                    UpdateField(docField, word, inputFile);
                    ReleaseCOMObject(docField);
                }
            }

            foreach (Range range in storyRanges)
            {
                UpdateFieldsInRange(range, word, inputFile);
                ReleaseCOMObject(range);
            }

            ReleaseCOMObject(storyRanges);
            ReleaseCOMObject(docFields);
        }

        // update fields in a header or footer
        private static void UpdateHeaderFooterFields(HeadersFooters list, Microsoft.Office.Interop.Word.Application word, String filename)
        {
            foreach (HeaderFooter item in list)
            {
                if (item.Exists && !item.LinkToPrevious)
                {
                    var range = item.Range;
                    UpdateFieldsInRange(range, word, filename);
                    ReleaseCOMObject(range);
                }
                ReleaseCOMObject(item);
            }
        }

        // update all fields in a range
        private static void UpdateFieldsInRange(Range range, Microsoft.Office.Interop.Word.Application word, String filename)
        {
            var rangeFields = range.Fields;
            if (rangeFields.Count > 0)
            {
                for (var i = 1; i <= rangeFields.Count; i++)
                {
                    var field = rangeFields[i];
                    UpdateField(field, word, filename);
                    ReleaseCOMObject(field);
                }
            }
            ReleaseCOMObject(rangeFields);
        }

        // Update a specific field
        private static void UpdateField(Field field, Microsoft.Office.Interop.Word.Application word, String filename)
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
                    ReleaseCOMObject(selection);
                    break;
            }
        }

        // Check if the document has any broken links from shapes and inline shapes.
        // We need to know this to determine if it is safe to perform
        // an update on all fields
        private static bool HasBrokenLinks(Microsoft.Office.Interop.Word.Document doc)
        {
            var hasBrokenLinks = false;
            var docShapes = doc.Shapes;
            hasBrokenLinks = HasBrokenLinksInShapeList<Shapes>(ref docShapes);
            if (!hasBrokenLinks)
            {
                // If there are no broken Shapes, then try the inline shapes list
                var inlineShapes = doc.InlineShapes;
                hasBrokenLinks = HasBrokenLinksInShapeList<InlineShapes>(ref inlineShapes);
                ReleaseCOMObject(inlineShapes);
            }
            ReleaseCOMObject(docShapes);
            return hasBrokenLinks;
        }

        // Loop through a list of shapes or inline shapes finding out if
        // any one has a broken reference
        private static bool HasBrokenLinksInShapeList<T>(ref T shapeList)
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
                    ReleaseCOMObject(sourceName);
                }
                ReleaseCOMObject(linkFormat);
                ReleaseCOMObject(shapeThing);
                if (hasBrokenLinks)
                {
                    // If there are broken links, we can break out now since we
                    // don't care about anything else
                    break;
                }
            }
            ReleaseCOMObject(items);
            return hasBrokenLinks;
        }

        // Use the OpenXML library to look for signatures
        protected static bool HasDigitalSignatures(string filename)
        {
            try
            {
                // Only work for things that look like OpenXml format
                if (!LooksLikeOpenXmlWord(filename))
                {
                    return false;
                }
                bool isSigned = false;
                var document = WordprocessingDocument.Open(filename, false);
                if (document != null)
                {
                    PackageDigitalSignatureManager dsm = new PackageDigitalSignatureManager(document.Package);
                    isSigned = dsm.IsSigned;
                    document.Close();
                    return isSigned;
                }
            }
            catch (Exception) { }

            return false;
        }

        protected static bool IsFileCorrupt(string filename)
        {
            // Only work for things that look like OpenXml format
            if (!LooksLikeOpenXmlWord(filename))
            {
                return false;
            }
            WordprocessingDocument document = null;

            try
            {
                document = WordprocessingDocument.Open(filename, false);
                document.Close();
            }
            catch (System.IO.FileFormatException)
            {
                return true;
            }
            catch(Exception) { }
            
            return false;
        }

        protected static bool LooksLikeOpenXmlWord(string filename)
        {
            // Only work for things that look like OpenXml format
            return (System.Text.RegularExpressions.Regex.IsMatch(filename, @"^.*\.doc[mx]?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase));
        }
    }
}
