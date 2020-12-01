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
    class WordConverter : Converter
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
            Application word = null;
            object oMissing = System.Reflection.Missing.Value;
            Template tmpl;
            String temporaryStorageDir = null;
            float wordVersion = 0;
            List<AppOption> wordOptionList = new List<AppOption>();
            try
            {
                String filename = (String)inputFile;
                Boolean hasSignatures = WordConverter.HasDigitalSignatures(filename);
                Boolean fileIsCorrupt = WordConverter.IsFileCorrupt(filename);
                Boolean visible = !(Boolean)options["hidden"];
                Boolean openAndRepair = !(Boolean)options["word_no_repair"];
                Boolean nowrite = (Boolean)options["readonly"] || fileIsCorrupt;
                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean includeTags = !(Boolean)options["excludetags"];
                Boolean bitmapMissingFonts = !(Boolean)options["word_ref_fonts"];
                Boolean autosave = options.ContainsKey("IsTempWord") && (Boolean)options["IsTempWord"];
                bool pdfa = (Boolean)options["pdfa"] ? true : false;
                String writePassword = "";
                String readPassword = "";
                int maxPages = 0;

                WdExportOptimizeFor quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                WdExportItem showMarkup = WdExportItem.wdExportDocumentContent;
                WdExportCreateBookmarks bookmarks = (Boolean)options["bookmarks"] ?
                    WdExportCreateBookmarks.wdExportCreateHeadingBookmarks :
                    WdExportCreateBookmarks.wdExportCreateNoBookmarks;
                Options wdOptions = null;
                Documents documents = null;
                Template normalTemplate = null;

                tmpl = null;
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
                        return (int)ExitCode.ApplicationError;
                    }
                }

                wdOptions = word.Options;
                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                // Issue #48 - we should allow control over whether the history is lost
                if (!(Boolean)options["word_keep_history"])
                {
                    word.DisplayRecentFiles = false;
                }
                word.DisplayDocumentInformationPanel = false;
                word.FeatureInstall = Microsoft.Office.Core.MsoFeatureInstall.msoFeatureInstallNone;
                wordVersion = (float)System.Convert.ToDecimal(word.Version, new CultureInfo("en-US"));

                // Set the Word options in a way that allows us to reset the options when we finish
                try
                {
                    wordOptionList.Add(new AppOption("AlertIfNotDefault", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("AllowReadingMode", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("PrecisePositioning", true, ref wdOptions));
                    wordOptionList.Add(new AppOption("UpdateFieldsAtPrint", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("UpdateLinksAtPrint", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("UpdateLinksAtOpen", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("UpdateFieldsWithTrackedChangesAtPrint", false, ref wdOptions));
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
                    wordOptionList.Add(new AppOption("AllowFastSave", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("BackgroundOpen", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("ShowMarkupOpenSave", false, ref wdOptions));
                    wordOptionList.Add(new AppOption("SaveInterval", 0, ref wdOptions));
                    wordOptionList.Add(new AppOption("PrintHiddenText", (Boolean)options["word_show_hidden"], ref wdOptions));
                }
                catch (SystemException)
                {
                }

                // Set up the PDF output quality
                if ((Boolean)options["print"])
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForPrint;
                }
                if ((Boolean)options["screen"])
                {
                    quality = WdExportOptimizeFor.wdExportOptimizeForOnScreen;
                }

                if ((Boolean)options["markup"])
                {
                    showMarkup = WdExportItem.wdExportDocumentWithMarkup;
                }

                if (!String.IsNullOrEmpty((String)options["password"]))
                {
                    readPassword = (String)options["password"];
                }

                if (!String.IsNullOrEmpty((String)options["writepassword"]))
                {
                    writePassword = (String)options["writepassword"];
                }

                // Large Word files may simply not print reliably - if the word_max_pages
                // configuration option is set, then we must close up and forget about 
                // converting the file.
                maxPages = (int)options[@"word_max_pages"];

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
                    autosave = false;
                    openAndRepair = false;
                }

                Document doc = null;
                try
                {
                    if ((bool)options["merge"] && !String.IsNullOrEmpty((string)options["template"]) &&
                        File.Exists((string)options["template"]) &&
                        System.Text.RegularExpressions.Regex.IsMatch((string)options["template"], @"^.*\.dot[mx]?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        // Create a new document based on a template
                        doc = documents.Add((string)options["template"]);
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
                    Console.WriteLine("Unable to open file");
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
                        docWinView.Type = WdViewType.wdPrintPreview;
                    }
                    catch (Exception) { }

                    // Handle markup
                    try
                    {
                        if ((Boolean)options["word_show_all_markup"])
                        {
                            options["word_show_comments"] = true;
                            options["word_show_revs_comments"] = true;
                            options["word_show_format_changes"] = true;
                            options["word_show_ink_annot"] = true;
                            options["word_show_ins_del"] = true;
                        }
                        if ((Boolean)options["word_show_comments"] ||
                            (Boolean)options["word_show_revs_comments"] ||
                            (Boolean)options["word_show_format_changes"] ||
                            (Boolean)options["word_show_ink_annot"] ||
                            (Boolean)options["word_show_ins_del"] ||
                            showMarkup == WdExportItem.wdExportDocumentWithMarkup)
                        {
                            docWinView.MarkupMode = (Boolean)options["word_markup_balloon"] ?
                                WdRevisionsMode.wdBalloonRevisions : WdRevisionsMode.wdInLineRevisions;
                        }
                        word.PrintPreview = false;
                        docWinView.RevisionsView = WdRevisionsView.wdRevisionsViewFinal;
                        docWinView.ShowRevisionsAndComments = (Boolean)options["word_show_revs_comments"];
                        docWinView.ShowComments = (Boolean)options["word_show_comments"];
                        docWinView.ShowFormatChanges = (Boolean)options["word_show_format_changes"];
                        docWinView.ShowInkAnnotations = (Boolean)options["word_show_ink_annot"];
                        docWinView.ShowInsertionsAndDeletions = (Boolean)options["word_show_ins_del"];
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
                        if (autosave) { doc.Save(); doc.Saved = true; }
                        doc.TrackMoves = false;
                        doc.TrackRevisions = false;
                        doc.TrackFormatting = false;

                        if ((Boolean)options["word_fix_table_columns"])
                        {
                            FixWordTableColumnWidths(doc);
                        }
                    }

                    normalTemplate.Saved = true;

                    // Hide the document window if need be
                    if ((Boolean)options["hidden"])
                    {
                        word.Visible = false;
                        var activeWin = word.ActiveWindow;
                        activeWin.Visible = false;
                        activeWin.WindowState = WdWindowState.wdWindowStateMinimize;
                        ReleaseCOMObject(activeWin);
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
                        UpdateDocumentFields(doc, word, inputFile, options);
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
                    if (autosave)
                    {
                        doc.Save();
                    }
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

                if (String.IsNullOrEmpty((string)options["printer"])) {
                    // No printer given, so export
                    try
                    {
                        doc.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF, false,
                        quality, WdExportRange.wdExportAllDocument,
                        1, 1, showMarkup, includeProps, true, bookmarks, includeTags, bitmapMissingFonts, pdfa);
                    } catch (Exception)
                    {
                        // Couldn't export, so see if there is a fallback printer
                        if (!String.IsNullOrEmpty((string)options["fallback_printer"])) {
                            PrintToGhostscript((string)options["fallback_printer"], outputFile, printFunc);
                        }
                        else
                        {
                            throw;
                        }
                    }
                } else
                {
                    PrintToGhostscript((string)options["printer"], outputFile, printFunc);
                }

                if (tmpl != null)
                {
                    tmpl.Saved = true;
                }

                object saveChanges = autosave ? WdSaveOptions.wdSaveChanges : WdSaveOptions.wdDoNotSaveChanges;
                if (nowrite)
                {
                    doc.Saved = true;
                }
                normalTemplate.Saved = true;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);

                // Reset options
                foreach (AppOption opt in wordOptionList)
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
        private static bool CloseWordApplication(Microsoft.Office.Interop.Word.Application word)
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
            public string Name { get; set; }
            public Boolean Value { get; set; }
            public Boolean OriginalValue { get; set; }
            public int IntValue { get; set; }
            public int OriginalIntValue { get; set; }
            protected Type VarType { get; set; }
            public AppOption(string name, Boolean value, ref Options wdOptions)
            {
                try
                {
                    Name = name;
                    Value = value;
                    VarType = typeof(Boolean);
                    OriginalValue = (Boolean)wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, wdOptions, null);

                    if (OriginalValue != value)
                    {
                        wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { value });
                    }
                }
                catch
                {
                    // We may be setting word options that are not available in the version of word
                    // being used, so just skip these errors
                }
            }
            public AppOption(string name, int value, ref Options wdOptions)
            {
                try
                {
                    Name = name;
                    IntValue = value;
                    VarType = typeof(int);
                    OriginalIntValue = (int)wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, wdOptions, null);

                    if (OriginalIntValue != value)
                    {
                        wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { value });
                    }
                }
                catch
                {
                    // We may be setting word options that are not available in the version of word
                    // being used, so just skip these errors
                }
            }

            // Allow the value on the options to be reset
            public void ResetValue(ref Options wdOptions)
            {
                if (VarType == typeof(Boolean))
                {
                    if (Value != this.OriginalValue)
                    {
                        wdOptions.GetType().InvokeMember(Name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { OriginalValue });
                    }
                }
                else
                {
                    if (IntValue != OriginalIntValue)
                    {
                        wdOptions.GetType().InvokeMember(Name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { OriginalIntValue });
                    }
                }
            }
        }

        // Update all the fields in a document
        private static void UpdateDocumentFields(Microsoft.Office.Interop.Word.Document doc, Microsoft.Office.Interop.Word.Application word, String inputFile, Hashtable options)
        {
            // Update fields quickly if it is safe to do so. We have
            // to check for broken links as they may raise Word dialogs or leave broken content
            if ((Boolean)options["word_field_quick_update"] ||
                ((Boolean)options["word_field_quick_update_safe"] && !HasBrokenLinks(doc)))
            {
                var fields = doc.Fields;
                foreach (Field f in fields)
                {
                    if (f.Type == WdFieldType.wdFieldFillIn)
                    {
                        f.Unlink();
                    }
                }
                fields.Update();
                ReleaseCOMObject(fields);
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
                if (!LooksLinkOpenXmlWord(filename))
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
            if (!LooksLinkOpenXmlWord(filename))
            {
                return false;
            }
            WordprocessingDocument document = null;

            try
            {
                document = WordprocessingDocument.Open(filename, false);
            }
            catch (System.IO.FileFormatException)
            {
                return true;
            }

            document.Close();
            return false;
        }

        protected static bool LooksLinkOpenXmlWord(string filename)
        {
            // Only work for things that look like OpenXml format
            return (System.Text.RegularExpressions.Regex.IsMatch(filename, @"^.*\.doc[mx]?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase));
        }
    }
}
