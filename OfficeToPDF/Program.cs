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
using System.Drawing.Printing;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;

namespace OfficeToPDF
{
    // Set up the exit codes
    [Flags]
    public enum ExitCode : int
    {
        Success = 0,
        Failed = 1,
        UnknownError = 2,
        PasswordFailure = 4,
        InvalidArguments = 8,
        FileOpenFailure = 16,
        UnsupportedFileFormat = 32,
        FileNotFound = 64,
        DirectoryNotFound = 128,
        WorksheetNotFound = 256,
        EmptyWorksheet = 512,
        PDFProtectedDocument = 1024,
        ApplicationError = 2048,
        NoPrinters = 4096
    }

    public enum MergeMode : int
    {
        None = 0,
        Prepend = 1,
        Append = 2
    }

    public enum MetaClean : int
    {
        None = 0,
        Basic = 1,
        Full = 2
    }

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ArgParser options = new ArgParser();
            List<PDFBookmark> documentBookmarks = new List<PDFBookmark>();

            // We need some printers to keep office happy
            Dictionary<string,bool> installedPrinters = GetInstalledPrinters();
            if (installedPrinters.Count <= 0)
            {
                Console.WriteLine("There are no installed printers, so conversion can not proceed");
                Exit(ExitCode.Failed | ExitCode.NoPrinters);
            }

            ExitCode result = options.Parse(args, installedPrinters);
            if (result != ExitCode.Success)
                Exit(result);


            // Need to error here, as we need input and output files as the
            // arguments to this script
            if (options.filesSeen != 1 && options.filesSeen != 2)
            {
                options.ShowHelpAndExit();
            }

            // Make sure we only choose one of /screen or /print options
            if ((Boolean)options["screen"] && (Boolean)options["print"])
            {
                Console.WriteLine("You can only use one of /screen or /print - not both");
                Exit(ExitCode.Failed | ExitCode.InvalidArguments);
            }

            // Make sure the input file looks like something we can handle (i.e. has an office
            // filename extension)
            Regex fileMatch = new Regex(@"\.(((ppt|pps|pot|do[ct]|xls|xlt)[xm]?)|xps|xlsb|od[spt]|rtf|csv|vsd[xm]?|vd[xw]|em[fz]|dwg|dxf|wmf|pub|msg|vcf|ics|mpp|svg|txt|html?|wpd)$", RegexOptions.IgnoreCase);
            if (fileMatch.Matches(options.files[0]).Count != 1)
            {
                Console.WriteLine("Input file can not be handled. Must be Word, PowerPoint, Excel, Outlook, Publisher, XPS or Visio");
                Exit(ExitCode.Failed | ExitCode.UnsupportedFileFormat);
            }

            if (options.filesSeen == 1)
            {
                // If only one file is seen, we just swap the extension
                options.files[1] = Path.ChangeExtension(options.files[0], "pdf");
            }
            else
            {
                // If the second file is a directory, then we want to create the PDF
                // with the same name as the original (changing the extension to pdf),
                // but in the directory given by the path
                if (Directory.Exists(options.files[1]))
                {
                    options.files[1] = Path.Combine(options.files[1], Path.GetFileNameWithoutExtension(options.files[0]) + ".pdf");
                }
            }

            String inputFile = "";
            String outputFile = "";
            String finalOutputFile = "";

            // Make sure the input file exists and is readable
            FileInfo info;
            try
            {
                info = new FileInfo(options.files[0]);

                if (info == null || !info.Exists)
                {
                    Console.WriteLine("Input file doesn't exist");
                    Exit(ExitCode.Failed | ExitCode.FileNotFound);
                }
                inputFile = info.FullName;
                options["original_filename"] = info.Name;
                options["original_basename"] = info.Name.Substring(0, info.Name.Length - info.Extension.Length);
            }
            catch
            {
                Console.WriteLine("Unable to open input file");
                Exit(ExitCode.Failed | ExitCode.FileOpenFailure);
            }

            // Stop people using the template as the input file
            if (!String.IsNullOrEmpty((string)options["template"]) &&
                inputFile.Equals((string)options["template"]))
            {
                Console.WriteLine("Input file must be different from the template file");
                Exit(ExitCode.Failed | ExitCode.InvalidArguments);
            }

            // Make sure the destination location exists
            FileInfo outputInfo = new FileInfo(options.files[1]);
            // Remove the destination unless we're doing a PDF merge
            if (outputInfo != null)
            {
                outputFile = finalOutputFile = outputInfo.FullName;
                if (outputInfo.Exists)
                {
                    if ((MergeMode)options["pdf_merge"] == MergeMode.None)
                    {
                        // We are not merging, so delete the final destination
                        System.IO.File.Delete(outputInfo.FullName);
                    }
                    else
                    {
                        // We are merging, so make a temporary file
                        outputFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".pdf";
                    }
                }
                else
                {
                    // If there is no current output, no need to merge
                    options["pdf_merge"] = MergeMode.None;
                }
            }
            else
            {
                Console.WriteLine("Unable to determine output location");
                Exit(ExitCode.Failed | ExitCode.DirectoryNotFound);
            }

            if (!System.IO.Directory.Exists(outputInfo.DirectoryName))
            {
                Console.WriteLine("Output directory does not exist");
                Exit(ExitCode.Failed | ExitCode.DirectoryNotFound);
            }

            // We want the input and output files copied to a working area where
            // we can manipulate them
            if (options.has_working_dir)
            {
                // Create a local temporary area and put the input and output in separate
                // areas
                string workingInput = Path.Combine(options.working_dir, "input");
                string workingOutput = Path.Combine(options.working_dir, "output");
                System.IO.Directory.CreateDirectory(workingInput);
                System.IO.Directory.CreateDirectory(workingOutput);
                string workingSource = Path.Combine(workingInput, Path.GetFileName(inputFile));
                string workingDest = Path.Combine(workingOutput, Path.GetFileName(outputFile));
                options.working_orig_dest = outputFile;
                File.Copy(inputFile, workingSource);
                inputFile = workingSource;
                outputFile = workingDest;
                if (options.verbose)
                {
                    Console.WriteLine("Created working directory: {0}", options.working_dir);
                }
            }

            // Now, do the cleverness of determining what the extension is, and so, which
            // conversion class to pass it to
            ConverterFactory factory = new ConverterFactory();
            ExitCode converted = ExitCode.UnknownError;
            Match extMatch = fileMatch.Match(inputFile);
            if (extMatch.Success)
            {
                // Set and environment variable so Office application VBA
                // code can check for un-attended conversion and avoid showing
                // blocking dialogs
                Environment.SetEnvironmentVariable("OFFICE2PDF_AUTO_CONVERT", "1");

                if (options.verbose)
                {
                    Console.WriteLine("Converting {0} to {1}", inputFile, finalOutputFile);
                }

                string extension = extMatch.Groups[1].ToString().ToLower();

                IConverter converter = factory.Create(extension);

                converted = converter.Convert(inputFile, outputFile, options, ref documentBookmarks);
            }

            // Clear up the working directory and restore the expected output
            if (options.has_working_dir)
            {
                if (File.Exists(outputFile))
                {
                    if (options.verbose)
                    {
                        Console.WriteLine("Copying working file {0} to {1}", outputFile, options.working_orig_dest);
                    }
                    File.Copy(outputFile, options.working_orig_dest);
                    outputFile = options.working_orig_dest;   
                }
                Directory.Delete(options.working_dir, true);
            }

            if (converted != ExitCode.Success)
            {
                Console.WriteLine("Did not convert");
                // Return the general failure code and the specific failure code
                Exit(ExitCode.Failed | converted);
            }
            else
            {
                if (options.verbose)
                {
                    Console.WriteLine("Completed Conversion");
                }

                if (documentBookmarks.Count > 0)
                {
                    AddPDFBookmarks(outputFile, documentBookmarks, options, null);
                }
                
                // Determine if we have to post-process the PDF
                if (options.postProcessPDF)
                {
                    PostProcessPDFFile(outputFile, finalOutputFile, options, options.postProcessPDFSecurity);
                }

                Exit(ExitCode.Success);
            }
        }

        private static void Exit(ExitCode exit) => Environment.Exit((int)exit);

        // Add any bookmarks returned by the conversion process
        private static void AddPDFBookmarks(String generatedFile, List<PDFBookmark> bookmarks, ArgParser options, PdfOutline parent)
        {
            var hasParent = parent != null;
            if (options.verbose)
            {
                Console.WriteLine("Adding {0} bookmarks {1}", bookmarks.Count, hasParent ? "as a sub-bookmark" : "to the PDF");
            }

            var srcPdf = hasParent ? parent.Owner : OpenPDFFile(generatedFile, options);
            if (srcPdf != null)
            {
                foreach (var bookmark in bookmarks)
                {
                    var page = srcPdf.Pages[bookmark.page - 1];
                    // Work out what level to add the bookmark
                    var outline = hasParent ? parent.Outlines.Add(bookmark.title, page) : srcPdf.Outlines.Add(bookmark.title, page);
                    if (bookmark.children != null && bookmark.children.Count > 0)
                    {
                        AddPDFBookmarks(generatedFile, bookmark.children, options, outline);
                    }
                }
                if (!hasParent)
                {
                    srcPdf.Save(generatedFile);
                }
            }
        }

        // There can be issues if we're copying the generated PDF to a location
        // that may lock files (e.g. for virus scanning) that prevents us from 
        // immediately opening the PDF in PDFSharp
        private static PdfDocument OpenPDFFile(string file, ArgParser options, PdfDocumentOpenMode mode = PdfDocumentOpenMode.Modify, string password = null)
        {
            int tries = 10;
            while (tries-- > 0)
            {
                try
                {
                    if (options.verbose)
                    {
                        Console.WriteLine("Opening {0} in PDF Reader", file);
                    }
                    if (password == null)
                    {
                       return PdfReader.Open(file, mode);
                    }
                    else
                    {
                        return PdfReader.Open(file, password, mode);
                    }
                }
                catch (System.UnauthorizedAccessException)
                {
                    if (options.verbose)
                    {
                        Console.WriteLine("Re-trying PDF open of {0}", file);
                    }
                    Thread.Sleep(500);
                }
            }
            return null;
        }

        // Perform some post-processing on the generated PDF
        private static void PostProcessPDFFile(String generatedFile, String finalFile, ArgParser options, Boolean postProcessPDFSecurity)
        {
            // Handle PDF merging
            if ((MergeMode)options["pdf_merge"] != MergeMode.None)
            {
                if (options.verbose)
                {
                    Console.WriteLine("Merging with existing PDF");
                }
                PdfDocument srcDoc;
                PdfDocument dstDoc = null;
                if ((MergeMode)options["pdf_merge"] == MergeMode.Append)
                {
                    srcDoc = OpenPDFFile(generatedFile, options, PdfDocumentOpenMode.Import);
                    dstDoc = ReadExistingPDFDocument(finalFile, generatedFile, ((string)options["pdf_owner_pass"]).Trim(), PdfDocumentOpenMode.Modify, options);
                }
                else
                {
                    dstDoc = OpenPDFFile(generatedFile, options);
                    srcDoc = ReadExistingPDFDocument(finalFile, generatedFile, ((string)options["pdf_owner_pass"]).Trim(), PdfDocumentOpenMode.Import, options);
                }
                int pages = srcDoc.PageCount;
                for (int pi = 0; pi < pages; pi++)
                {
                    PdfPage page = srcDoc.Pages[pi];
                    dstDoc.AddPage(page);
                }
                dstDoc.Save(finalFile);
                File.Delete(generatedFile);
            }

            if (options["pdf_page_mode"] != null || options["pdf_layout"] != null ||
                (MetaClean)options["pdf_clean_meta"] != MetaClean.None || postProcessPDFSecurity)
            {

                PdfDocument pdf = OpenPDFFile(finalFile, options);

                if (options["pdf_page_mode"] != null)
                {
                    if (options.verbose)
                    {
                        Console.WriteLine("Setting PDF Page mode");
                    }
                    pdf.PageMode = (PdfPageMode)options["pdf_page_mode"];
                }
                if (options["pdf_layout"] != null)
                {
                    if (options.verbose)
                    {
                        Console.WriteLine("Setting PDF layout");
                    }
                    pdf.PageLayout = (PdfPageLayout)options["pdf_layout"];
                }

                if ((MetaClean)options["pdf_clean_meta"] != MetaClean.None)
                {
                    if (options.verbose)
                    {
                        Console.WriteLine("Cleaning PDF meta-data");
                    }
                    pdf.Info.Creator = "";
                    pdf.Info.Keywords = "";
                    pdf.Info.Author = "";
                    pdf.Info.Subject = "";
                    //pdf.Info.Producer = "";
                    if ((MetaClean)options["pdf_clean_meta"] == MetaClean.Full)
                    {
                        pdf.Info.Title = "";
                        pdf.Info.CreationDate = System.DateTime.Today;
                        pdf.Info.ModificationDate = System.DateTime.Today;
                    }
                }

                // See if there are security changes needed
                if (postProcessPDFSecurity)
                {
                    PdfSecuritySettings secSettings = pdf.SecuritySettings;
                    if (((string)options["pdf_owner_pass"]).Trim().Length != 0)
                    {
                        
                        // Set the owner password
                        if (options.verbose)
                        {
                            Console.WriteLine("Setting PDF owner password");
                        }
                        secSettings.OwnerPassword = ((string)options["pdf_owner_pass"]).Trim();
                    }
                    if (((string)options["pdf_user_pass"]).Trim().Length != 0)
                    {
                        // Set the user password
                        // Set the owner password
                        if (options.verbose)
                        {
                            Console.WriteLine("Setting PDF user password");
                        }
                        secSettings.UserPassword = ((string)options["pdf_user_pass"]).Trim();
                    }

                    secSettings.PermitAccessibilityExtractContent = !(Boolean)options["pdf_restrict_accessibility_extraction"];
                    secSettings.PermitAnnotations = !(Boolean)options["pdf_restrict_annotation"];
                    secSettings.PermitAssembleDocument = !(Boolean)options["pdf_restrict_assembly"];
                    secSettings.PermitExtractContent = !(Boolean)options["pdf_restrict_extraction"];
                    secSettings.PermitFormsFill = !(Boolean)options["pdf_restrict_forms"];
                    secSettings.PermitModifyDocument = !(Boolean)options["pdf_restrict_modify"];
                    secSettings.PermitPrint = !(Boolean)options["pdf_restrict_print"];
                    secSettings.PermitFullQualityPrint = !(Boolean)options["pdf_restrict_full_quality"];
                }
                pdf.Save(finalFile);
                pdf.Close();
            }
        }

        static void CheckOptionIsInteger(ref Hashtable options, string optionKey, string optionName, string optionValue)
        {
            if (Regex.IsMatch(optionValue, @"^\d+$"))
            {
                options[optionKey] = (int)Convert.ToInt32(optionValue);
            }
            else
            {
                Console.WriteLine("{0} ({1}) is invalid", optionName, optionValue);
                Exit(ExitCode.Failed | ExitCode.InvalidArguments);
            }
        }

        static PdfDocument ReadExistingPDFDocument(String filename, String generatedFilename, String password, PdfDocumentOpenMode mode, ArgParser options)
        {
            PdfDocument dstDoc = null;
            try
            {

                dstDoc = OpenPDFFile(filename, options, mode);
            }
            catch (PdfReaderException)
            {
                if (password.Length != 0)
                {
                    try
                    {
                        dstDoc = OpenPDFFile(filename, options, mode, password);
                    }
                    catch (PdfReaderException)
                    {
                        if (File.Exists(generatedFilename))
                        {
                            File.Delete(generatedFilename);
                        }
                        Console.WriteLine("Unable to modify a protected PDF. Invalid owner password given.");
                        // Return the general failure code and the specific failure code
                        Exit(ExitCode.PDFProtectedDocument);
                    }
                }
                else
                {
                    // There is no owner password, we can not open this
                    if (File.Exists(generatedFilename))
                    {
                        File.Delete(generatedFilename);
                    }
                    Console.WriteLine("Unable to modify a protected PDF. Use the /pdf_owner_pass option to specify an owner password.");
                    // Return the general failure code and the specific failure code
                    Exit(ExitCode.PDFProtectedDocument);
                }
            }
            return dstDoc;
        }

        private static Dictionary<string, bool> GetInstalledPrinters()
        {
            Dictionary<string, bool> printers = new Dictionary<string, bool>();
            try
            {
                foreach (string name in PrinterSettings.InstalledPrinters)
                {
                    printers[name.ToLowerInvariant()] = true;
                }
            }
            catch (Exception) { }
            return printers;
        }
    } 
}
