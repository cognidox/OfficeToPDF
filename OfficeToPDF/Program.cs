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
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Security.AccessControl;
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
            string[] files = new string[2];
            int filesSeen = 0;
            Boolean postProcessPDF = false;
            Boolean postProcessPDFSecurity = false;
            Hashtable options = new Hashtable();
            List<PDFBookmark> documentBookmarks = new List<PDFBookmark>();

            // Loop through the input, grabbing switches off the command line
            options["hidden"] = false;
            options["markup"] = false;
            options["readonly"] = false;
            options["bookmarks"] = false;
            options["print"] = true;
            options["screen"] = false;
            options["pdfa"] = false;
            options["verbose"] = false;
            options["excludeprops"] = false;
            options["excludetags"] = false;
            options["noquit"] = true;
            options["merge"] = false;
            options["template"] = "";
            options["password"] = "";
            options["printer"] = "";
            options["fallback_printer"] = "";
            options["working_dir"] = "";
            options["has_working_dir"] = false;
            options["excel_show_formulas"] = false;
            options["excel_show_headings"] = false;
            options["excel_auto_macros"] = false;
            options["excel_template_macros"] = false;
            options["excel_active_sheet"] = false;
            options["excel_no_link_update"] = false;
            options["excel_no_recalculate"] = false;
            options["excel_max_rows"] = (int) 0;
            options["excel_worksheet"] = (int) 0;
            options["excel_delay"] = (int) 0;
            options["word_field_quick_update"] = false;
            options["word_field_quick_update_safe"] = false;
            options["word_no_field_update"] = false;
            options["word_header_dist"] = (float) -1;
            options["word_footer_dist"] = (float) -1;
            options["word_max_pages"] = (int) 0;
            options["word_ref_fonts"] = false;
            options["word_keep_history"] = false;
            options["word_no_repair"] = false;
            options["word_show_comments"] = false;
            options["word_show_revs_comments"] = false;
            options["word_show_format_changes"] = false;
            options["word_show_ink_annot"] = false;
            options["word_show_ins_del"] = false;
            options["word_markup_balloon"] = false;
            options["word_show_all_markup"] = false;
            options["word_fix_table_columns"] = false;
            options["original_filename"] = "";
            options["original_basename"] = "";
            options["powerpoint_output"] = "";
            options["pdf_page_mode"] = null;
            options["pdf_layout"] = null;
            options["pdf_merge"] = (int) MergeMode.None;
            options["pdf_clean_meta"] = (int)MetaClean.None;
            options["pdf_owner_pass"] = "";
            options["pdf_user_pass"] = "";
            options["pdf_restrict_annotation"] = false;
            options["pdf_restrict_extraction"] = false;
            options["pdf_restrict_assembly"] = false;
            options["pdf_restrict_forms"] = false;
            options["pdf_restrict_modify"] = false;
            options["pdf_restrict_print"] = false;
            options["pdf_restrict_annotation"] = false;
            options["pdf_restrict_accessibility_extraction"] = false;
            options["pdf_restrict_full_quality"] = false;

            // We need some printers to keep office happy
            Dictionary<string,bool> installedPrinters = GetInstalledPrinters();
            if (installedPrinters.Count <= 0)
            {
                Console.WriteLine("There are no installed printers, so conversion can not proceed");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.NoPrinters));
            }

            // Strings used in error messages for different options
            var optionNameMap = new Dictionary<string, string>()
            {
                { "excel_max_rows", "Maximum number of rows" },
                { "excel_worksheet", "Excel worksheet" },
                { "word_max_pages", "Maximum number of pages" },
                { "excel_delay", "Excel delay milliseconds" }
            };

            Regex switches = new Regex(@"^/(version|hidden|markup|readonly|bookmarks|merge|noquit|print|(fallback_)?printer|screen|pdfa|template|writepassword|password|help|verbose|exclude(props|tags)|excel_(delay|max_rows|show_formulas|show_headings|auto_macros|template_macros|active_sheet|worksheet|no_recalculate|no_link_update)|powerpoint_(output)|word_(header_dist|footer_dist|ref_fonts|no_field_update|field_quick_update(_safe)?|max_pages|keep_history|no_repair|fix_table_columns|show_(comments|revs_comments|format_changes|ink_annot|ins_del|all_markup)|markup_balloon)|pdf_(page_mode|append|prepend|layout|clean_meta|owner_pass|user_pass|restrict_(annotation|extraction|assembly|forms|modify|print|accessibility_extraction|full_quality))|working_dir|\?)$", RegexOptions.IgnoreCase);
            for (int argIdx = 0; argIdx < args.Length; argIdx++)
            {
                string item = args[argIdx];
                // see if this starts with a /
                Match m = Regex.Match(item, @"^/");
                if (m.Success)
                {
                    // This is an option
                    Match itemMatch = switches.Match(item);
                    if (itemMatch.Success)
                    {
                        if (itemMatch.Groups[1].Value.ToLower().Equals("help") ||
                            itemMatch.Groups[1].Value.Equals("?"))
                        {
                            ShowHelp();
                        }
                        switch (itemMatch.Groups[1].Value.ToLower())
                        {
                            case "pdf_page_mode":
                                if (argIdx + 1 < args.Length)
                                {
                                    postProcessPDF = true;
                                    var pageMode = args[argIdx + 1];
                                    pageMode = pageMode.ToLower();
                                    switch (pageMode)
                                    {
                                        case "full":
                                            options["pdf_page_mode"] = PdfPageMode.FullScreen;
                                            break;
                                        case "none":
                                            options["pdf_page_mode"] = PdfPageMode.UseNone;
                                            break;
                                        case "bookmarks":
                                            options["pdf_page_mode"] = PdfPageMode.UseOutlines;
                                            break;
                                        case "thumbs":
                                            options["pdf_page_mode"] = PdfPageMode.UseThumbs;
                                            break;
                                        default:
                                            Console.WriteLine("Invalid PDF page mode ({0}). It must be one of full, none, outline or thumbs", args[argIdx + 1]);
                                            Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                            break;
                                    }
                                    argIdx++;
                                }
                                break;
                            case "pdf_clean_meta":
                                if (argIdx + 1 < args.Length)
                                {
                                    postProcessPDF = true;
                                    var cleanType = args[argIdx + 1];
                                    cleanType = cleanType.ToLower();
                                    switch (cleanType)
                                    {
                                        case "basic":
                                            options["pdf_clean_meta"] = MetaClean.Basic;
                                            break;
                                        case "full":
                                            options["pdf_clean_meta"] = MetaClean.Full;
                                            break;
                                        default:
                                            Console.WriteLine("Invalid PDF meta-data clean value ({0}). It must be one of full or basic", args[argIdx + 1]);
                                            Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                            break;
                                    }
                                    argIdx++;
                                }
                                break;
                            case "pdf_layout":
                                if (argIdx + 1 < args.Length)
                                {
                                    postProcessPDF = true;
                                    var pdfLayout = args[argIdx + 1];
                                    pdfLayout = pdfLayout.ToLower();
                                    switch (pdfLayout)
                                    {
                                        case "onecol":
                                            options["pdf_layout"] = PdfPageLayout.OneColumn;
                                            break;
                                        case "single":
                                            options["pdf_layout"] = PdfPageLayout.SinglePage;
                                            break;
                                        case "twocolleft":
                                            options["pdf_layout"] = PdfPageLayout.TwoColumnLeft;
                                            break;
                                        case "twocolright":
                                            options["pdf_layout"] = PdfPageLayout.TwoColumnRight;
                                            break;
                                        case "twopageleft":
                                            options["pdf_layout"] = PdfPageLayout.TwoPageLeft;
                                            break;
                                        case "twopageright":
                                            options["pdf_layout"] = PdfPageLayout.TwoPageRight;
                                            break;
                                        default:
                                            Console.WriteLine("Invalid PDF layout ({0}). It must be one of onecol, single, twocolleft, twocolright, twopageleft or twopageright", args[argIdx + 1]);
                                            Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                            break;
                                    }
                                    argIdx++;
                                }
                                break;
                            case "pdf_owner_pass":
                            case "pdf_user_pass":
                                if (argIdx + 1 < args.Length)
                                {
                                    postProcessPDF = true;
                                    postProcessPDFSecurity = true;
                                    var pass = args[argIdx + 1];
                                    // Set the password
                                    options[itemMatch.Groups[1].Value.ToLower()] = pass;
                                    argIdx++;
                                }
                                break;
                            case "template":
                                // Only accept the next option if there are enough options
                                if (argIdx + 1 < args.Length)
                                {
                                    if (File.Exists(args[argIdx + 1]))
                                    {
                                        FileInfo templateInfo = new FileInfo(args[argIdx + 1]);
                                        options[itemMatch.Groups[1].Value.ToLower()] = templateInfo.FullName;
                                    }
                                    else
                                    {
                                        Console.WriteLine("Unable to find {0} {1}", itemMatch.Groups[1].Value.ToLower(), args[argIdx + 1]);
                                    }
                                    argIdx++;
                                }
                                break;
                            case "powerpoint_output":
                                // Only accept the next option if there are enough options
                                if (argIdx + 1 < args.Length)
                                {
                                    bool validOutputType = false;
                                    PowerpointConverter.GetOutputType(args[argIdx + 1], ref validOutputType);
                                    if (!validOutputType)
                                    {
                                        Console.WriteLine("Invalid PowerPoint output type");
                                        Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                    }
                                    options["powerpoint_output"] = args[argIdx + 1];
                                    argIdx++;
                                }
                                break;
                            case "working_dir":
                                // Allow for a local working directory where files are manipulated
                                if (argIdx + 1 < args.Length)
                                {
                                    if (Directory.Exists(args[argIdx + 1]))
                                    {
                                        // Need to check the directory is writable
                                        bool workingDirectoryWritable = false;
                                        try
                                        {
                                            AuthorizationRuleCollection arc = Directory.GetAccessControl(args[argIdx + 1]).GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));
                                            foreach (FileSystemAccessRule rule in arc)
                                            {
                                                if (rule.AccessControlType == AccessControlType.Allow)
                                                {
                                                    workingDirectoryWritable = true;
                                                    break;
                                                }
                                            }
                                        }
                                        catch (Exception)
                                        {}
                                        if (workingDirectoryWritable)
                                        {
                                            int maxTries = 20;
                                            while (maxTries-- > 0)
                                            {
                                                options["working_dir"] = Path.Combine(args[argIdx + 1], Guid.NewGuid().ToString());
                                                if (!Directory.Exists((string)options["working_dir"]))
                                                {
                                                    DirectoryInfo di = Directory.CreateDirectory((string)options["working_dir"]);
                                                    if (di.Exists)
                                                    {
                                                        options["has_working_dir"] = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (maxTries <= 0)
                                            {
                                                Console.WriteLine("A working directory can not be created");
                                                Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                            }
                                        }
                                        else
                                        {
                                            // The working directory must be writable
                                            Console.WriteLine("The working directory {0} is not writable", args[argIdx + 1]);
                                            Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                        }
                                    }
                                    else
                                    {
                                        // We need a real directory to work in, so there is an error here
                                        Console.WriteLine("Unable to find working directory {0}", args[argIdx + 1]);
                                        Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                    }
                                    argIdx++;
                                }
                                break;
                            case "excel_max_rows":
                            case "excel_worksheet":
                            case "excel_delay":
                            case "word_max_pages":
                                // Only accept the next option if there are enough options
                                if (argIdx + 1 < args.Length)
                                {
                                    CheckOptionIsInteger(ref options, itemMatch.Groups[1].Value.ToLower(), optionNameMap[itemMatch.Groups[1].Value.ToLower()], args[argIdx + 1]);
                                    argIdx++;
                                }
                                break;
                            case "word_header_dist":
                            case "word_footer_dist":
                                // Only accept the next option if there are enough options
                                if (argIdx + 1 < args.Length)
                                {
                                    if (Regex.IsMatch(args[argIdx + 1], @"^[\d\.]+$"))
                                    {
                                        try
                                        {
                                            options[itemMatch.Groups[1].Value.ToLower()] = (float)Convert.ToDouble(args[argIdx + 1]);
                                        }
                                        catch (Exception)
                                        {
                                            Console.WriteLine("Header/Footer distance ({0}) is invalid", args[argIdx + 1]);
                                            Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Header/Footer distance ({0}) is invalid", args[argIdx + 1]);
                                        Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                    }
                                    argIdx++;
                                }
                                break;
                            case "password":
                            case "writepassword":
                            case "printer":
                            case "fallback_printer":
                                // Only accept the next option if there are enough options
                                string optname = itemMatch.Groups[1].Value.ToLower();
                                if (argIdx + 1 < args.Length)
                                {
                                    options[optname] = args[argIdx + 1];
                                    argIdx++;
                                }
                                if (optname.Equals("printer") || optname.Equals("fallback_printer"))
                                {
                                    if (!installedPrinters.ContainsKey(((string)options[optname]).ToLowerInvariant())) {
                                        // The requested printer did not exists
                                        Console.WriteLine("The printer \"{0}\" is not installed", options[optname]);
                                        Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                    }
                                }
                                break;
                            case "screen":
                                options["print"] = false;
                                options[itemMatch.Groups[1].Value.ToLower()] = true;
                                break;
                            case "print":
                                options["screen"] = false;
                                options[itemMatch.Groups[1].Value.ToLower()] = true;
                                break;
                            case "version":
                                Assembly asm = Assembly.GetExecutingAssembly();
                                FileVersionInfo fv = System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                                Console.WriteLine(String.Format("{0}", fv.FileVersion));
                                Environment.Exit((int)ExitCode.Success);
                                break;
                            case "pdf_append":
                                if ((MergeMode)options["pdf_merge"] != MergeMode.None)
                                {
                                    Console.WriteLine("Only one of /pdf_append or /pdf_prepend can be used");
                                    Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                }
                                postProcessPDF = true;
                                options["pdf_merge"] = MergeMode.Append;
                                break;
                            case "pdf_prepend":
                                if ((MergeMode)options["pdf_merge"] != MergeMode.None)
                                {
                                    Console.WriteLine("Only one of /pdf_append or /pdf_prepend can be used");
                                    Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                }
                                postProcessPDF = true;
                                options["pdf_merge"] = MergeMode.Prepend;
                                break;
                            case "pdf_restrict_annotation":
                            case "pdf_restrict_extraction":
                            case "pdf_restrict_assembly":
                            case "pdf_restrict_forms":
                            case "pdf_restrict_modify":
                            case "pdf_restrict_print":
                            case "pdf_restrict_full_quality":
                            case "pdf_restrict_accessibility_extraction":
                                postProcessPDFSecurity = true;
                                options[itemMatch.Groups[1].Value.ToLower()] = true;
                                break;
                            default:
                                options[itemMatch.Groups[1].Value.ToLower()] = true;
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Unknown option: {0}", item);
                        Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                    }
                }
                else if (filesSeen < 2)
                {
                    files[filesSeen++] = item;
                }
            }

            // Need to error here, as we need input and output files as the
            // arguments to this script
            if (filesSeen != 1 && filesSeen != 2)
            {
                ShowHelp();
            }

            // Make sure we only choose one of /screen or /print options
            if ((Boolean)options["screen"] && (Boolean)options["print"])
            {
                Console.WriteLine("You can only use one of /screen or /print - not both");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
            }

            // Make sure the input file looks like something we can handle (i.e. has an office
            // filename extension)
            Regex fileMatch = new Regex(@"\.(((ppt|pps|pot|do[ct]|xls|xlt)[xm]?)|xps|xlsb|od[spt]|rtf|csv|vsd[xm]?|vd[xw]|em[fz]|dwg|dxf|wmf|pub|msg|vcf|ics|mpp|svg|txt|html?|wpd)$", RegexOptions.IgnoreCase);
            if (fileMatch.Matches(files[0]).Count != 1)
            {
                Console.WriteLine("Input file can not be handled. Must be Word, PowerPoint, Excel, Outlook, Publisher, XPS or Visio");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.UnsupportedFileFormat));
            }

            if (filesSeen == 1)
            {
                // If only one file is seen, we just swap the extension
                files[1] = Path.ChangeExtension(files[0], "pdf");
            }
            else
            {
                // If the second file is a directory, then we want to create the PDF
                // with the same name as the original (changing the extension to pdf),
                // but in the directory given by the path
                if (Directory.Exists(files[1]))
                {
                    files[1] = Path.Combine(files[1], Path.GetFileNameWithoutExtension(files[0]) + ".pdf");
                }
            }

            String inputFile = "";
            String outputFile = "";
            String finalOutputFile = "";

            // Make sure the input file exists and is readable
            FileInfo info;
            try
            {
                info = new FileInfo(files[0]);

                if (info == null || !info.Exists)
                {
                    Console.WriteLine("Input file doesn't exist");
                    Environment.Exit((int)(ExitCode.Failed | ExitCode.FileNotFound));
                }
                inputFile = info.FullName;
                options["original_filename"] = info.Name;
                options["original_basename"] = info.Name.Substring(0, info.Name.Length - info.Extension.Length);
            }
            catch
            {
                Console.WriteLine("Unable to open input file");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.FileOpenFailure));
            }

            // Stop people using the template as the input file
            if (!String.IsNullOrEmpty((string)options["template"]) &&
                inputFile.Equals((string)options["template"]))
            {
                Console.WriteLine("Input file must be different from the template file");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
            }

            // Make sure the destination location exists
            FileInfo outputInfo = new FileInfo(files[1]);
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
                Environment.Exit((int)(ExitCode.Failed | ExitCode.DirectoryNotFound));
            }

            if (!System.IO.Directory.Exists(outputInfo.DirectoryName))
            {
                Console.WriteLine("Output directory does not exist");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.DirectoryNotFound));
            }

            // We want the input and output files copied to a working area where
            // we can manipulate them
            if ((bool)options["has_working_dir"])
            {
                // Create a local temporary area and put the input and output in separate
                // areas
                string workingInput = Path.Combine((string)options["working_dir"], "input");
                string workingOutput = Path.Combine((string)options["working_dir"], "output");
                System.IO.Directory.CreateDirectory(workingInput);
                System.IO.Directory.CreateDirectory(workingOutput);
                string workingSource = Path.Combine(workingInput, Path.GetFileName(inputFile));
                string workingDest = Path.Combine(workingOutput, Path.GetFileName(outputFile));
                options["working_orig_dest"] = outputFile;
                File.Copy(inputFile, workingSource);
                inputFile = workingSource;
                outputFile = workingDest;
                if ((Boolean)options["verbose"])
                {
                    Console.WriteLine("Created working directory: {0}", (string)options["working_dir"]);
                }
            }

            // Now, do the cleverness of determining what the extension is, and so, which
            // conversion class to pass it to
            int converted = (int)ExitCode.UnknownError;
            Match extMatch = fileMatch.Match(inputFile);
            if (extMatch.Success)
            {
                // Set and environment variable so Office application VBA
                // code can check for un-attended conversion and avoid showing
                // blocking dialogs
                Environment.SetEnvironmentVariable("OFFICE2PDF_AUTO_CONVERT", "1");

                if ((Boolean)options["verbose"])
                {
                    Console.WriteLine("Converting {0} to {1}", inputFile, finalOutputFile);
                }
                switch (extMatch.Groups[1].ToString().ToLower())
                {
                    case "rtf":
                    case "odt":
                    case "doc":
                    case "dot":
                    case "docx":
                    case "dotx":
                    case "docm":
                    case "dotm":
                    case "txt":
                    case "html":
                    case "htm":
                    case "wpd":
                        // Word
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with Word converter");
                        }
                        converted = WordConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "csv":
                    case "ods":
                    case "xls":
                    case "xlsx":
                    case "xlt":
                    case "xltx":
                    case "xlsm":
                    case "xltm":
                    case "xlsb":
                        // Excel
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with Excel converter");
                        }
                        converted = ExcelConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "odp":
                    case "ppt":
                    case "pptx":
                    case "pptm":
                    case "pot":
                    case "potm":
                    case "potx":
                    case "pps":
                    case "ppsx":
                    case "ppsm":
                        // Powerpoint
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with Powerpoint converter");
                        }
                        converted = PowerpointConverter.Convert(inputFile, outputFile, options, ref documentBookmarks);
                        break;
                    case "vsd":
                    case "vsdm":
                    case "vsdx":
                    case "vdx":
                    case "vdw":
                    case "svg":
                    case "emf":
                    case "emz":
                    case "dwg":
                    case "dxf":
                    case "wmf":
                        // Visio
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with Visio converter");
                        }
                        converted = VisioConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "pub":
                        // Publisher
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with Publisher converter");
                        }
                        converted = PublisherConverter.Convert(inputFile, outputFile, options, ref documentBookmarks);
                        break;
                    case "xps":
                        // XPS
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with XPS converter");
                        }
                        converted = XpsConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "msg":
                    case "vcf":
                    case "ics":
                        // Outlook
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with Outlook converter");
                        }
                        converted = OutlookConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "mpp":
                        // Project
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Converting with Project converter");
                        }
                        converted = ProjectConverter.Convert(inputFile, outputFile, options);
                        break;
                }
            }

            // Clear up the working directory and restore the expected output
            if ((bool)options["has_working_dir"])
            {
                if (File.Exists(outputFile))
                {
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Copying working file {0} to {1}", outputFile, (string)options["working_orig_dest"]);
                    }
                    File.Copy(outputFile, (string)options["working_orig_dest"]);
                    outputFile = (string)options["working_orig_dest"];
                    
                }
                Directory.Delete((string)options["working_dir"], true);
            }

            if (converted != (int)ExitCode.Success)
            {
                Console.WriteLine("Did not convert");
                // Return the general failure code and the specific failure code
                Environment.Exit((int)ExitCode.Failed | converted);
            }
            else
            {
                if ((Boolean)options["verbose"])
                {
                    Console.WriteLine("Completed Conversion");
                }

                if (documentBookmarks.Count > 0)
                {
                    AddPDFBookmarks(outputFile, documentBookmarks, options, null);
                }
                
                // Determine if we have to post-process the PDF
                if (postProcessPDF)
                {
                    PostProcessPDFFile(outputFile, finalOutputFile, options, postProcessPDFSecurity);
                }

                Environment.Exit((int)ExitCode.Success);
            }
        }

        // Add any bookmarks returned by the conversion process
        private static void AddPDFBookmarks(String generatedFile, List<PDFBookmark> bookmarks, Hashtable options, PdfOutline parent)
        {
            var hasParent = parent != null;
            if ((Boolean)options["verbose"])
            {
                Console.WriteLine("Adding {0} bookmarks {1}", bookmarks.Count, (hasParent ? "as a sub-bookmark" : "to the PDF"));
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
        private static PdfDocument OpenPDFFile(string file, Hashtable options, PdfDocumentOpenMode mode = PdfDocumentOpenMode.Modify, string password = null)
        {
            int tries = 10;
            while (tries-- > 0)
            {
                try
                {
                    if ((Boolean)options["verbose"])
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
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Re-trying PDF open of {0}", file);
                    }
                    Thread.Sleep(500);
                }
            }
            return null;
        }

        // Perform some post-processing on the generated PDF
        private static void PostProcessPDFFile(String generatedFile, String finalFile, Hashtable options, Boolean postProcessPDFSecurity)
        {
            // Handle PDF merging
            if ((MergeMode)options["pdf_merge"] != MergeMode.None)
            {
                if ((Boolean)options["verbose"])
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
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Setting PDF Page mode");
                    }
                    pdf.PageMode = (PdfPageMode)options["pdf_page_mode"];
                }
                if (options["pdf_layout"] != null)
                {
                    if ((Boolean)options["verbose"])
                    {
                        Console.WriteLine("Setting PDF layout");
                    }
                    pdf.PageLayout = (PdfPageLayout)options["pdf_layout"];
                }

                if ((MetaClean)options["pdf_clean_meta"] != MetaClean.None)
                {
                    if ((Boolean)options["verbose"])
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
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Setting PDF owner password");
                        }
                        secSettings.OwnerPassword = ((string)options["pdf_owner_pass"]).Trim();
                    }
                    if (((string)options["pdf_user_pass"]).Trim().Length != 0)
                    {
                        // Set the user password
                        // Set the owner password
                        if ((Boolean)options["verbose"])
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
                Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
            }
        }

        static PdfDocument ReadExistingPDFDocument(String filename, String generatedFilename, String password, PdfDocumentOpenMode mode, Hashtable options)
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
                        Environment.Exit((int)ExitCode.PDFProtectedDocument);
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
                    Environment.Exit((int)ExitCode.PDFProtectedDocument);
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

        static void ShowHelp()
        {
            Console.Write(@"Converts Office documents to PDF from the command line.
Handles Office files:
  doc, dot, docx, dotx, docm, dotm, rtf, odt, txt, htm, html, wpd, ppt, pptx,
  pptm, pps, ppsx, ppsm, pot, potm, potx, odp, xls, xlsx, xlsm, xlt, xltm,
  xltx, xlsb, csv, ods, vsd, vsdm, vsdx, svg, vdx, vdw, emf, emz, dwg, dxf, wmf,
  pub, mpp, ics, vcf, msg, xps

OfficeToPDF.exe [switches] input_file [output_file]

  Available switches:

  /bookmarks    - Create bookmarks in the PDF when they are supported by the
                  Office application
  /hidden       - Attempt to hide the Office application window when converting
  /markup       - Show document markup when creating PDFs with Word
  /readonly     - Load the input file in read only mode where possible
  /print        - Create high-quality PDFs optimised for print (default)
  /screen       - Create PDFs optimised for screen display
  /pdfa         - Produce ISO 19005-1 (PDF/A) compliant PDFs
  /excludeprops - Do not include properties in the PDF
  /excludetags  - Do not include tags in the PDF
  /noquit       - Do not quit already running Office applications once the conversion is done
  /verbose      - Print out messages as it runs
  /password <pass>          - Use <pass> as the password to open the document with
  /writepassword <pass>     - Use <pass> as the write password to open the document with
  /template <template_path> - Use a .dot, .dotx or .dotm template when
                              converting with Word. When converting Excel, use the
                              page settings from the first worksheet in the template
                              document
  /fallback_printer <name>  - Convert by printing the document to the postscript printer with name
                              <name>. Only operates if exporting from Word fails.
  /printer <name>           - Convert by printing the document to the postscript printer with name
                              <name>. 
  /excel_active_sheet       - Only convert the active worksheet
  /excel_auto_macros        - Run Auto_Open macros in Excel files before conversion
  /excel_show_formulas      - Show formulas in the generated PDF
  /excel_delay <ms>         - Number of milliseconds to pause Excel for during file processing
  /excel_show_headings      - Show row and column headings
  /excel_max_rows <rows>    - If any worksheet in a spreadsheet document has more
                              than this number of rows, do not attempt to convert
                              the file. Applies when converting with Excel.
  /excel_no_link_update     - Do not update links when opening Excel files.
  /excel_no_recalculate     - Skip automatic re-calculation of formulas in the workbook.
  /excel_template_macros    - Run Auto_Open macros in the /template document before conversion by Excel
  /excel_worksheet <num>    - Only convert worksheet <num> in the workbook. First sheet is 1
  /powerpoint_output <type> - Controls what is generated by output. Possible values are slides, notes,
                              outline, build_slides, handouts and multi-page handouts using handout2,
                              handout3, handout4, handout6 and handout9. The default is slides.   
  /word_header_dist <pts>   - The distance (in points) from the header to the top of
                              the page.
  /word_footer_dist <pts>   - The distance (in points) from the footer to the bottom
                              of the page.
  /word_field_quick_update  - Perform a fast update of fields in Word before conversion.
  /word_field_quick_update_safe - Perform a fast update for fields only if there are no broken linked files.
  /word_fix_table_columns   - Fix table column widths in cases where table body columns do not match header
                              column widths.
  /word_keep_history        - Do not clear Word's recent files list.
  /word_max_pages <pages>   - Do not attempt conversion of a Word document if it has more than
                              this number of pages.
  /word_no_field_update     - Do not update fields when creating the PDF.
  /word_no_repair           - Do not attempt to repair a Word document when opening.
  /word_ref_fonts           - When fonts are not available, a reference to the font is used in
                              the generated PDF rather than a bitmapped version. The default is
                              for a bitmap of the text to be used.
  /word_show_comments       - Show comments when /markup is used.
  /word_show_revs_comments  - Show revisions and comments when /markup is used.
  /word_show_format_changes - Show format changes when /markup is used.
  /word_show_ink_annot      - Show ink annotations when /markup is used.
  /word_show_ins_del        - Show all markup when /markup is used.
  /word_show_all_markup     - Show all markup content when /markup is used.
  /word_markup_balloon      - Show balloon style markup messages rather than inline ones.
  /pdf_clean_meta <type>    - Allows for some meta-data to be removed from the generated PDF.
                              <type> can be:
                                basic - removes author, keywords, creator and subject
                                full  - removes all that basic removes and also the title
  /pdf_layout <layout>      - Controls how the pages layout in Acrobat Reader. <layout> can be
                              one of the following values:
                                onecol       - show pages as a single scrolling column
                                single       - show pages one at a time
                                twocolleft   - show pages in two columns, with odd-numbered pages on the left
                                twocolright  - show pages in two columns, with odd-numbered pages on the right
                                twopageleft  - show pages two at a time, with odd-numbered pages on the left
                                twopageright - show pages two at a time, with odd-numbered pages on the right
  /pdf_page_mode <mode>     - Controls how the PDF will open with Acrobat Reader. <mode> can be
                              one of the following values:
                                full      - the PDF will open in fullscreen mode
                                bookmarks - the PDF will open with the bookmarks visible
                                thumbs    - the PDF will open with the thumbnail view visible
                                none      - the PDF will open without the navigation bar visible
  /pdf_append                - Append the generated PDF to the end of the PDF destination.
  /pdf_prepend               - Prepend the generated PDF to the start of the PDF destination.
  /pdf_owner_pass <password> - Set the owner password on the PDF. Needed to make modifications to the PDF.
  /pdf_user_pass  <password> - Set the user password on the PDF. Needed to open the PDF.
  /pdf_restrict_accessibility_extraction - Prevent all content extraction without the owner password.
  /pdf_restrict_annotation   - Prevent annotations on the PDF without the owner password.
  /pdf_restrict_assembly     - Prevent rotation, removal or insertion of pages without the owner password.
  /pdf_restrict_extraction   - Prevent content extraction without the owner password.
  /pdf_restrict_forms        - Prevent form entry without the owner password.
  /pdf_restrict_full_quality - Prevent full quality printing without the owner password.
  /pdf_restrict_modify       - Prevent modification without the owner password.
  /pdf_restrict_print        - Prevent printing without the owner password.
  /version                   - Print out the version of OfficeToPDF and exit.
  /working_dir <path>        - A path to copy the input file into temporarily when running the conversion.
  
  input_file  - The filename of the Office document to convert
  output_file - The filename of the PDF to create. If not given, the input filename
                will be used with a .pdf extension
");
            Environment.Exit((int)ExitCode.Success);
        }
    } 
}
