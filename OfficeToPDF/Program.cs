/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010/2013/2016
 *  Copyright (C) 2011-2016 Cognidox Ltd
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
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp;

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
        EmptyWorksheet = 512
    }

    public enum MergeMode : int
    {
        None = 0,
        Prepend = 1,
        Append = 2
    }

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string[] files = new string[2];
            int filesSeen = 0;
            Boolean postProcessPDF = false;
            Hashtable options = new Hashtable();

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
            options["noquit"] = false;
            options["merge"] = false;
            options["template"] = "";
            options["password"] = "";
            options["excel_show_formulas"] = false;
            options["excel_show_headings"] = false;
            options["excel_auto_macros"] = false;
            options["excel_active_sheet"] = false;
            options["excel_max_rows"] = (int) 0;
            options["excel_worksheet"] = (int) 0;
            options["word_header_dist"] = (float) -1;
            options["word_footer_dist"] = (float) -1;
            options["pdf_page_mode"] = null;
            options["pdf_layout"] = null;
            options["pdf_merge"] = (int) MergeMode.None;
            Regex switches = new Regex(@"^/(version|hidden|markup|readonly|bookmarks|merge|noquit|print|screen|pdfa|template|writepassword|password|help|verbose|exclude(props|tags)|excel_max_rows|excel_show_formulas|excel_show_headings|excel_auto_macros|excel_active_sheet|excel_worksheet|word_header_dist|word_footer_dist|pdf_(page_mode|append|prepend|layout)|\?)$", RegexOptions.IgnoreCase);
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
                            showHelp();
                        }
                        switch (itemMatch.Groups[1].Value.ToLower())
                        {
                            case "pdf_page_mode":
                                if (argIdx + 2 < args.Length)
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
                            case "pdf_layout":
                                if (argIdx + 2 < args.Length)
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
                            case "template":
                                // Only accept the next option if there are enough options
                                if (argIdx + 2 < args.Length)
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
                            case "excel_max_rows":
                                // Only accept the next option if there are enough options
                                if (argIdx + 2 < args.Length)
                                {
                                    if (Regex.IsMatch(args[argIdx + 1], @"^\d+$"))
                                    {
                                        options[itemMatch.Groups[1].Value.ToLower()] = (int) Convert.ToInt32(args[argIdx + 1]);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Maximum number of rows ({0}) is invalid", args[argIdx + 1]);
                                        Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                    }
                                    argIdx++;
                                }
                                break;
                            case "excel_worksheet":
                                // Only accept the next option if there are enough options
                                if (argIdx + 2 < args.Length)
                                {
                                    if (Regex.IsMatch(args[argIdx + 1], @"^\d+$"))
                                    {
                                        options[itemMatch.Groups[1].Value.ToLower()] = (int)Convert.ToInt32(args[argIdx + 1]);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Excel worksheet ({0}) is invalid", args[argIdx + 1]);
                                        Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
                                    }
                                    argIdx++;
                                }
                                break;
                            case "word_header_dist":
                            case "word_footer_dist":
                                // Only accept the next option if there are enough options
                                if (argIdx + 2 < args.Length)
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
                                // Only accept the next option if there are enough options
                                if (argIdx + 2 < args.Length)
                                {
                                    options[itemMatch.Groups[1].Value.ToLower()] = args[argIdx + 1];
                                    argIdx++;
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
                showHelp();
            }

            // Make sure we only choose one of /screen or /print options
            if ((Boolean)options["screen"] && (Boolean)options["print"])
            {
                Console.WriteLine("You can only use one of /screen or /print - not both");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.InvalidArguments));
            }

            // Make sure the input file looks like something we can handle (i.e. has an office
            // filename extension)
            Regex fileMatch = new Regex(@"\.(((ppt|pps|pot|do[ct]|xls|xlt)[xm]?)|xlsb|od[spt]|rtf|csv|vsd[xm]?|vd[xw]|em[fz]|dwg|dxf|wmf|pub|msg|vcf|ics|mpp|svg|txt|html?|wpd)$", RegexOptions.IgnoreCase);
            if (fileMatch.Matches(files[0]).Count != 1)
            {
                Console.WriteLine("Input file can not be handled. Must be Word, PowerPoint, Excel, Outlook, Publisher or Visio");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.UnsupportedFileFormat));
            }

            if (filesSeen == 1)
            {
                files[1] = Path.ChangeExtension(files[0], "pdf");
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
            }
            catch
            {
                Console.WriteLine("Unable to open input file");
                Environment.Exit((int)(ExitCode.Failed | ExitCode.FileOpenFailure));
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


            // Now, do the cleverness of determining what the extension is, and so, which
            // conversion class to pass it to
            int converted = (int)ExitCode.UnknownError;
            Match extMatch = fileMatch.Match(inputFile);
            if (extMatch.Success)
            {
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
                        converted = PowerpointConverter.Convert(inputFile, outputFile, options);
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
                        converted = PublisherConverter.Convert(inputFile, outputFile, options);
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
                
                // Determine if we have to post-process the PDF
                if (postProcessPDF)
                {
                    // Handle PDF merging
                    if ((MergeMode)options["pdf_merge"] != MergeMode.None)
                    {
                        if ((Boolean)options["verbose"])
                        {
                            Console.WriteLine("Merging with existing PDF");
                        }
                        PdfDocument srcDoc;
                        PdfDocument dstDoc;
                        if ((MergeMode)options["pdf_merge"] == MergeMode.Append)
                        {
                            // Open the destination and generated PDFs
                            srcDoc = PdfReader.Open(outputFile, PdfDocumentOpenMode.Import);
                            dstDoc = PdfReader.Open(finalOutputFile, PdfDocumentOpenMode.Modify);
                        }
                        else
                        {
                            srcDoc = PdfReader.Open(finalOutputFile, PdfDocumentOpenMode.Import);
                            dstDoc = PdfReader.Open(outputFile, PdfDocumentOpenMode.Modify);
                        }
                        int pages = srcDoc.PageCount;
                        for (int pi = 0; pi < pages; pi++)
                        {
                            PdfPage page = srcDoc.Pages[pi];
                            dstDoc.AddPage(page);
                        }
                        dstDoc.Save(finalOutputFile);
                        File.Delete(outputFile);
                    }

                    if (options["pdf_page_mode"] != null || options["pdf_layout"] != null)
                    {
                        
                        PdfDocument pdf = PdfReader.Open(finalOutputFile, PdfDocumentOpenMode.Modify);

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
                        pdf.Save(finalOutputFile);
                        pdf.Close();
                    }
                }
                Environment.Exit((int)ExitCode.Success);
            }
        }

        static void showHelp()
        {
            Console.Write(@"Converts Office documents to PDF from the command line.
Handles Office files:
  doc, dot, docx, dotx, docm, dotm, rtf, odt, txt, htm, html, wpd, ppt, pptx,
  pptm, pps, ppsx, ppsm, pot, potm, potx, odp, xls, xlsx, xlsm, xlt, xltm,
  xltx, xlsb, csv, ods, vsd, vsdm, vsdx, svg, vdx, vdw, emf, emz, dwg, dxf, wmf,
  pub, mpp, ics, vcf, msg

OfficeToPDF.exe [/bookmarks] [/hidden] [/readonly] input_file [output_file]

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
                              converting with Word
  /excel_active_sheet       - Only convert the active worksheet
  /excel_auto_macros        - Run Auto_Open macros in Excel files before conversion
  /excel_show_formulas      - Show formulas in the generated PDF
  /excel_show_headings      - Show row and column headings
  /excel_max_rows <rows>    - If any worksheet in a spreadsheet document has more
                              than this number of rows, do not attempt to convert
                              the file. Applies when converting with Excel.
  /excel_worksheet <num>    - Only convert worksheet <num> in the workbook. First sheet is 1.
  /word_header_dist <pts>   - The distance (in points) from the header to the top of
                              the page.
  /word_footer_dist <pts>   - The distance (in points) from the footer to the bottom
                              of the page.
  /pdf_layout <layout>      - Controls how the pages layout in Acrobat Reader. <layout> can be
                              one of the following values:
                                onecol       - show pages as a single scolling column
                                single       - show pages one at a time
                                twocolleft   - show pages in two columns, with oddnumbered pages on the left
                                twocolright  - show pages in two columns, with oddnumbered pages on the right
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
  /version                   - Print out the version of OfficeToPDF and exit.
  
  input_file  - The filename of the Office document to convert
  output_file - The filename of the PDF to create. If not given, the input filename
                will be used with a .pdf extension
");
            Environment.Exit((int)ExitCode.Success);
        }
    } 
}
