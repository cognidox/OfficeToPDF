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
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;
using System.Text.RegularExpressions;
using PdfSharp.Pdf;

namespace OfficeToPDF
{
    internal class ArgParser : Hashtable
    {
        public Action<ExitCode> Exit { get; private set; }
        public Action<string, object[]> Output { get; private set; }

        // Strings used in error messages for different options
        private readonly Dictionary<string, string> optionNameMap = new Dictionary<string, string>()
            {
                { "excel_max_rows", "Maximum number of rows" },
                { "excel_worksheet", "Excel worksheet" },
                { "word_max_pages", "Maximum number of pages" },
                { "excel_delay", "Excel delay milliseconds" },
                { "timeout", "Timeout in seconds to wait for generation of the pdf" }
            };

        public ArgParser()
            : this(value => Environment.Exit((int)value), (format, args) => Console.WriteLine(format, args))
        { }

        private void WriteLine(string msg) => Output(msg, Array.Empty<object>());

        private void WriteLine(string format, object arg0) => Output(format, new[] { arg0 });

        private void WriteLine(string format, object arg0, object arg1) => Output(format, new[] { arg0, arg1 });

        public ArgParser(Action<ExitCode> exit, Action<string, object[]> output)
        {
            Exit = exit;
            Output = output;

            this["hidden"] = false;
            this["markup"] = false;
            this["readonly"] = false;
            this["bookmarks"] = false;
            this["print"] = true;
            this["screen"] = false;
            this["pdfa"] = false;
            this["verbose"] = false;
            this["excludeprops"] = false;
            this["excludetags"] = false;
            this["noquit"] = false;
            this["merge"] = false;
            this["template"] = "";
            this["password"] = "";
            this["printer"] = "";
            this["fallback_printer"] = "";
            this["working_dir"] = "";
            this["has_working_dir"] = false;
            this["excel_show_formulas"] = false;
            this["excel_show_headings"] = false;
            this["excel_auto_macros"] = false;
            this["excel_template_macros"] = false;
            this["excel_active_sheet"] = false;
            this["excel_no_link_update"] = false;
            this["excel_no_recalculate"] = false;
            this["excel_no_map_papersize"] = false;
            this["excel_max_rows"] = 0;
            this["excel_active_sheet_on_max_rows"] = false;
            this["excel_worksheet"] = 0;
            this["excel_delay"] = 0;
            this["word_field_quick_update"] = false;
            this["word_field_quick_update_safe"] = false;
            this["word_no_field_update"] = false;
            this["word_header_dist"] = (float)-1;
            this["word_footer_dist"] = (float)-1;
            this["word_max_pages"] = 0;
            this["word_ref_fonts"] = false;
            this["word_keep_history"] = false;
            this["word_no_repair"] = false;
            this["word_show_comments"] = false;
            this["word_show_revs_comments"] = false;
            this["word_show_format_changes"] = false;
            this["word_show_hidden"] = false;
            this["word_show_ink_annot"] = false;
            this["word_show_ins_del"] = false;
            this["word_markup_balloon"] = false;
            this["word_show_all_markup"] = false;
            this["word_fix_table_columns"] = false;
            this["word_no_map_papersize"] = false;
            this["original_filename"] = "";
            this["original_basename"] = "";
            this["powerpoint_output"] = "";
            this["pdf_page_mode"] = null;
            this["pdf_layout"] = null;
            this["pdf_merge"] = MergeMode.None;
            this["pdf_clean_meta"] = MetaClean.None;
            this["pdf_owner_pass"] = "";
            this["pdf_user_pass"] = "";
            this["pdf_restrict_annotation"] = false;
            this["pdf_restrict_extraction"] = false;
            this["pdf_restrict_assembly"] = false;
            this["pdf_restrict_forms"] = false;
            this["pdf_restrict_modify"] = false;
            this["pdf_restrict_print"] = false;
            this["pdf_restrict_accessibility_extraction"] = false;
            this["pdf_restrict_full_quality"] = false;
            this["timeout"] = 0;
        }

        public string[] files = new string[2];
        public int filesSeen = 0;
        public Boolean postProcessPDF = false;
        public Boolean postProcessPDFSecurity = false;

        public int timeout => TryGetKeyValue<int>();

        private T TryGetKeyValue<T>([CallerMemberName] string key = null) =>
            this.ContainsKey(Normalise(key)) ? (T)this[Normalise(key)] : default(T);

        private void SetKeyValue<T>(T value, [CallerMemberName] string key = null) =>
            this[Normalise(key)] = value;

        private static string Normalise(string key) => key.ToLowerInvariant();

        public bool hidden => TryGetKeyValue<bool>();
        public bool markup => TryGetKeyValue<bool>();
        public bool @readonly => TryGetKeyValue<bool>();
        public bool bookmarks => TryGetKeyValue<bool>();
        public bool print
        {
            get => TryGetKeyValue<bool>();
            private set => SetKeyValue(value);
        }
        public bool screen
        {
            get => TryGetKeyValue<bool>();
            private set => SetKeyValue(value);
        }
        public bool pdfa => TryGetKeyValue<bool>();
        public bool verbose => TryGetKeyValue<bool>();
        public bool excludeprops => TryGetKeyValue<bool>();
        public bool excludetags => TryGetKeyValue<bool>();
        public bool noquit => TryGetKeyValue<bool>();
        public bool merge => TryGetKeyValue<bool>();
        public string template => TryGetKeyValue<string>();
        public string password => TryGetKeyValue<string>();
        public string printer => TryGetKeyValue<string>();
        public string fallback_printer => TryGetKeyValue<string>();
        public string working_dir
        {
            get => TryGetKeyValue<string>();
            private set => SetKeyValue(value);
        }
        public bool has_working_dir
        {
            get => TryGetKeyValue<bool>();
            private set => SetKeyValue(value);
        }

        public bool word_keep_history => TryGetKeyValue<bool>();
        public bool word_no_repair => TryGetKeyValue<bool>();
        public string writepassword => TryGetKeyValue<string>();
        public int word_max_pages => TryGetKeyValue<int>();
        public bool word_ref_fonts => TryGetKeyValue<bool>();
        public bool IsTempWord
        {
            get => TryGetKeyValue<bool>();
            set => SetKeyValue(value);
        }
        public bool word_show_hidden => TryGetKeyValue<bool>();
        public bool word_no_map_papersize => TryGetKeyValue<bool>();
        public bool word_show_all_markup => TryGetKeyValue<bool>();
        public bool word_show_comments
        {
            get => TryGetKeyValue<bool>();
            set => SetKeyValue(value);
        }
        public bool word_show_revs_comments
        {
            get => TryGetKeyValue<bool>();
            set => SetKeyValue(value);
        }
        public bool word_show_format_changes
        {
            get => TryGetKeyValue<bool>();
            set => SetKeyValue(value);
        }
        public bool word_show_ink_annot
        {
            get => TryGetKeyValue<bool>();
            set => SetKeyValue(value);
        }
        public bool word_show_ins_del
        {
            get => TryGetKeyValue<bool>();
            set => SetKeyValue(value);
        }
        public bool word_markup_balloon => TryGetKeyValue<bool>();
        public bool word_fix_table_columns => TryGetKeyValue<bool>();
        public bool word_no_field_update => TryGetKeyValue<bool>();
        public float word_header_dist => TryGetKeyValue<float>();
        public float word_footer_dist => TryGetKeyValue<float>();
        public bool word_field_quick_update => TryGetKeyValue<bool>();
        public bool word_field_quick_update_safe => TryGetKeyValue<bool>();

        public string working_orig_dest
        {
            get => TryGetKeyValue<string>();
            set => SetKeyValue(value);
        }

        public bool excel_no_map_papersize => TryGetKeyValue<bool>();
        public bool excel_active_sheet => TryGetKeyValue<bool>();
        public bool excel_active_sheet_on_max_rows => TryGetKeyValue<bool>();
        public bool excel_no_recalculate => TryGetKeyValue<bool>();
        public bool excel_show_headings => TryGetKeyValue<bool>();
        public bool excel_show_formulas => TryGetKeyValue<bool>();
        public bool excel_no_link_update => TryGetKeyValue<bool>();
        public bool excel_auto_macros => TryGetKeyValue<bool>();
        public int excel_max_rows => TryGetKeyValue<int>();
        public int excel_worksheet => TryGetKeyValue<int>();
        public int excel_delay => TryGetKeyValue<int>();

        public string powerpoint_output
        {
            get => TryGetKeyValue<string>();
            private set => SetKeyValue(value);
        }

        public PdfPageMode? pdf_page_mode
        {
            get => TryGetKeyValue<PdfPageMode?>();
            private set => SetKeyValue(value);
        }
        public PdfPageLayout? pdf_layout
        {
            get => TryGetKeyValue<PdfPageLayout?>();
            private set => SetKeyValue(value);
        }
        public MergeMode pdf_merge
        {
            get => TryGetKeyValue<MergeMode>();
            set => SetKeyValue(value);
        }
        public MetaClean pdf_clean_meta
        {
            get => TryGetKeyValue<MetaClean>();
            private set => SetKeyValue(value);
        }
        public string pdf_owner_pass => TryGetKeyValue<string>();
        public string pdf_user_pass => TryGetKeyValue<string>();
        public bool pdf_restrict_annotation => TryGetKeyValue<bool>();
        public bool pdf_restrict_extraction => TryGetKeyValue<bool>();
        public bool pdf_restrict_assembly => TryGetKeyValue<bool>();
        public bool pdf_restrict_forms => TryGetKeyValue<bool>();
        public bool pdf_restrict_modify => TryGetKeyValue<bool>();
        public bool pdf_restrict_print => TryGetKeyValue<bool>();
        public bool pdf_restrict_accessibility_extraction => TryGetKeyValue<bool>();
        public bool pdf_restrict_full_quality => TryGetKeyValue<bool>();

        public string original_filename
        {
            get => TryGetKeyValue<string>();
            set => SetKeyValue(value);
        }
        public string original_basename
        {
            get => TryGetKeyValue<string>();
            set => SetKeyValue(value);
        }

        private ExitCode CheckOptionIsInteger(string optionKey, string optionName, string optionValue)
        {
            if (Regex.IsMatch(optionValue, @"^\d+$"))
            {
                this[optionKey] = Convert.ToInt32(optionValue);
                return ExitCode.Success;
            }
            WriteLine("{0} ({1}) is invalid", optionName, optionValue);
            return ExitCode.Failed | ExitCode.InvalidArguments;
        }


        public ExitCode Parse(string[] args, Dictionary<string, bool> installedPrinters)
        {
            // Loop through the input, grabbing switches off the command line

            Regex switches = new Regex(@"^/(version|hidden|markup|readonly|bookmarks|merge|noquit|print|(fallback_)?printer|screen|pdfa|template|writepassword|password|help|verbose|exclude(props|tags)|excel_(delay|max_rows|show_formulas|show_headings|auto_macros|template_macros|active_sheet|active_sheet_on_max_rows|worksheet|no_recalculate|no_link_update|no_map_papersize)|powerpoint_(output)|word_(show_hidden|header_dist|footer_dist|ref_fonts|no_field_update|field_quick_update(_safe)?|max_pages|keep_history|no_repair|fix_table_columns|show_(comments|revs_comments|format_changes|ink_annot|ins_del|all_markup)|markup_balloon|no_map_papersize)|pdf_(page_mode|append|prepend|layout|clean_meta|owner_pass|user_pass|restrict_(annotation|extraction|assembly|forms|modify|print|accessibility_extraction|full_quality))|working_dir|timeout|\?)$", RegexOptions.IgnoreCase);
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
                            ShowHelpAndExit();
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
                                            this.pdf_page_mode = PdfPageMode.FullScreen;
                                            break;
                                        case "none":
                                            this.pdf_page_mode = PdfPageMode.UseNone;
                                            break;
                                        case "bookmarks":
                                            this.pdf_page_mode = PdfPageMode.UseOutlines;
                                            break;
                                        case "thumbs":
                                            this.pdf_page_mode = PdfPageMode.UseThumbs;
                                            break;
                                        default:
                                            WriteLine("Invalid PDF page mode ({0}). It must be one of full, none, outline or thumbs", args[argIdx + 1]);
                                            return ExitCode.Failed | ExitCode.InvalidArguments;
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
                                            this.pdf_clean_meta = MetaClean.Basic;
                                            break;
                                        case "full":
                                            this.pdf_clean_meta = MetaClean.Full;
                                            break;
                                        default:
                                            WriteLine("Invalid PDF meta-data clean value ({0}). It must be one of full or basic", args[argIdx + 1]);
                                            return ExitCode.Failed | ExitCode.InvalidArguments;
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
                                            this.pdf_layout = PdfPageLayout.OneColumn;
                                            break;
                                        case "single":
                                            this.pdf_layout = PdfPageLayout.SinglePage;
                                            break;
                                        case "twocolleft":
                                            this.pdf_layout = PdfPageLayout.TwoColumnLeft;
                                            break;
                                        case "twocolright":
                                            this.pdf_layout = PdfPageLayout.TwoColumnRight;
                                            break;
                                        case "twopageleft":
                                            this.pdf_layout = PdfPageLayout.TwoPageLeft;
                                            break;
                                        case "twopageright":
                                            this.pdf_layout = PdfPageLayout.TwoPageRight;
                                            break;
                                        default:
                                            WriteLine("Invalid PDF layout ({0}). It must be one of onecol, single, twocolleft, twocolright, twopageleft or twopageright", args[argIdx + 1]);
                                            return ExitCode.Failed | ExitCode.InvalidArguments;
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
                                    this[itemMatch.Groups[1].Value.ToLower()] = pass;
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
                                        this[itemMatch.Groups[1].Value.ToLower()] = templateInfo.FullName;
                                    }
                                    else
                                    {
                                        WriteLine("Unable to find {0} {1}", itemMatch.Groups[1].Value.ToLower(), args[argIdx + 1]);
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
                                        WriteLine("Invalid PowerPoint output type");
                                        return ExitCode.Failed | ExitCode.InvalidArguments;
                                    }
                                    this.powerpoint_output = args[argIdx + 1];
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
                                        catch
                                        { /* NOOP */ }
                                        if (workingDirectoryWritable)
                                        {
                                            int maxTries = 20;
                                            while (maxTries-- > 0)
                                            {
                                                this.working_dir = Path.Combine(args[argIdx + 1], Guid.NewGuid().ToString());
                                                if (!Directory.Exists(this.working_dir))
                                                {
                                                    DirectoryInfo di = Directory.CreateDirectory(this.working_dir);
                                                    if (di.Exists)
                                                    {
                                                        this.has_working_dir = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (maxTries <= 0)
                                            {
                                                WriteLine("A working directory can not be created");
                                                return ExitCode.Failed | ExitCode.InvalidArguments;
                                            }
                                        }
                                        else
                                        {
                                            // The working directory must be writable
                                            WriteLine("The working directory {0} is not writable", args[argIdx + 1]);
                                            return ExitCode.Failed | ExitCode.InvalidArguments;
                                        }
                                    }
                                    else
                                    {
                                        // We need a real directory to work in, so there is an error here
                                        WriteLine("Unable to find working directory {0}", args[argIdx + 1]);
                                        return ExitCode.Failed | ExitCode.InvalidArguments;
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
                                    ExitCode result = CheckOptionIsInteger(itemMatch.Groups[1].Value.ToLower(), optionNameMap[itemMatch.Groups[1].Value.ToLower()], args[argIdx + 1]);
                                    if (result != ExitCode.Success)
                                        return result;

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
                                            this[itemMatch.Groups[1].Value.ToLower()] = (float)Convert.ToDouble(args[argIdx + 1]);
                                        }
                                        catch
                                        {
                                            WriteLine("Header/Footer distance ({0}) is invalid", args[argIdx + 1]);
                                            return ExitCode.Failed | ExitCode.InvalidArguments;
                                        }
                                    }
                                    else
                                    {
                                        WriteLine("Header/Footer distance ({0}) is invalid", args[argIdx + 1]);
                                        return ExitCode.Failed | ExitCode.InvalidArguments;
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
                                    this[optname] = args[argIdx + 1];
                                    argIdx++;
                                }
                                if (optname.Equals("printer") || optname.Equals("fallback_printer"))
                                {
                                    if (!installedPrinters.ContainsKey(((string)this[optname]).ToLowerInvariant()))
                                    {
                                        // The requested printer did not exists
                                        WriteLine("The printer \"{0}\" is not installed", this[optname]);
                                        return ExitCode.Failed | ExitCode.InvalidArguments;
                                    }
                                }
                                break;
                            case "screen":
                                this.print = false;
                                this.screen = true;
                                break;
                            case "print":
                                this.screen = false;
                                this.print = true;
                                break;
                            case "version":
                                Assembly asm = Assembly.GetExecutingAssembly();
                                FileVersionInfo fv = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                                WriteLine("{0}", fv.FileVersion);
                                Exit(ExitCode.Success);
                                break;
                            case "pdf_append":
                                if (this.pdf_merge != MergeMode.None)
                                {
                                    WriteLine("Only one of /pdf_append or /pdf_prepend can be used");
                                    return ExitCode.Failed | ExitCode.InvalidArguments;
                                }
                                postProcessPDF = true;
                                this.pdf_merge = MergeMode.Append;
                                break;
                            case "pdf_prepend":
                                if (this.pdf_merge != MergeMode.None)
                                {
                                    WriteLine("Only one of /pdf_append or /pdf_prepend can be used");
                                    return ExitCode.Failed | ExitCode.InvalidArguments;
                                }
                                postProcessPDF = true;
                                this.pdf_merge = MergeMode.Prepend;
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
                                this[itemMatch.Groups[1].Value.ToLower()] = true;
                                break;
                            case "timeout":
                                // Only accept the next option if there are enough options
                                if (argIdx + 1 < args.Length)
                                {
                                    ExitCode result = CheckOptionIsInteger(itemMatch.Groups[1].Value.ToLower(), optionNameMap[itemMatch.Groups[1].Value.ToLower()], args[argIdx + 1]);
                                    if (result != ExitCode.Success)
                                        return result;

                                    argIdx++;
                                }
                                break;
                            default:
                                this[itemMatch.Groups[1].Value.ToLower()] = true;
                                break;
                        }
                    }
                    else
                    {
                        WriteLine("Unknown option: {0}", item);
                        return ExitCode.Failed | ExitCode.InvalidArguments;
                    }
                }
                else if (filesSeen < 2)
                {
                    files[filesSeen++] = item;
                }
            }

            return ExitCode.Success;
        }

        public void ShowHelpAndExit()
        {
            WriteLine(@"Converts Office documents to PDF from the command line.
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
  /excel_active_sheet_on_max_rows - Only convert the active sheet if another worksheet has too many rows.
  /excel_auto_macros        - Run Auto_Open macros in Excel files before conversion
  /excel_show_formulas      - Show formulas in the generated PDF
  /excel_delay <ms>         - Number of milliseconds to pause Excel for during file processing
  /excel_show_headings      - Show row and column headings
  /excel_max_rows <rows>    - If any worksheet in a spreadsheet document has more
                              than this number of rows, do not attempt to convert
                              the file. Applies when converting with Excel.
  /excel_no_link_update     - Do not update links when opening Excel files.
  /excel_no_recalculate     - Skip automatic re-calculation of formulas in the workbook.
  /excel_no_map_papersize   - Do not map papersize to the local paper size.
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
  /word_no_map_papersize    - Do not map papersize to the local paper size.
  /word_ref_fonts           - When fonts are not available, a reference to the font is used in
                              the generated PDF rather than a bitmapped version. The default is
                              for a bitmap of the text to be used.
  /word_show_comments       - Show comments when /markup is used.
  /word_show_revs_comments  - Show revisions and comments when /markup is used.
  /word_show_format_changes - Show format changes when /markup is used.
  /word_show_hidden         - Show hidden text.
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
            Exit(ExitCode.Success);
        }
    }
}
