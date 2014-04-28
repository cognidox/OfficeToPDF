/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010
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
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeToPDF
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string[] files = new string[2];
            int filesSeen = 0;
            Hashtable options = new Hashtable();

            // Loop through the input, grabbing switches off the command line
            options["hidden"] = false;
            options["readonly"] = false;
            options["bookmarks"] = false;
            options["print"] = false;
            options["template"] = "";
            Regex switches = new Regex(@"^/(hidden|readonly|bookmarks|print|template|help|\?)$", RegexOptions.IgnoreCase);
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
                            case "template":
                                // Only accept the next option if there are enough options
                                if (argIdx + 3 < args.Length)
                                {
                                    if (File.Exists(args[argIdx + 1]))
                                    {
                                        options[itemMatch.Groups[1].Value.ToLower()] = args[argIdx + 1];
                                    }
                                    else
                                    {
                                        Console.WriteLine("Unable to find {0} {1}", itemMatch.Groups[1].Value.ToLower(), args[argIdx + 1]);
                                    }
                                    argIdx++;

                                }
                                break;
                            default:
                                options[itemMatch.Groups[1].Value.ToLower()] = true;
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Unknown option: {0}", item);
                        Environment.Exit(1);
                    }
                }
                else if (filesSeen < 2)
                {
                    files[filesSeen++] = item;
                }
            }

            // Need to error here, as we need input and output files as the
            // arguments to this script
            if (filesSeen != 2)
            {
                showHelp();
            }

            // Make sure the input file looks like something we can handle (i.e. has an office
            // filename extension)
            Regex fileMatch = new Regex(@"\.(((ppt|pot|do[ct]|xls)[xm]?)|od[cpt]|rtf|csv|vsd|pub|msg|vcf|ics|mpp)$", RegexOptions.IgnoreCase);
            if (fileMatch.Matches(files[0]).Count != 1)
            {
                Console.WriteLine("Input file can not be handled. Must be Word, PowerPoint, Excel, Outlook, Publisher or Visio");
                Environment.Exit(1);
            }

            String inputFile;
            String outputFile;

            // Make sure the input file exists and is readable
            FileInfo info = new FileInfo(files[0]);
            if (info == null || !info.Exists)
            {
                Console.WriteLine("Input file doesn't exist");
                Environment.Exit(1);
            }
            inputFile = info.FullName;

            // Make sure the destination location exists
            FileInfo outputInfo = new FileInfo(files[1]);
            if (outputInfo != null && outputInfo.Exists)
            {
                System.IO.File.Delete(outputInfo.FullName);
            }
            if (!System.IO.Directory.Exists(outputInfo.DirectoryName))
            {
                Console.WriteLine("Output directory does not exist");
                Environment.Exit(1);
            }
            outputFile = outputInfo.FullName;

            // Now, do the cleverness of determining what the extension is, and so, which
            // conversion class to pass it to
            Boolean converted = false;
            Match extMatch = fileMatch.Match(inputFile);
            if (extMatch.Success)
            {
                switch (extMatch.Groups[1].ToString())
                {
                    case "rtf":
                    case "odt":
                    case "doc":
                    case "dot":
                    case "docx":
                    case "dotx":
                    case "docm":
                    case "dotm":
                        // Word
                        converted = WordConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "csv":
                    case "odc":
                    case "xls":
                    case "xlsx":
                    case "xlsm":
                        // Excel
                        converted = ExcelConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "odp":
                    case "ppt":
                    case "pptx":
                    case "pptm":
                    case "pot":
                    case "potm":
                    case "potx":
                        // Powerpoint
                        converted = PowerpointConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "vsd":
                        // Visio
                        converted = VisioConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "pub":
                        // Publisher
                        converted = PublisherConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "msg":
                    case "vcf":
                    case "ics":
                        // Outlook
                        converted = OutlookConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "mpp":
                        // Project
                        converted = ProjectConverter.Convert(inputFile, outputFile, options);
                        break;
                }
            }
            if (!converted)
            {
                Console.WriteLine("Did not convert");
                Environment.Exit(1);
            }
            Environment.Exit(0);
        }

        static void showHelp()
        {
            Console.Write(@"Converts Office documents to PDF from the command line.
Handles Office files:
  doc, dot, docx, dotx, docm, dotm, rtf, odt, ppt, pptx, pptm, odp,
  xls, xlsx, xlsm, csv, odc, vsd, pub, mpp, ics, vcf, msg

OfficeToPDF.exe [/bookmarks] [/hidden] [/readonly] input_file output_file

  /bookmarks  Create bookmarks in the PDF when they are supported by the
              Office application
  /hidden     Attempt to hide the Office application window when converting
  /readonly   Load the input file in read only mode where possible
  /print      Create high-quality PDFs optimised for print
  /template <template_path> use a .dot, .dotx or .dotm template when converting with Word

  input_file  The filename of the Office document to convert
  output_file The filename of the PDF to create
");
            Environment.Exit(0);
        }
    } 
}
