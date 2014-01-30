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
            Regex switches = new Regex(@"^/(hidden|readonly|bookmarks|print|help|\?)$", RegexOptions.IgnoreCase);
            foreach (string item in args)
            {
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
                        options[itemMatch.Groups[1].Value.ToLower()] = true;
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
            Regex fileMatch = new Regex(@"\.(((ppt|do[ct]|xls)[xm]?)|vsd|pub|msg|vcf|ics)$", RegexOptions.IgnoreCase);
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
                    case "doc":
                    case "dot":
                    case "docx":
                    case "dotx":
                    case "docm":
                    case "dotm":
                        // Word
                        converted = WordConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "xls":
                    case "xlsx":
                    case "xlsm":
                        // Excel
                        converted = ExcelConverter.Convert(inputFile, outputFile, options);
                        break;
                    case "ppt":
                    case "pptx":
                    case "pptm":
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
  doc, dot, docx, dotx, docm, dotm, ppt, pptx, pptm, xls, xlsx, xlsm, vsd, pub

OfficeToPDF.exe [/bookmarks] [/hidden] [/readonly] input_file output_file

  /bookmarks  Create bookmarks in the PDF when they are supported by the
              Office application
  /hidden     Attempt to hide the Office application window when converting
  /readonly   Load the input file in read only mode where possible
  /print      Create high-quality PDFs optimised for print

  input_file  The filename of the Office document to convert
  output_file The filename of the PDF to create
");
            Environment.Exit(0);
        }
    } 
}
