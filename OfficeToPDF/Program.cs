/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010
 *  Copyright (C) 2011  Cognidox Ltd
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
            // Need to error here, as we need input and output files as the
            // arguments to this script
            if (args.Length != 2)
            {
                Console.WriteLine("Please provide the input and output files as arguments");
                Environment.Exit(1);
            }

            // Make sure the input file looks like something we can handle (i.e. has an office
            // filename extension)
            Regex fileMatch = new Regex(@"\.(((ppt|do[ct]|xls)[xm]?)|vsd|pub)$", RegexOptions.IgnoreCase);
            if (fileMatch.Matches(args[0]).Count != 1)
            {
                Console.WriteLine("Input file can not be handled. Must be Word, PowerPoint, Excel, Publisher or Visio");
                Environment.Exit(1);
            }

            String inputFile;
            String outputFile;

            // Make sure the input file exists and is readable
            FileInfo info = new FileInfo(args[0]);
            if (info == null || !info.Exists)
            {
                Console.WriteLine("Input file doesn't exist");
                Environment.Exit(1);
            }
            inputFile = info.FullName;

            // Make sure the destination location exists
            FileInfo outputInfo = new FileInfo(args[1]);
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
                        converted = WordConverter.Convert(inputFile, outputFile);
                        break;
                    case "xls":
                    case "xlsx":
                    case "xlsm":
                        // Excel
                        converted = ExcelConverter.Convert(inputFile, outputFile);
                        break;
                    case "ppt":
                    case "pptx":
                    case "pptm":
                        // Powerpoint
                        converted = PowerpointConverter.Convert(inputFile, outputFile);
                        break;
                    case "vsd":
                        // Visio
                        converted = VisioConverter.Convert(inputFile, outputFile);
                        break;
                    case "pub":
                        // Publisher
                        converted = PublisherConverter.Convert(inputFile, outputFile);
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
    }
}
