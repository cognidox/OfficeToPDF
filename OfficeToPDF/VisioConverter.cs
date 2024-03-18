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
using System.Globalization;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Visio files
    /// </summary>
    class VisioConverter : Converter, IConverter
    {
        ExitCode IConverter.Convert(String inputFile, String outputFile, ArgParser options, ref List<PDFBookmark> bookmarks)
        {
            if (options.verbose)
            {
                Console.WriteLine("Converting with Visio converter");
            }
            return Convert(inputFile, outputFile, options);
        }

        public static ExitCode StartVisio(ref Boolean running, ref InvisibleApp visio)
        {
            try
            {
                visio = (Microsoft.Office.Interop.Visio.InvisibleApp)Marshal.GetActiveObject("Visio.Application");
            }
            catch (System.Exception)
            {
                visio = new Microsoft.Office.Interop.Visio.InvisibleApp();
                running = false;
            }
            return ExitCode.Success;
        }

        static ExitCode Convert(String inputFile, String outputFile, ArgParser options)
        {
            Boolean running = options.noquit;
            Microsoft.Office.Interop.Visio.InvisibleApp visio = null;
            String tmpFile = null;
            String extension = "vsd";
            IWatchdog watchdog = new NullWatchdog();
            try
            {
                ExitCode result = StartVisio(ref running, ref visio);
                if (result != ExitCode.Success)
                    return result;

                watchdog = WatchdogFactory.CreateStarted(visio, options.timeout);

                Regex extReg = new Regex("\\.(\\w+)$");
                Match match = extReg.Match(inputFile);
                if (match.Success)
                {
                    extension = match.Groups[1].Value;
                }

                // We can only convert svg, vsdx and vsdm files with Visio 2013
                if (System.Convert.ToDouble(visio.Version.ToString(), new CultureInfo("en-US")) < 15 &&
                    ((String.Compare(extension, "vsdx", true) == 0) ||
                    (String.Compare(extension, "vsdm", true) == 0) ||
                    (String.Compare(extension, "vdw", true) == 0) ||
                    (String.Compare(extension, "vdx", true) == 0) ||
                    (String.Compare(extension, "svg", true) == 0)))
                {
                    Console.WriteLine("File type not supported in Visio version {0}", visio.Version);
                    return ExitCode.UnsupportedFileFormat;
                }

                bool pdfa = options.pdfa;
                short flags = 0;
                if (options.@readonly)
                {
                    flags += 2;
                }
                if (!options.hidden)
                {
                    visio.Visible = true;
                }
                VisDocExIntent quality = VisDocExIntent.visDocExIntentPrint;
                if (options.print)
                {
                    quality = VisDocExIntent.visDocExIntentPrint;
                }
                if (options.screen)
                {
                    quality = VisDocExIntent.visDocExIntentScreen;
                }
                Boolean includeProps = !options.excludeprops;
                Boolean includeTags = !options.excludetags;

                var documents = visio.Documents;
                documents.OpenEx(inputFile, flags);

                // Try and avoid dialogs about versions and convert non-visio files to
                // visio to get ready for printing
                if ((String.Compare(extension, "svg", true) == 0) ||
                    (String.Compare(extension, "wmf", true) == 0) ||
                    (String.Compare(extension, "emf", true) == 0) ||
                    (String.Compare(extension, "emz", true) == 0) ||
                    (String.Compare(extension, "dwg", true) == 0) ||
                    (String.Compare(extension, "dxf", true) == 0))
                {
                    extension = "vsd";
                }
                var tmpDirectory = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString();
                System.IO.Directory.CreateDirectory(tmpDirectory);
                tmpFile = System.IO.Path.Combine(tmpDirectory, options.original_basename) + "." + extension;

                var activeDoc = visio.ActiveDocument;
                activeDoc.SaveAs(tmpFile);
                activeDoc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, outputFile, quality, VisPrintOutRange.visPrintAll, 1, -1, false, true, includeProps, includeTags, pdfa);
                activeDoc.Close();

                ReleaseCOMObject(documents);
                ReleaseCOMObject(activeDoc);
                return ExitCode.Success;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return ExitCode.UnknownError;
            }
            finally
            {
                watchdog.Stop();

                if (tmpFile != null)
                {
                    System.IO.File.Delete(tmpFile);
                    System.IO.Directory.Delete(System.IO.Path.GetDirectoryName(tmpFile));
                }
                if (visio != null && !running)
                {
                    CloseVisioApplication(visio);
                }
                ReleaseCOMObject(visio);
            }
        }

        internal static void CloseVisioApplication(InvisibleApp visio)
        {
            try
            {
                visio.Quit();
            }
            catch (COMException)
            {
                // NOOP - The watchdog may have gone off
            }
        }
    }
}
