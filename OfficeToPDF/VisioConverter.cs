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
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using MSCore = Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Visio files
    /// </summary>
    class VisioConverter: Converter
    {
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            Boolean running = (Boolean)options["noquit"];
            Microsoft.Office.Interop.Visio.InvisibleApp app = null;
            String tmpFile = null;
            String extension = "vsd";
            try
            {
                try
                {
                    app = (Microsoft.Office.Interop.Visio.InvisibleApp)Marshal.GetActiveObject("Visio.Application");
                }
                catch (System.Exception)
                {
                    app = new Microsoft.Office.Interop.Visio.InvisibleApp();
                    running = false;
                }
                Regex extReg = new Regex("\\.(\\w+)$");
                Match match = extReg.Match(inputFile);
                if (match.Success)
                {
                    extension = match.Groups[1].Value;
                }

                // We can only convert svg, vsdx and vsdm files with Visio 2013
                if (System.Convert.ToDouble(app.Version.ToString()) < 15 &&
                    ((String.Compare(extension, "vsdx") == 0) ||
                    (String.Compare(extension, "vsdm") == 0) ||
                    (String.Compare(extension, "svg") == 0)))
                {
                    Console.WriteLine("File type not supported in Visio version {0}", app.Version);
                    return (int)ExitCode.UnsupportedFileFormat;
                }

                bool pdfa = (Boolean)options["pdfa"] ? true : false;
                short flags = 0;
                if ((Boolean)options["readonly"])
                {
                    flags += 2;
                }
                if (!(Boolean)options["hidden"])
                {
                    app.Visible = true;
                }
                VisDocExIntent quality = VisDocExIntent.visDocExIntentPrint;
                if ((Boolean)options["print"])
                {
                    quality = VisDocExIntent.visDocExIntentPrint;
                }
                if ((Boolean)options["screen"])
                {
                    quality = VisDocExIntent.visDocExIntentScreen;
                }
                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean includeTags = !(Boolean)options["excludetags"];

                var documents = app.Documents;
                documents.OpenEx(inputFile, flags);

                // Try and avoid dialogs about versions and convert SVG files to
                // visio to get ready for printing
                if (String.Compare(extension, "svg") == 0)
                {
                    extension = "vsd";
                }
                tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + "." + extension;
                
                var activeDoc = app.ActiveDocument;
                activeDoc.SaveAs(tmpFile);
                activeDoc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, outputFile, quality, VisPrintOutRange.visPrintAll, 1, -1, false, true, includeProps, includeTags, pdfa);
                activeDoc.Close();

                Converter.releaseCOMObject(documents);
                Converter.releaseCOMObject(activeDoc);
                return (int)ExitCode.Success;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            finally
            {
                if (tmpFile != null)
                {
                    System.IO.File.Delete(tmpFile);
                }
                if (app != null && !running)
                {
                    app.Quit();
                }
                Converter.releaseCOMObject(app);
            }
        }
    }
}
