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
using MSCore = Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Visio files
    /// </summary>
    class VisioConverter: Converter
    {
        public static new Boolean Convert(String inputFile, String outputFile, Hashtable options)
        {
            Microsoft.Office.Interop.Visio.InvisibleApp app = null;
            String tmpFile = null;
            try
            {
                app = new Microsoft.Office.Interop.Visio.InvisibleApp();
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
                VisDocExIntent quality = VisDocExIntent.visDocExIntentScreen;
                if ((Boolean)options["print"])
                {
                    quality = VisDocExIntent.visDocExIntentPrint;
                }

                var documents = app.Documents;
                documents.OpenEx(inputFile, flags);

                // Try and avoid dialogs about versions
                tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".";
                Regex extReg = new Regex("\\.(\\w+)$");
                Match match = extReg.Match(inputFile);
                if (match.Success)
                {
                    tmpFile += match.Groups[1].Value;
                }
                else
                {
                    tmpFile += "vsd";
                }
                var activeDoc = app.ActiveDocument;
                activeDoc.SaveAs(tmpFile);
                activeDoc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, outputFile, quality, VisPrintOutRange.visPrintAll, 1, -1, false, true, true, true, pdfa);
                activeDoc.Close();

                Converter.releaseCOMObject(documents);
                Converter.releaseCOMObject(activeDoc);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            finally
            {
                if (tmpFile != null)
                {
                    System.IO.File.Delete(tmpFile);
                }
                if (app != null)
                {
                    app.Quit();
                }
                Converter.releaseCOMObject(app);
            }
        }
    }
}
