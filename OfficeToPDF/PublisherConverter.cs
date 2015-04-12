﻿/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010
 *  Copyright (C) 2011-2015 Cognidox Ltd
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
using System.Runtime.InteropServices;
using MSCore = Microsoft.Office.Core;
using Microsoft.Office.Interop.Publisher;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Microsoft Publisher files
    /// </summary>
    class PublisherConverter: Converter
    {
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            Boolean running = (Boolean)options["noquit"];
            Microsoft.Office.Interop.Publisher.Application app = null;
            String tmpFile = null;
            try
            {
                try
                {
                    app = (Microsoft.Office.Interop.Publisher.Application)Marshal.GetActiveObject("Publisher.Application");
                }
                catch (System.Exception)
                {
                    app = new Microsoft.Office.Interop.Publisher.Application();
                    running = false;
                }
                Boolean nowrite = (Boolean)options["readonly"];
                bool pdfa = (Boolean)options["pdfa"] ? true : false;
                if ((Boolean)options["hidden"])
                {
                    var activeWin = app.ActiveWindow;
                    activeWin.Visible = false;
                    Converter.releaseCOMObject(activeWin);
                }
                app.Open(inputFile, nowrite, false, PbSaveOptions.pbDoNotSaveChanges);
                PbFixedFormatIntent quality = PbFixedFormatIntent.pbIntentStandard;
                if ((Boolean)options["print"])
                {
                    quality = PbFixedFormatIntent.pbIntentPrinting;
                }
                if ((Boolean)options["screen"])
                {
                    quality = PbFixedFormatIntent.pbIntentMinimum;
                }
                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean includeTags = !(Boolean)options["excludetags"];

                // Try and avoid dialogs about versions
                tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".pub";
                var activeDocument = app.ActiveDocument;
                activeDocument.SaveAs(tmpFile, PbFileFormat.pbFilePublication, false);
                activeDocument.ExportAsFixedFormat(PbFixedFormatType.pbFixedFormatTypePDF, outputFile, quality, includeProps, -1, -1, -1, -1, -1, -1, -1, true, PbPrintStyle.pbPrintStyleDefault, includeTags, true, pdfa);
                activeDocument.Close();

                Converter.releaseCOMObject(activeDocument);
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
                if (app != null)
                {
                    ((Microsoft.Office.Interop.Publisher._Application)app).Quit();
                }
                Converter.releaseCOMObject(app);
            }
        }
    }
}
