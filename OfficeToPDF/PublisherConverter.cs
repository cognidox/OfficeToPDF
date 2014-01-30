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
using MSCore = Microsoft.Office.Core;
using Microsoft.Office.Interop.Publisher;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Microsoft Publisher files
    /// </summary>
    class PublisherConverter: Converter
    {
        public static new Boolean Convert(String inputFile, String outputFile, Hashtable options)
        {
            Microsoft.Office.Interop.Publisher.Application app = null;
            String tmpFile = null;
            try
            {
                Boolean nowrite = (Boolean)options["readonly"];
                app = new Microsoft.Office.Interop.Publisher.Application();
                if ((Boolean)options["hidden"])
                {
                    app.ActiveWindow.Visible = false;
                }
                app.Open(inputFile, nowrite, false, PbSaveOptions.pbDoNotSaveChanges);
                PbFixedFormatIntent quality = PbFixedFormatIntent.pbIntentStandard;
                if ((Boolean)options["print"])
                {
                    quality = PbFixedFormatIntent.pbIntentPrinting;
                }

                // Try and avoid dialogs about versions
                tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".pub";
                app.ActiveDocument.SaveAs(tmpFile, PbFileFormat.pbFilePublication, false);
                app.ActiveDocument.ExportAsFixedFormat(PbFixedFormatType.pbFixedFormatTypePDF, outputFile, quality, true);
                app.ActiveDocument.Close();
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
                    ((Microsoft.Office.Interop.Publisher._Application)app).Quit();
                    app = null;
                }
            }
        }
    }
}
