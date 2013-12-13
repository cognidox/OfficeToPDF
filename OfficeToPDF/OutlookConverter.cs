/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010
 *  Copyright (C) 2011-2013 Cognidox Ltd
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
using Microsoft.Office.Interop.Outlook;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Outlook msg files
    /// </summary>
    class OutlookConverter: Converter
    {
        public static new Boolean Convert(String inputFile, String outputFile, Hashtable options)
        {
            Microsoft.Office.Interop.Outlook.Application app = null;
            String tmpDocFile = null;
            try
            {
                app = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem message = null;
                message = (MailItem) app.Session.OpenSharedItem(inputFile);
                if (message == null)
                {
                    return false;
                }
                tmpDocFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".doc";
                message.SaveAs(tmpDocFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);
                if (!File.Exists(tmpDocFile))
                {
                    return false;
                }
                bool converted = false;
                converted = WordConverter.Convert(tmpDocFile, outputFile, options);
                return converted;
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            finally
            {
                if (tmpDocFile != null)
                {
                    System.IO.File.Delete(tmpDocFile);
                }
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
            }
        }
    }
}
