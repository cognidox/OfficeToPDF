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
                FileInfo fi = new FileInfo(inputFile);
                // Create a temporary doc file from the message
                tmpDocFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".doc";
                switch(fi.Extension)
                {
                    case ".msg":
                        Microsoft.Office.Interop.Outlook.MailItem message = null;
                        message = (MailItem) app.Session.OpenSharedItem(inputFile);
                        if (message == null)
                        {
                            return false;
                        }
                        message.SaveAs(tmpDocFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);
                        break;
                    case ".vcf":
                        Microsoft.Office.Interop.Outlook.ContactItem contact = null;
                        contact = (ContactItem)app.Session.OpenSharedItem(inputFile);
                        if (contact == null)
                        {
                            return false;
                        }
                        contact.SaveAs(tmpDocFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);
                        break;
                    case ".ics":
                        Microsoft.Office.Interop.Outlook.AppointmentItem appointment = null;
                        appointment = (AppointmentItem)app.Session.OpenSharedItem(inputFile);
                        if (appointment == null)
                        {
                            return false;
                        }
                        appointment.SaveAs(tmpDocFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);
                        break;
                }

                
                if (!File.Exists(tmpDocFile))
                {
                    return false;
                }
                // Convert the doc file to a PDF
                return WordConverter.Convert(tmpDocFile, outputFile, options);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            finally
            {
                if (tmpDocFile != null && File.Exists(tmpDocFile))
                {
                    System.IO.File.Delete(tmpDocFile);
                }
                if (app != null)
                {
                    ((Microsoft.Office.Interop.Outlook._Application)app).Quit();
                    app = null;
                }
            }
        }
    }
}
