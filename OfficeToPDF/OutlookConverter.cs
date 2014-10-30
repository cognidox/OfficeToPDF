/**
 *  OfficeToPDF command line PDF conversion for Office 2007/2010/2013
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
using System.Runtime.InteropServices;
using MSCore = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Outlook msg files
    /// </summary>
    class OutlookConverter: Converter
    {
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            Boolean running = (Boolean)options["noquit"];
            Microsoft.Office.Interop.Outlook.Application app = null;
            String tmpDocFile = null;
            try
            {
                try
                {
                    app = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch(System.Exception)
                {
                    app = new Microsoft.Office.Interop.Outlook.Application();
                    running = false;
                }
                var session = app.Session;
                FileInfo fi = new FileInfo(inputFile);
                // Create a temporary doc file from the message
                tmpDocFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".doc";
                switch(fi.Extension)
                {
                    case ".msg":
                        var message = (MailItem) session.OpenSharedItem(inputFile);
                        if (message == null)
                        {
                            return (int)ExitCode.FileOpenFailure;
                        }
                        message.SaveAs(tmpDocFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);
                        Converter.releaseCOMObject(message);
                        Converter.releaseCOMObject(session);
                        break;
                    case ".vcf":
                        var contact = (ContactItem)session.OpenSharedItem(inputFile);
                        if (contact == null)
                        {
                            return (int)ExitCode.FileOpenFailure;
                        }
                        contact.SaveAs(tmpDocFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);
                        Converter.releaseCOMObject(contact);
                        Converter.releaseCOMObject(session);
                        break;
                    case ".ics":
                        var item = session.OpenSharedItem(inputFile);
                        string itemType = (string)(string)item.GetType().InvokeMember("MessageClass", System.Reflection.BindingFlags.GetProperty, null, item, null);
                        switch (itemType)
                        {
                            case "IPM.Appointment":
                                var appointment = (AppointmentItem)item;
                                if (appointment != null)
                                {
                                    appointment.SaveAs(tmpDocFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);
                                }
                                break;
                            default:
                                Console.WriteLine("Unable to convert ICS type " + itemType);
                                break;
                        }
                        Converter.releaseCOMObject(item);
                        Converter.releaseCOMObject(session);
                        break;
                }

                if (!File.Exists(tmpDocFile))
                {
                    return (int)ExitCode.UnknownError;
                }
                // Convert the doc file to a PDF
                return WordConverter.Convert(tmpDocFile, outputFile, options);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            finally
            {
                if (tmpDocFile != null && File.Exists(tmpDocFile))
                {
                    System.IO.File.Delete(tmpDocFile);
                }
                // If we were not already running, quit and release the outlook object
                if (app != null && !running)
                {
                    ((Microsoft.Office.Interop.Outlook._Application)app).Quit();
                }
                Converter.releaseCOMObject(app);
            }
        }
    }
}
