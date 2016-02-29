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
using System.Runtime.InteropServices;
using MSCore = Microsoft.Office.Core;
using MSProject = Microsoft.Office.Interop.MSProject;


namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Project mpp files
    /// </summary>
    class ProjectConverter: Converter
    {
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            Boolean running = (Boolean)options["noquit"];
            MSProject.Application app = null;
            object missing = System.Reflection.Missing.Value;
            try
            {
                try
                {
                    app = (MSProject.Application)Marshal.GetActiveObject("MSProject.Application");
                }
                catch (System.Exception)
                {
                    app = new MSProject.Application();
                    running = false;
                }
                System.Type type = app.GetType();
                if (type.GetMethod("DocumentExport") == null || System.Convert.ToDouble(app.Version.ToString()) < 14)
                {
                    Console.WriteLine("Not implemented with Office version {0}", app.Version);
                    return (int)ExitCode.UnsupportedFileFormat;
                }

                app.ShowWelcome = false;
                app.DisplayAlerts = false;
                app.DisplayPlanningWizard = false;
                app.DisplayWizardErrors = false;

                Boolean includeProps = !(Boolean)options["excludeprops"];
                Boolean markup = (Boolean)options["markup"];
                
                FileInfo fi = new FileInfo(inputFile);
                switch(fi.Extension)
                {
                    case ".mpp":
                        MSProject.Project project = null;
                        if (app.FileOpenEx(inputFile, false, MSProject.PjMergeType.pjDoNotMerge,missing, missing, missing, missing, missing, missing, missing, missing, MSProject.PjPoolOpen.pjDoNotOpenPool, missing, missing, false, missing)) {
                            project = app.ActiveProject;
                        }
                        if (project == null)
                        {
                            return (int)ExitCode.UnknownError;
                        }
                        app.DocumentExport(outputFile, MSProject.PjDocExportType.pjPDF, includeProps, markup, false, missing, missing);
                        app.FileCloseEx(MSProject.PjSaveType.pjDoNotSave, missing, missing);
                        break;
                }
                return File.Exists(outputFile) ? (int)ExitCode.Success : (int)ExitCode.UnknownError;
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return (int)ExitCode.UnknownError;
            }
            finally
            {
                if (app != null && !running)
                {
                    ((MSProject.Application)app).Quit();
                }
                Converter.releaseCOMObject(app);
            }
        }
    }
}
