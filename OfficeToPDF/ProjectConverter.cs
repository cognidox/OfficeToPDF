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
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using MSProject = Microsoft.Office.Interop.MSProject;


namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of Project mpp files
    /// </summary>
    class ProjectConverter: Converter, IConverter
    {
        ExitCode IConverter.Convert(String inputFile, String outputFile, ArgParser options, ref List<PDFBookmark> bookmarks)
        {
            if (options.verbose)
            {
                Console.WriteLine("Converting with Project converter");
            }
            return Convert(inputFile, outputFile, options);
        }

        public static ExitCode StartProject(ref Boolean running, ref MSProject.Application project)
        {
            try
            {
                project = (MSProject.Application)Marshal.GetActiveObject("MSProject.Application");
            }
            catch (System.Exception)
            {
                project = new MSProject.Application();
                running = false;
            }
            return ExitCode.Success;
        }

        static ExitCode Convert(String inputFile, String outputFile, ArgParser options)
        {
            Boolean running = options.noquit;
            MSProject.Application app = null;
            object missing = System.Reflection.Missing.Value;
            IWatchdog watchdog = new NullWatchdog();
            try
            {
                ExitCode result = StartProject(ref running, ref app);
                if (result != ExitCode.Success)
                    return result;

                watchdog = WatchdogFactory.CreateStarted(app, options.timeout);

                System.Type type = app.GetType();
                if (type.GetMethod("DocumentExport") == null || System.Convert.ToDouble(app.Version.ToString(), new CultureInfo("en-US")) < 14)
                {
                    Console.WriteLine("Not implemented with Office version {0}", app.Version);
                    return ExitCode.UnsupportedFileFormat;
                }

                app.ShowWelcome = false;
                app.DisplayAlerts = false;
                app.DisplayPlanningWizard = false;
                app.DisplayWizardErrors = false;

                Boolean includeProps = !options.excludeprops;
                Boolean markup = options.markup;
                
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
                            return ExitCode.UnknownError;
                        }
                        app.DocumentExport(outputFile, MSProject.PjDocExportType.pjPDF, includeProps, markup, false, missing, missing);
                        app.FileCloseEx(MSProject.PjSaveType.pjDoNotSave, missing, missing);
                        break;
                }
                return File.Exists(outputFile) ? ExitCode.Success : ExitCode.UnknownError;
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return ExitCode.UnknownError;
            }
            finally
            {
                watchdog.Stop();

                if (app != null && !running)
                {
                    CloseProjectApplication(app);
                }
                ReleaseCOMObject(app);
            }
        }

        internal static void CloseProjectApplication(MSProject.Application app)
        {
            try
            {
                ((MSProject.Application)app).Quit();
            }
            catch (COMException)
            {
                // NOOP - The watchdog may have gone off
            }
        }
    }
}
