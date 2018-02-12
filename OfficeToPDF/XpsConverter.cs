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
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using PdfSharp.Xps;

namespace OfficeToPDF
{
    /// <summary>
    /// Handle conversion of XPS files
    /// </summary>
    class XpsConverter: Converter
    {
        public static new int Convert(String inputFile, String outputFile, Hashtable options)
        {
            try
            {

                PdfSharp.Xps.XpsConverter.Convert(inputFile, outputFile, 0);
                if (System.IO.File.Exists(outputFile))
                {
                    return (int)ExitCode.Success;
                }
            }
            catch (Exception) { }
            return (int)ExitCode.UnknownError;
        }
    }
}
