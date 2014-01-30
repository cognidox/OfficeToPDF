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
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace OfficeToPDF
{
    /// <summary>
    /// Base converter class that all conversion handlers implement
    /// </summary>
    class Converter
    {
        /// <summary>
        /// Converts an input file to an output PDF
        /// </summary>
        /// <param name="inputFile">Full path of the input file</param>
        /// <param name="outputFile">Full path of the file to output PDF</param>
        /// <returns></returns>
        public static Boolean Convert(String inputFile, String outputFile, Hashtable options)
        {
            return false;
        }
    }
}
