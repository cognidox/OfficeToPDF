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
using Microsoft.Office.Interop.Word;

namespace OfficeToPDF
{
    // We want to be able to reset the options in Word so it doesn't affect subsequent
    // usage
    internal class AppOption
    {
        public string Name { get; set; }
        public Boolean Value { get; set; }
        public Boolean OriginalValue { get; set; }
        public int IntValue { get; set; }
        public int OriginalIntValue { get; set; }
        protected Type VarType { get; set; }
        public AppOption(string name, Boolean value, ref Options wdOptions)
        {
            try
            {
                Name = name;
                Value = value;
                VarType = typeof(Boolean);
                OriginalValue = (Boolean)wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, wdOptions, null);

                if (OriginalValue != value)
                {
                    wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { value });
                }
            }
            catch
            {
                // We may be setting word options that are not available in the version of word
                // being used, so just skip these errors
            }
        }
        public AppOption(string name, int value, ref Options wdOptions)
        {
            try
            {
                Name = name;
                IntValue = value;
                VarType = typeof(int);
                OriginalIntValue = (int)wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, wdOptions, null);

                if (OriginalIntValue != value)
                {
                    wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { value });
                }
            }
            catch
            {
                // We may be setting word options that are not available in the version of word
                // being used, so just skip these errors
            }
        }

        // Allow the value on the options to be reset
        public void ResetValue(ref Options wdOptions)
        {
            if (VarType == typeof(Boolean))
            {
                if (Value != this.OriginalValue)
                {
                    wdOptions.GetType().InvokeMember(Name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { OriginalValue });
                }
            }
            else
            {
                if (IntValue != OriginalIntValue)
                {
                    wdOptions.GetType().InvokeMember(Name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { OriginalIntValue });
                }
            }
        }
    }
}
