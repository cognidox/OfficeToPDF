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
using Microsoft.Office.Interop.Word;

namespace OfficeToPDF
{
    internal interface IAppOption
    {
        void ResetValue(ref Options wdOptions);
    }

    // We want to be able to reset the options in Word so it doesn't affect subsequent
    // usage
    internal class AppOption<T> : IAppOption
    {
        public string Name { get; }
        public T Value { get; }
        public T OriginalValue { get; }

        public AppOption(string name, T value, ref Options wdOptions)
        {
            Name = name;
            Value = value;

            try
            {
                OriginalValue = (T)wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, wdOptions, null);

                if (Equals(Value, OriginalValue))
                    return;

                SetProperty(ref wdOptions, Name, Value);
            }
            catch
            {
                // We may be setting word options that are not available in the version of word
                // being used, so just skip these errors

                OriginalValue = Value; // Don't try and restore setting on reset
            }
        }

        private static bool Equals(T lhs, T rhs) => EqualityComparer<T>.Default.Equals(lhs, rhs);

        private static void SetProperty(ref Options wdOptions, string name, object value) =>
            wdOptions.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, wdOptions, new Object[] { value });

        // Allow the value on the options to be reset
        public void ResetValue(ref Options wdOptions)
        {
            if (Equals(Value, OriginalValue))
                return;

            try
            {
                SetProperty(ref wdOptions, Name, OriginalValue);
            }
            catch
            {
                // We may be setting word options that are not available in the version of word
                // being used, so just skip these errors
            }
        }
    }

    internal static class AppOptionFactory
    { 
        public static AppOption<T> Create<T>(string name, T value, ref Options wdOptions) =>
            new AppOption<T>(name, value, ref wdOptions);
    }
}
