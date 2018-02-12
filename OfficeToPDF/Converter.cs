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
using System.IO;
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
        /// <param name="options">A set of options passed in from the main program</param>
        /// <returns>0 on success, or an error code on failure</returns>
        public static int Convert(String inputFile, String outputFile, Hashtable options)
        {
            return (int)ExitCode.UnknownError;
        }

        /// <summary>
        /// Converts the input file to a PDF and updates a reference to a set of bookmarks
        /// </summary>
        /// <param name="inputFile">Full path of the input file</param>
        /// <param name="outputFile">Full path of the file to output PDF</param>
        /// <param name="options">A set of options passed in from the main program</param>
        /// <param name="bookmarks">A reference to bookmarks that need to be added to the PDF</param>
        /// <returns>0 on success, or an error code on failure</returns>
        public static int Convert(String inputFile, String outputFile, Hashtable options, ref List<PDFBookmark> bookmarks)
        {
            return (int)ExitCode.UnknownError;
        }

        // Clean up COM objects
        protected static void releaseCOMObject(object obj)
        {
            try
            {
                if (null != obj && System.Runtime.InteropServices.Marshal.IsComObject(obj))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                }
            }
            catch { }
            finally
            {
                obj = null;
            }
        }


        /// <summary>
        /// Detects if a given office document is protected by a password or not.
        /// Supported formats: Word, Excel and PowerPoint (both legacy and OpenXml).
        /// Source:
        /// http://stackoverflow.com/a/26522122
        /// </summary>
        /// <param name="fileName">Path to an office document.</param>
        /// <returns>True if document is protected by a password, false otherwise.</returns>
        protected static bool IsPasswordProtected(string fileName)
        {
            using (var stream = File.OpenRead(fileName))
                return IsPasswordProtected(stream);
        }

        /// <summary>
        /// Detects if a given office document is protected by a password or not.
        /// Supported formats: Word, Excel and PowerPoint (both legacy and OpenXml).
        /// </summary>
        /// <param name="stream">Office document stream.</param>
        /// <returns>True if document is protected by a password, false otherwise.</returns>
        protected static bool IsPasswordProtected(Stream stream)
        {
            // minimum file size for office file is 4k
            if (stream.Length < 4096)
                return false;

            // read file header
            stream.Seek(0, SeekOrigin.Begin);
            var compObjHeader = new byte[0x20];
            ReadFromStream(stream, compObjHeader);

            // check if we have plain zip file
            if (compObjHeader[0] == 'P' && compObjHeader[1] == 'K')
            {
                // this is a plain OpenXml document (not encrypted)
                return false;
            }

            // check compound object magic bytes
            if (compObjHeader[0] != 0xD0 || compObjHeader[1] != 0xCF)
            {
                // unknown document format
                return false;
            }

            int sectionSizePower = compObjHeader[0x1E];
            if (sectionSizePower < 8 || sectionSizePower > 16)
            {
                // invalid section size
                return false;
            }
            int sectionSize = 2 << (sectionSizePower - 1);

            const int defaultScanLength = 32768;
            long scanLength = Math.Min(defaultScanLength, stream.Length);

            // read header part for scan
            stream.Seek(0, SeekOrigin.Begin);
            var header = new byte[scanLength];
            ReadFromStream(stream, header);

            // check if we detected password protection
            if (ScanForPassword(stream, header, sectionSize))
                return true;

            // if not, try to scan footer as well

            // read footer part for scan
            stream.Seek(-scanLength, SeekOrigin.End);
            var footer = new byte[scanLength];
            ReadFromStream(stream, footer);

            // finally return the result
            return ScanForPassword(stream, footer, sectionSize);
        }

        static void ReadFromStream(Stream stream, byte[] buffer)
        {
            int bytesRead, count = buffer.Length;
            while (count > 0 && (bytesRead = stream.Read(buffer, 0, count)) > 0)
                count -= bytesRead;
            if (count > 0) throw new EndOfStreamException();
        }

        static bool ScanForPassword(Stream stream, byte[] buffer, int sectionSize)
        {
            const string afterNamePadding = "\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0";
            const string encryptedPackageName = "E\0n\0c\0r\0y\0p\0t\0e\0d\0P\0a\0c\0k\0a\0g\0e" + afterNamePadding;
            const string encryptedSummaryName = "E\0n\0c\0r\0y\0p\0t\0e\0d\0S\0u\0m\0m\0a\0r\0y" + afterNamePadding;
            const string wordDocumentName = "W\0o\0r\0d\0D\0o\0c\0u\0m\0e\0n\0t" + afterNamePadding;
            const string workbookName = "W\0o\0r\0k\0b\0o\0o\0k" + afterNamePadding;

            try
            {
                var bufferString = Encoding.ASCII.GetString(buffer, 0, buffer.Length);

                // try to detect password protection used in new OpenXml documents
                // by searching for "EncryptedPackage" or "EncryptedSummary" streams
                // (old .ppt documents use this stream as well)
                if (bufferString.Contains(encryptedPackageName) ||
                    bufferString.Contains(encryptedSummaryName))
                    return true;

                // try to detect password protection for legacy Office documents

                // check for Word header
                int headerOffset = bufferString.IndexOf(wordDocumentName, StringComparison.InvariantCulture);
                int sectionId;
                const int coBaseOffset = 0x200;
                const int sectionIdOffset = 0x74;
                if (headerOffset >= 0)
                {
                    sectionId = BitConverter.ToInt32(buffer, headerOffset + sectionIdOffset);
                    int sectionOffset = coBaseOffset + sectionId * sectionSize;
                    const int fibScanSize = 0x10;
                    if (sectionOffset + fibScanSize > stream.Length)
                        return false; // invalid document
                    var fibHeader = new byte[fibScanSize];
                    stream.Seek(sectionOffset, SeekOrigin.Begin);
                    ReadFromStream(stream, fibHeader);
                    short properties = BitConverter.ToInt16(fibHeader, 0x0A);
                    // check for fEncrypted FIB bit
                    const short fEncryptedBit = 0x0100;
                    return (properties & fEncryptedBit) == fEncryptedBit;
                }

                // check for Excel header
                headerOffset = bufferString.IndexOf(workbookName, StringComparison.InvariantCulture);
                if (headerOffset >= 0)
                {
                    sectionId = BitConverter.ToInt32(buffer, headerOffset + sectionIdOffset);
                    int sectionOffset = coBaseOffset + sectionId * sectionSize;
                    const int streamScanSize = 0x100;
                    if (sectionOffset + streamScanSize > stream.Length)
                        return false; // invalid document
                    var workbookStream = new byte[streamScanSize + sizeof(short)];
                    stream.Seek(sectionOffset, SeekOrigin.Begin);
                    ReadFromStream(stream, workbookStream);
                    short record = BitConverter.ToInt16(workbookStream, 0);
                    short recordSize = BitConverter.ToInt16(workbookStream, sizeof(short));
                    const short bofMagic = 0x0809;
                    const short eofMagic = 0x000A;
                    const short filePassMagic = 0x002F;
                    if (record != bofMagic)
                        return false; // invalid BOF
                    // scan for FILEPASS record until the end of the buffer
                    int offset = sizeof(short) * 2 + recordSize;
                    int recordsLeft = 16; // simple infinite loop check just in case
                    do
                    {
                        record = BitConverter.ToInt16(workbookStream, offset);
                        if (record == filePassMagic)
                            return true;
                        recordSize = BitConverter.ToInt16(workbookStream, sizeof(short) + offset);
                        offset += sizeof(short) * 2 + recordSize;
                        recordsLeft--;
                    } while (record != eofMagic && recordsLeft > 0);
                }
            }
            catch (Exception ex)
            {
                // BitConverter exceptions may be related to document format problems
                // so we just treat them as "password not detected" result
                if (ex is ArgumentOutOfRangeException)
                    return false;
                // respect all the rest exceptions
                throw;
            }

            return false;
        }
    }
}
