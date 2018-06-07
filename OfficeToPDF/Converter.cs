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
using System.Text;
using System.Threading;
using System.Xml.Linq;
using System.Windows.Forms;
using Ghostscript.NET.Processor;
using OpenMcdf;
using OpenMcdf.Extensions;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeToPDF
{
    public delegate void PrintDocument(string destination, string printerName);

    /// <summary>
    /// Base converter class that all conversion handlers implement
    /// </summary>
    class Converter
    {
        private enum DocumentFileTypes
        {
            Unknown = 0,
            CDF = 1,
            OpenXML = 2
        }

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
        protected static void ReleaseCOMObject(object obj)
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
            {
                return IsPasswordProtected(stream);
            }
        }

        /// <summary>
        /// Detects if a given office document is protected by a password or not.
        /// Supported formats: Word, Excel and PowerPoint (both legacy and OpenXml).
        /// </summary>
        /// <param name="stream">Office document stream.</param>
        /// <returns>True if document is protected by a password, false otherwise.</returns>
        protected static bool IsPasswordProtected(Stream stream)
        {
            var compObjHeader = new byte[0x20];
            if (FileStreamDocumentType(stream, ref compObjHeader) != DocumentFileTypes.CDF)
            {
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
            {
                return true;
            }

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
                    {
                        return false; // invalid document
                    }
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

        // Return what we thing the document type is
        private static DocumentFileTypes FileStreamDocumentType(Stream stream)
        {
            byte[] compObjHeader = new byte[0x20];
            return FileStreamDocumentType(stream, ref compObjHeader);
        }

        // Return what we thing the document type is plus a filled in byte array of the document header
        private static DocumentFileTypes FileStreamDocumentType(Stream stream, ref byte[] compObjHeader)
        {
            // Minimum file size for office file is 4k
            if (stream.Length < 4096)
            {
                return DocumentFileTypes.Unknown;
            }

            // read file header
            stream.Seek(0, SeekOrigin.Begin);
            
            ReadFromStream(stream, compObjHeader);
            
            // check if we have plain zip file
            if (compObjHeader[0] == 'P' && compObjHeader[1] == 'K')
            {
                // this is a plain OpenXml document (not encrypted)
                return DocumentFileTypes.OpenXML;
            }

            // check compound object magic bytes
            if (compObjHeader[0] != 0xD0 || compObjHeader[1] != 0xCF)
            {
                // unknown document format
                return DocumentFileTypes.Unknown;
            }
            return DocumentFileTypes.CDF;
        }
        
        /// <summary>
        /// Detects if a given office document is should be opened in read-only
        /// mode based on the file extended properties
        /// </summary>
        /// <param name="filename">Office document path.</param>
        /// <returns>True if document is should be opened read-only, false otherwise.</returns>
        protected static bool IsReadOnlyEnforced(String filename)
        {
            DocumentFileTypes documentType;
            using (var stream = File.OpenRead(filename))
            {
                documentType = FileStreamDocumentType(stream);
            }
            switch (documentType)
            {
                case DocumentFileTypes.CDF:
                    return IsCDFReadOnlyEnforced(filename);
                case DocumentFileTypes.OpenXML:
                    return IsOpenXMLReadOnlyEnforced(filename);
                default:
                    return false;
            }
        }

        // Return true if a compound document is enforcing read-only
        // Looks in the Summary Information stream
        private static bool IsCDFReadOnlyEnforced(string filename)
        {
            CompoundFile cf = new CompoundFile(fileName: filename);
            if (null == cf)
            {
                return false;
            }
            CFStream summaryInfo = cf.RootStorage.GetStream("\x05SummaryInformation");
            if (null != summaryInfo)
            {
                // Interested in the doc security setting
                OpenMcdf.Extensions.OLEProperties.PropertySetStream ps = summaryInfo.AsOLEProperties();
                int securityIdx = ps.PropertySet0.PropertyIdentifierAndOffsets.FindIndex(x => x.PropertyIdentifier == OpenMcdf.Extensions.OLEProperties.PropertyIdentifiersSummaryInfo.PIDSI_DOC_SECURITY);
                if (securityIdx >= 0 && securityIdx < ps.PropertySet0.Properties.Count) {
                    int security = (int)ps.PropertySet0.Properties[securityIdx].PropertyValue;
                    // See if read-only is enforced https://msdn.microsoft.com/en-us/library/windows/desktop/aa371587(v=vs.85).aspx
                    if ((security & 4) == 4)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        // Return true if an open XML document is enforcing read-only
        private static bool IsOpenXMLReadOnlyEnforced(string filename)
        {
            // Read an OpenXML type document
            using (Package package = Package.Open(path: filename, packageMode: FileMode.Open, packageAccess: FileAccess.Read))
            {
                if (null == package)
                {
                    return false;
                }
                try
                {
                    // Document security is set in the extended properties
                    // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc845474(v%3doffice.14)
                    string extendedType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
                    PackageRelationshipCollection extendedProps = package.GetRelationshipsByType(extendedType);
                    if (null != extendedProps)
                    {
                        IEnumerator extendedPropsList = extendedProps.GetEnumerator();
                        if (extendedPropsList.MoveNext())
                        {
                            Uri extendedPropsUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), ((PackageRelationship)extendedPropsList.Current).TargetUri);
                            PackagePart props = package.GetPart(extendedPropsUri);
                            if (null != props)
                            {
                                // Read the internal docProps/app.xml XML file
                                XDocument xmlDoc = XDocument.Load(props.GetStream());
                                XElement securityEl = xmlDoc.Root.Element(XName.Get("DocSecurity", xmlDoc.Root.GetDefaultNamespace().NamespaceName));
                                if (null != securityEl)
                                {
                                    if (!String.IsNullOrWhiteSpace(securityEl.Value))
                                    {
                                        // See https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc840043%28v%3doffice.14%29
                                        return ((Int16.Parse(securityEl.Value) & 4) == 4);
                                    }
                                }

                                package.Close();
                                // PowerPoint doesn't use DocSecurity (*sigh*) so need another check
                                XElement appEl = xmlDoc.Root.Element(XName.Get("Application", xmlDoc.Root.GetDefaultNamespace().NamespaceName));
                                if (null != appEl)
                                {
                                    if (!String.IsNullOrWhiteSpace(appEl.Value) &&
                                        appEl.Value.IndexOf("PowerPoint", StringComparison.InvariantCultureIgnoreCase) >= 0)
                                    {
                                        PresentationDocument presentationDocument = PresentationDocument.Open(path: filename, isEditable: false);
                                        if (null != presentationDocument &&
                                            presentationDocument.PresentationPart.Presentation.ModificationVerifier != null) {
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception) { }
                return false;
            }
        }

        // Print out a document to a file which we will pass to ghostscript
        protected static void PrintToGhostscript(string printer, string outputFilename, PrintDocument printFunc)
        {
            String postscriptFilePath = "";
            String postscriptFile = "";
            try
            {
                // Create a temporary location to output to
                postscriptFilePath = Path.GetTempFileName();
                File.Delete(postscriptFilePath);
                Directory.CreateDirectory(postscriptFilePath);
                postscriptFile = Path.Combine(postscriptFilePath, Guid.NewGuid() + ".ps");

                // Set up the printer
                PrintDialog printDialog = new PrintDialog
                {
                    AllowPrintToFile = true,
                    PrintToFile = true
                };
                System.Drawing.Printing.PrinterSettings printerSettings = printDialog.PrinterSettings;
                printerSettings.PrintToFile = true;
                printerSettings.PrinterName = printer;
                printerSettings.PrintFileName = postscriptFile;

                // Call the appropriate printer function (changes based on the office application)
                printFunc(postscriptFile, printerSettings.PrinterName);
                ReleaseCOMObject(printerSettings);
                ReleaseCOMObject(printDialog);
                
                // Call ghostscript
                GhostscriptProcessor gsproc = new GhostscriptProcessor();
                List<string> gsArgs = new List<string>
                    {
                        "gs",
                        "-dBATCH",
                        "-dNOPAUSE",
                        "-dQUIET",
                        "-dSAFER",
                        "-dNOPROMPT",
                        "-sDEVICE=pdfwrite",
                        String.Format("-sOutputFile=\"{0}\"", string.Join(@"\\", outputFilename.Split(new string[] { @"\" }, StringSplitOptions.None))),
                        @"-f",
                        postscriptFile
                    };
                gsproc.Process(gsArgs.ToArray());
            }
            finally {
                // Clean up the temporary files
                if (!String.IsNullOrWhiteSpace(postscriptFilePath) && Directory.Exists(postscriptFilePath))
                {
                    if (!String.IsNullOrWhiteSpace(postscriptFile) && File.Exists(postscriptFile))
                    {
                        // Make sure ghostscript is not holding onto the postscript file
                        for (var i = 0; i < 60; i++)
                        {
                            try
                            {
                                File.Delete(postscriptFile);
                                break;
                            }
                            catch (IOException)
                            {
                                Thread.Sleep(500);
                            }
                        }
                    }
                    Directory.Delete(postscriptFilePath);
                }
            }
        }

    }
}
