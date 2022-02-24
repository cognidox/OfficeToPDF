using System;
using System.Collections.Generic;

namespace OfficeToPDF
{
    internal class NullConverter : IConverter
    {
        public string Extension { get; }
        public NullConverter(string extension) => Extension = extension;

        int IConverter.Convert(string inputFile, string outputFile, ArgParser options, ref List<PDFBookmark> bookmarks)
        {
            if (options.verbose)
            {
                Console.WriteLine($"Unsupported document extension '{Extension}'.");
            }
            return (int)ExitCode.UnsupportedFileFormat;
        }
    }
}
