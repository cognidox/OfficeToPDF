namespace OfficeToPDF
{
    internal class ConverterFactory
    {
        public IConverter Create(string extension)
        {
            switch (extension)
            {
                case "rtf":
                case "odt":
                case "doc":
                case "dot":
                case "docx":
                case "dotx":
                case "docm":
                case "dotm":
                case "txt":
                case "html":
                case "htm":
                case "wpd":
                    return new WordConverter();

                case "csv":
                case "ods":
                case "xls":
                case "xlsx":
                case "xlt":
                case "xltx":
                case "xlsm":
                case "xltm":
                case "xlsb":
                    return new ExcelConverter();

                case "odp":
                case "ppt":
                case "pptx":
                case "pptm":
                case "pot":
                case "potm":
                case "potx":
                case "pps":
                case "ppsx":
                case "ppsm":
                    return new PowerpointConverter();

                case "vsd":
                case "vsdm":
                case "vsdx":
                case "vdx":
                case "vdw":
                case "svg":
                case "emf":
                case "emz":
                case "dwg":
                case "dxf":
                case "wmf":
                    return new VisioConverter();

                case "pub":
                    return new PublisherConverter();

                case "xps":
                    return new XpsConverter();

                case "msg":
                case "vcf":
                case "ics":
                    return new OutlookConverter();

                case "mpp":
                    return new ProjectConverter();
            }

            return new NullConverter(extension);
        }
    }
}
