using NUnit.Framework;
using System.Collections.Generic;

namespace OfficeToPDF.Tests
{
    [TestFixture]
    public class ConverterFactoryTests
    {
        public static IEnumerable<object[]> WordExtensions =>
            new[]
            {
                new object[] { "rtf" },
                new object[] { "odt" },
                new object[] { "doc" },
                new object[] { "dot" },
                new object[] { "docx" },
                new object[] { "dotx" },
                new object[] { "docm" },
                new object[] { "dotm" },
                new object[] { "txt" },
                new object[] { "html" },
                new object[] { "htm" },
                new object[] { "wpd" },
            };

        [TestCaseSource(nameof(WordExtensions))]
        public void FactoryReturnsCorrectTypeForWordExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<WordConverter>());
        }

        public static IEnumerable<object[]> ExcelExtensions =>
            new[]
            {
                new object[] { "csv" },
                new object[] { "ods" },
                new object[] { "xls" },
                new object[] { "xlsx" },
                new object[] { "xlt" },
                new object[] { "xltx" },
                new object[] { "xlsm" },
                new object[] { "xltm" },
                new object[] { "xlsb" }
            };

        [TestCaseSource(nameof(ExcelExtensions))]
        public void FactoryReturnsCorrectTypeForExcelExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<ExcelConverter>());
        }

        public static IEnumerable<object[]> PowerPointExtensions =>
            new[]
            {
                new object[] { "odp" },
                new object[] { "ppt" },
                new object[] { "pptx" },
                new object[] { "pptm" },
                new object[] { "pot" },
                new object[] { "potm" },
                new object[] { "potx" },
                new object[] { "pps" },
                new object[] { "ppsx" },
                new object[] { "ppsm" }
            };

        [TestCaseSource(nameof(PowerPointExtensions))]
        public void FactoryReturnsCorrectTypeForPowerPointExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<PowerpointConverter>());
        }

        public static IEnumerable<object[]> VisioExtensions =>
            new[]
            {
                new object[] { "vsd" },
                new object[] { "vsdm" },
                new object[] { "vsdx" },
                new object[] { "vdx" },
                new object[] { "vdw" },
                new object[] { "svg" },
                new object[] { "emf" },
                new object[] { "emz" },
                new object[] { "dwg" },
                new object[] { "dxf" },
                new object[] { "wmf" }
            };


        [TestCaseSource(nameof(VisioExtensions))]
        public void FactoryReturnsCorrectTypeForVisioExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<VisioConverter>());
        }

        public static IEnumerable<object[]> PublisherExtensions =>
            new[]
            {
                new object[] { "pub" }
            };


        [TestCaseSource(nameof(PublisherExtensions))]
        public void FactoryReturnsCorrectTypeForPublisherExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<PublisherConverter>());
        }

        public static IEnumerable<object[]> XpsExtensions =>
            new[]
            {
                new object[] { "xps" }
            };


        [TestCaseSource(nameof(XpsExtensions))]
        public void FactoryReturnsCorrectTypeForXpsExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<XpsConverter>());
        }

        public static IEnumerable<object[]> OutlookExtensions =>
            new[]
            {
                new object[] { "msg" },
                new object[] { "vcf" },
                new object[] { "ics" }
            };


        [TestCaseSource(nameof(OutlookExtensions))]
        public void FactoryReturnsCorrectTypeForOutlookExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<OutlookConverter>());
        }

        public static IEnumerable<object[]> ProjectExtensions =>
            new[]
            {
                new object[] { "mpp" }
            };


        [TestCaseSource(nameof(ProjectExtensions))]
        public void FactoryReturnsCorrectTypeForProjectExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<ProjectConverter>());
        }

        public static IEnumerable<object[]> InvalidExtensions =>
            new[]
            {
                new object[] { "xyz" },
                new object[] { "abc" },
                new object[] { "123" },
                new object[] { "_-&" },
                new object[] { "" }
            };


        [TestCaseSource(nameof(InvalidExtensions))]
        public void FactoryReturnsCorrectTypeForInvalidExtensions(string extension)
        {
            var factory = new ConverterFactory();

            var result = factory.Create(extension);

            Assert.That(result, Is.InstanceOf<NullConverter>());
        }
    }
}
