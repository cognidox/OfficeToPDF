using NUnit.Framework;
using PdfSharp.Pdf;
using System.Collections.Generic;
using System.Diagnostics;

namespace OfficeToPDF.Tests
{
    [TestFixture]
    public class ArgParserTests
    {
        [Test]
        public void WhenConstructedThenSetContainsTheExpectedNumber()
        {
            ArgParser parser = new ArgParser();

            Assert.That(parser.Count, Is.EqualTo(67));
        }

        [Test]
        public void WhenConstructedTheSetContainsTheExpectedKeys()
        {
            var expected = new[]
            {
                #region Key values
                "bookmarks",
                "excel_active_sheet_on_max_rows",
                "excel_active_sheet",
                "excel_auto_macros",
                "excel_delay",
                "excel_max_rows",
                "excel_no_link_update",
                "excel_no_map_papersize",
                "excel_no_recalculate",
                "excel_show_formulas",
                "excel_show_headings",
                "excel_template_macros",
                "excel_worksheet",
                "excludeprops",
                "excludetags",
                "fallback_printer",
                "has_working_dir",
                "hidden",
                "markup",
                "merge",
                "noquit",
                "original_basename",
                "original_filename",
                "password",
                "pdf_clean_meta",
                "pdf_layout",
                "pdf_merge",
                "pdf_owner_pass",
                "pdf_page_mode",
                "pdf_restrict_accessibility_extraction",
                "pdf_restrict_annotation",
                "pdf_restrict_assembly",
                "pdf_restrict_extraction",
                "pdf_restrict_forms",
                "pdf_restrict_full_quality",
                "pdf_restrict_modify",
                "pdf_restrict_print",
                "pdf_user_pass",
                "pdfa",
                "powerpoint_output",
                "print",
                "printer",
                "readonly",
                "screen",
                "template",
                "timeout",
                "verbose",
                "word_field_quick_update_safe",
                "word_field_quick_update",
                "word_fix_table_columns",
                "word_footer_dist",
                "word_header_dist",
                "word_keep_history",
                "word_markup_balloon",
                "word_max_pages",
                "word_no_field_update",
                "word_no_map_papersize",
                "word_no_repair",
                "word_ref_fonts",
                "word_show_all_markup",
                "word_show_comments",
                "word_show_format_changes",
                "word_show_hidden",
                "word_show_ink_annot",
                "word_show_ins_del",
                "word_show_revs_comments",
                "working_dir"
                #endregion
            };

            ArgParser parser = new ArgParser();

            Assert.That(parser.Keys, Is.EquivalentTo(expected));
        }

        [Test]
        public void WhenParsingVersionArgThenExitCalled()
        {
            var exit = false;

            ArgParser parser = new ArgParser(_ => exit = true, (format, args) => Trace.WriteLine(string.Format(format, args)));

            parser.Parse(new[] { "/version" }, new Dictionary<string, bool>());

            Assert.That(exit, Is.True);
        }

        [Test]
        public void WhenParsingQueryArgThenHelpShown()
        {
            var exit = false;
            var captured = "";

            ArgParser parser = new ArgParser(_ => exit = true, (format, args) => { captured = format; Trace.WriteLine(string.Format(format, args)); });

            parser.Parse(new[] { "/?" }, new Dictionary<string, bool>());

            Assert.That(exit, Is.True);
            Assert.That(captured.StartsWith("Converts Office documents to PDF from the command line."));
        }

        [Test]
        public void WhenParsingHelpArgThenHelpShown()
        {
            var exit = false;
            var captured = "";

            ArgParser parser = new ArgParser(_ => exit = true, (format, args) => { captured = format; Trace.WriteLine(string.Format(format, args)); });

            parser.Parse(new[] { "/help" }, new Dictionary<string, bool>());

            Assert.That(exit, Is.True);
            Assert.That(captured.StartsWith("Converts Office documents to PDF from the command line."));
        }

        [Test]
        public void WhenParsingTimeoutArgThenResultIsSuccess()
        {
            ArgParser parser = new ArgParser();

            var result = parser.Parse(new[] { "/timeout", "100" }, new Dictionary<string, bool>());

            Assert.That(result, Is.EqualTo(ExitCode.Success));
        }

        [Test]
        public void WhenParsingTimeoutArgThenValueCaptured()
        {
            ArgParser parser = new ArgParser();

            parser.Parse(new[] { "/timeout", "100" }, new Dictionary<string, bool>());

            Assert.That(parser.ContainsKey("timeout"), Is.True);
            Assert.That(parser["timeout"], Is.EqualTo(100));
        }

        [Test]
        public void WhenParsingTimeoutArgThenPropertyContainsValue()
        {
            ArgParser parser = new ArgParser();

            parser.Parse(new[] { "/timeout", "100" }, new Dictionary<string, bool>());

            Assert.That(parser.ContainsKey("timeout"), Is.True);
            Assert.That(parser.timeout, Is.EqualTo(100));
        }

        public static IEnumerable<object[]> PdfPageModeArgs =>
            new[]
            {
                new object[] { "full", PdfPageMode.FullScreen },
                new object[] { "none", PdfPageMode.UseNone },
                new object[] { "bookmarks", PdfPageMode.UseOutlines },
                new object[] { "thumbs", PdfPageMode.UseThumbs },
                new object[] { default(string), default(PdfPageMode?) }
            };

        [TestCaseSource(nameof(PdfPageModeArgs))]
        public void WhenParsingPdfPageModeArgThenPropertyContainsCorrectValue(string option, PdfPageMode? expected)
        {
            ArgParser parser = new ArgParser();

            if (!string.IsNullOrEmpty(option))
                parser.Parse(new[] { "/pdf_page_mode", option }, new Dictionary<string, bool>());

            Assert.That(parser.ContainsKey("pdf_page_mode"), Is.True);
            Assert.That(parser.pdf_page_mode, Is.EqualTo(expected));
        }

        [TestCase("/readonly", true)]
        [TestCase(default(string), false)]
        public void WhenParsingReadonlyArgThenPropertyContainsCorrectValue(string option, bool expected)
        {
            ArgParser parser = new ArgParser();

            if (!string.IsNullOrEmpty(option))
                parser.Parse(new[] { option }, new Dictionary<string, bool>());

            Assert.That(parser.ContainsKey("readonly"), Is.True);
            Assert.That(parser.@readonly, Is.EqualTo(expected));
        }
    }
}
