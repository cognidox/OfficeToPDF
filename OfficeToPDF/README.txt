Office To PDF
=============

This is the Cognidox Ltd Office To PDF tool. It can be used to convert
Microsoft Office 2003, 2007, 2010, 2013 or 2016 documents to PDF from the
command line.

In order to run the tool, .net 4 and one of MS Office 2007, 2010, 2013
or 2016 must be installed.

It is distributed under the Apache License v2.0:
http://opensource.org/licenses/apache2.0

More information about installation and usage can be found on the Cognidox
project page: 
https://www.cognidox.com/officetopdf-converter/


Supported file types:
---------------------

The following file types can be converted:

 * Word       - .doc, .dot,  .docx, .dotx, .docm, .dotm, .rtf, .odt, .txt, .htm, .html, .wpd
 * Excel      - .xls, .xlsx, .xlsm, .xlt, .xltm, .xltx, .xlsb, .csv, .ods
 * Powerpoint - .ppt, .pptx, .pptm, .pot, .potm, .potx, .pps, .ppsx, .ppsm, .odp
 * Visio      - .vsd, .vsdm, .vsdx, [.vdx, .vdw, .dwg, .dfx, .wmf, .emf, .emz, .svg require Visio >= 2013]
 * Publisher  - .pub
 * Outlook    - .msg, .vcf, .ics
 * Project    - .mpp [requires Project >= 2010]


Usage:
------

In order to use the tool, simply call officetopdf.exe with two arguments -
the source Office document and the destination PDF document. e.g.

officetopdf.exe somefile.docx somefile.pdf

Switches:
---------

The following optional switches can be used:

  /bookmarks    - create bookmarks in the PDF when they are supported by the Office application
  /readonly     - attempts to open the source document in read-only mode
  /hidden       - attempts to minimise the Office application when converting
  /markup       - show document markup when creating PDFs with Word
  /print        - create high-quality PDFs optimised for print
  /pdfa         - produce ISO 19005-1 (PDF/A) compliant PDFs
  /excludeprops - Do not include properties in the PDF
  /excludetags  - Do not include tags in the PDF
  /noquit       - Do not quit already running Office applications once the conversion is done
  /verbose      - print out messages as it runs
  /password <pass>          - use <pass> as the password to open the document with
  /writepassword <pass>     - use <pass> as the write password to open the document with
  /merge                    - when using a template, create a new file from the template and merge
                              the text from the document to convert into the new file
  /template <template_path> - use a .dot, .dotx or .dotm template when converting with Word
  /excel_active_sheet       - only convert the active worksheet
  /excel_auto_macros        - run Auto_Open macros in Excel files before conversion
  /excel_show_formulas      - show formulas in the PDF when converting with Excel
  /excel_show_headings      - show row and column headings
  /excel_max_rows <rows>    - if any worksheet in a spreadsheet document has more
                              than this number of rows, do not attempt to convert
                              the file. Applies when converting with Excel
  /excel_worksheet <num>    - only convert worksheet <num> in the workbook. First sheet is 1
  /word_header_dist <pts>   - the distance (in points) from the header to the top of
                              the page
  /word_footer_dist <pts>   - the distance (in points) from the footer to the bottom
                              of the page
  /word_ref_fonts           - when fonts are not available, a reference to the font is used in
                              the generated PDF rather than a bitmapped version. The default is
                              for a bitmap of the text to be used
  /pdf_clean_meta <type>    - Allows for some meta-data to be removed from the generated PDF.
                              <type> can be:
                                basic - removes author, keywords, creator and subject
                                full  - removes all that basic removes and also the title
  /pdf_layout <layout>      - controls how the pages layout in Acrobat Reader. <layout> can be
                              one of the following values:
                                onecol       - show pages as a single scolling column
                                single       - show pages one at a time
                                twocolleft   - show pages in two columns, with oddnumbered pages on the left
                                twocolright  - show pages in two columns, with oddnumbered pages on the right
                                twopageleft  - show pages two at a time, with odd-numbered pages on the left
                                twopageright - show pages two at a time, with odd-numbered pages on the right
  /pdf_page_mode <mode>     - controls how the PDF will open with Acrobat Reader. <mode> can be
                              one of the following values:
                                full      - the PDF will open in fullscreen mode
                                bookmarks - the PDF will open with the bookmarks visible
                                thumbs    - the PDF will open with the thumbnail view visible
                                none      - the PDF will open without the navigation bar visible
  /pdf_append               - append the generated PDF to the end of the PDF destination
  /pdf_prepend              - prepend the generated PDF to the start of the PDF destination
  /version                  - print out the version of OfficeToPDF and exit


  Templates:
  ----------

  When converting documents with Word, the /template switch can be used to open
  the source document with a specific Word template file. If no /template switch
  is set, the default Normal.dotm template will be used.
  
  The template path must be the full path to the template file. e.g.
    /template c:\users\fred\Application Data\Microsoft\Templates\example.dotx
  
  For more information about Office template paths, see:
  http://office.microsoft.com/en-001/word-help/about-document-template-locations-HP003082725.aspx


  Error Codes:
  ------------

  The following error codes are returned by OfficeToPDF. Note that multiple errors are
  returned as a bitmask, so bitwise operations can test for multiple errors.

  0     - Success
  1     - Failure
  2     - Unknown Error
  4     - File protected by password
  8     - Invalid arguments
  16    - Unable to open the source file
  32    - Unsupported file format
  64    - Source file not found
  128   - Output directory not found
  256   - The requested worksheet was not found
  512   - Unable to use an empty worksheet

  To check for a specific error code after calling officetopdf.exe, use the batch
  "SET /A" command. e.g.

      SET /A "PASSWORDFAIL=(%ERRORLEVEL% & 4)"
      IF %PASSWORDFAIL% NEQ 0 (
          ECHO Password failed
      )


  Credits:
  --------

  Cognidox would like to thank all the people who have made suggestions
  through CodePlex.

  OfficeToPDF use the PDFSharp library: http://pdfsharp.codeplex.com/. This
  is licensed using the MIT License: http://pdfsharp.codeplex.com/license
