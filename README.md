# OfficeToPDF

## Giving back to the community

[Cognidox](https://www.cognidox.com/) would not exist without the help of many open source projects. Perl,
Apache and Solr are just a few of the excellent open source packages that help make CogniDox a leading document
management system. So, to show our appreciation, we've released a couple of open source projects such as
OfficeToPDF to help others.

---
## OfficeToPDF - what does it do?

OfficeToPDF is a **command line utility** that converts Microsoft Office 2003, 2007, 2010, 2013 and 2016
documents from their native format into PDF using Office's in-built PDF export features.

Most Office to PDF converter tools are intended as single-user desktop applications. OfficeToPDF is
useful (and unique) if you want to automatically create PDF files on a server-wide basis and free
individual users from an extra step of using the "Save as..."
command on their Office files. These PDF files can then be stored and managed on a separate server. 
This can be useful if, for example, a department has a policy of only distributing PDF versions of 
documents to people outside the department.

There are some technical requirements that must be met before you can use it:</p>

* .NET Framework 4
* Office 2016, 2013, 2010 **or** Office 2007

If you are using Office 2007, you will also need:

* Visual Studio 2010 Tools for Office Runtime [<a href="https://www.microsoft.com/en-GB/download/details.aspx?id=48217" target="_blank">Download</a>]
* 2007 Microsoft Office Add-in: Microsoft Save as PDF or XPS [<a href="http://www.microsoft.com/downloads/en/details.aspx?familyid=4d951911-3e7e-4ae6-b059-a2e79ed87041&displaylang=en" target="_blank">Download</a>]

It is distributed under the [**Apache 2.0**](https://github.com/cognidox/OfficeToPDF/blob/master/LICENSE.md) license.

---

## Supported File Types

The following file types can be converted:

* Word (.doc, .dot,&nbsp; .docx, .dotx, .docm, .dotm, .rtf, .wpd)
* Excel&nbsp; (.xls, .xlsx, .xlsm, .xlsb, .xlt, .xltx, .xltm, .csv)
* Powerpoint (.ppt, .pptx, .pptm, .pps, .ppsx, .ppsm, .pot, .potx, .potm)
* Visio (.vsd, .vsdx, .vsdm, .svg) [Requires &gt;= Visio 2013 for .svg, .vsdx and .vsdm support]
* Publisher (.pub) </li><li>Outlook (.msg, .vcf, .ics)
* Project (.mpp) [Requires Project &gt;= 2010 for .mpp support]
* OpenOffice (.odt, .odp, .ods)

Conversion of Visio, Publisher and Project&nbsp;files require that the Visio, Publisher and Project
applications are installed. These are not included in the Office standard package.

---

## Instructions

In order to use the tool, download the officetopdf.exe file and, from the command line, run officetopdf.exe with two arguments - 
the source Office document and the destination PDF document. e.g.

    Microsoft Windows [Version 6.1.7601]
    Copyright (c) 2009 Microsoft Corporation. All rights reserved.

    C:\Users\test> officetopdf.exe somefile.docx somefile.pdf


### Command line switches

The following optional switches can be used:

| Switch | Description |
| ------ | ----------- |
| /bookmarks                 | create bookmarks in the PDF when they are supported by the Office application |
| /readonly                  | attempts to open the source document in read-only mode |
| /print                     | create high-quality PDFs optimised for print |
| /hidden                    | attempts to minimise the Office application when converting |
| /template _template_       | use a .dot, .dotx or .dotm template when converting with Word |
| /markup                    | show document markup when creating PDFs with Word |
| /pdfa                      | produce ISO 19005-1 (PDF/A) compliant PDFs |
| /noquit                    | do not exit running Office applications |
| /excludeprops              | do not include properties in generated PDF |
| /excludetags               | do not include tags in generated PDF |
| /password _password_       | provide a read password to open the file with |
| /writepassword _password_  | provide a read/write password to open the file with |
| /merge                     | when using a template, create a new file from the template and merge the text from the document to convert into the new file |
| /excel_show_formulas       | show formulas in Excel |
| /excel_show_headings       | shows column and row headings |
| /excel_max_rows _rows_     | allow a limit on the number of rows to convert |
| /excel_active_sheet        | only convert the currently active worksheet in a spreadsheet |
| /excel_worksheet _num_     |  only convert worksheet _num_ in the workbook. First sheet is 1 |
| /excel_auto_macros         | run Auto_Open macros in Excel files before conversion |
| /excel_no_link_update      | do not update links when opening Excel files |
| /excel_no_recalculate      | skip automatic re-calculation of formulas in the workbook |
| /word_header_dist _pts_    | the distance (in points) from the header to the top of the page |
| /word_footer_dist _pts_    | the distance (in points) from the footer to the bottom of the page |
| /word_field_quick_update   | perform a fast update of fields in Word before conversion |
| /word_fix_table_columns    | update table column widths to match table heading column widths |
| /word_keep_history         | do not clear Word's recent files list |
| /word_max_pages _pages_    | do not attempt conversion of a Word document if it has more than this number of _pages_ |
| /word_no_field_update      | do not update fields when creating the PDF |
| /word_ref_fonts            | when fonts are not available, a reference to the font is used in the generated PDF rather than a bitmapped version. The default is for a bitmap of the text to be used |
| /word_show_comments        | show comments when /markup is used |
| /word_show_revs_comments   | show revisions and comments when /markup is used |
| /word_show_format_changes  | show format changes when /markup is used |
| /word_show_ink_annot       | show ink annotations when /markup is used |
| /word_show_ins_del         | show all markup when /markup is used |
| /word_show_all_markup      | show all markup content when /markup is used |
| /word_markup_balloon       | show balloon style markup messages rather than inline ones |
| /fallback_printer <name>   | print the document to postscript printer <name> for conversion when the main conversion routine fails. Requires Ghostscript to be installed |
| /printer <name>            | print the document to postscript printer <name> for conversion. Requires Ghostscript to be installed |
| /pdf_clean_meta _type_     | allows for some meta-data to be removed from the generated PDF<br>_type_ can be:<ul><li>basic - removes author, keywords, creator and subject</li><li>full - removes all that basic removes and also the title</li></ul> |
| /pdf_layout _layout_       | controls how the pages layout in Acrobat Reader<br>_layout_ can be one of the following values:<ul><li>onecol - show pages as a single scrolling column</li><li>single - show pages one at a time</li><li>twocolleft - show pages in two columns, with odd-numbered pages on the left</li><li>twocolright - show pages in two columns, with odd-numbered pages on the right</li><li>twopageleft - show pages two at a time, with odd-numbered pages on the left</li><li>twopageright - show pages two at a time, with odd-numbered pages on the right</li></ul> |
| /pdf_page_mode _mode_      | controls how the PDF will open with Acrobat Reader<br>_mode_ can be one of the following values:<ul><li>full - the PDF will open in fullscreen mode</li><li>bookmarks - the PDF will open with the bookmarks visible</li><li>thumbs - the PDF will open with the thumbnail view visible</li><li>none - the PDF will open without the navigation bar visible</li></ul> |
| /pdf_append                | append the generated PDF to the end of the PDF destination |
| /pdf_prepend               | prepend the generated PDF to the start of the PDF destination |
| /pdf_owner_pass _pass_     | set the owner password on the PDF. Needed to make modifications to the PDF |
| /pdf_user_pass _pass_      | set the user password on the PDF. Needed to open the PDF |
| /pdf_restrict_accessibility_extraction | Prevent all content extraction without the owner password |
| /pdf_restrict_annotation   | prevent annotations on the PDF without the owner password |
| /pdf_restrict_assembly     |  prevent rotation, removal or insertion of pages without the owner password |
| /pdf_restrict_extraction   |  prevent content extraction without the owner password |
| /pdf_restrict_forms        | prevent form entry without the owner password |
| /pdf_restrict_full_quality | prevent full quality printing without the owner password |
| /pdf_restrict_modify       | prevent modification without the owner password |
| /pdf_restrict_print        | prevent printing without the owner password |
| /verbose                   | print out messages as it runs |
| /version                   | print out the version of OfficeToPDF and exit |
| /working_dir _path_        | a path to copy the input file into temporarily when running the conversion |

---

## Error Codes

The following error codes are returned by OfficeToPDF. Note that multiple errors are returned as a bitmask,
so bitwise operations can test for multiple errors.

```
0 - Success
1 - Failure
2 - Unknown Error
4 - File protected by password
8 - Invalid arguments
16 - Unable to open the source file
32 - Unsupported file format
64 - Source file not found
128 - Output directory not found
256 - The requested worksheet was not found
512 - Unable to use an empty worksheet
1024 - Unable to modify or open a protected PDF
2048 - Raised when there is a problem calling an Office application
4096 - There are no printers installed, so Office conversion can not proceed
```

---

## About CogniDox

CogniDox is a web-based, document management software tool aimed primarily at supporting high-tech
product development. It enables better product lifecycle management and knowledge transfer from
developers to partners, clients and customers.

We provide highly-integrated support for system engineering workflows and the product lifecycle.
Plug-ins are provided for CMS products, software SCM tools, EDA tools, CRM systems, CAD tools and 
Help Desk applications. Our "Xtranet" solution enables companies to add a secure self-service
customer portal onto their public web sites.

Based in Cambridge, UK since its formation in 2008, CogniDox has an impressive portfolio of
blue chip users &ndash; see examples at https://www.cognidox.com/customers

