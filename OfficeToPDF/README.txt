Office To PDF
=============

This is the Cognidox Ltd Office To PDF tool. It can be used to convert
Microsoft Office 2003, 2007, 2010 & 2013 documents to PDF from the
command line.

In order to run the tool, .net 4 and one of MS Office 2007, 2010 or 2013
must be installed.

It is distributed under the Apache License v2.0:
http://opensource.org/licenses/apache2.0

More information about installation and usage can be found on the Cognidox
project page: 
http://www.cognidox.com/opensource/officetopdf


Supported file types:
---------------------

The following file types can be converted:

 * Word       - .doc, .dot,  .docx, .dotx, .docm, .dotm, .rtf, .odt
 * Excel      - .xls, .xlsx, .xlsm, .csv, .odc
 * Powerpoint - .ppt, .pptx, .pptm, .odp
 * Visio      - .vsd
 * Publisher  - .pub
 * Outlook    - .msg, .vcf, .ics
 * Project    - .mpp (Requires Office 2010 or greater)


Usage:
------

In order to use the tool, simply call officetopdf.exe with two arguments -
the source Office document and the destination PDF document. e.g.

officetopdf.exe somefile.docx somefile.pdf

Switches:
---------

The following optional switches can be used:

  /bookmarks - create bookmarks in the PDF when they are supported by the Office application
  /readonly  - attempts to open the source document in read-only mode
  /hidden    - attempts to minimise the Office application when converting
  /print     - create high-quality PDFs optimised for print
  /pdfa      - produce ISO 19005-1 (PDF/A) compliant PDFs
  /verbose   - Print out messages as it runs
  /template <template_path> - use a .dot, .dotx or .dotm template when converting with Word


  Templates:
  ----------

  When converting documents with Word, the /template switch can be used to open
  the source document with a specific Word template file. If no /template switch
  is set, the default Normal.dotm template will be used.
  
  The template path must be the full path to the template file. e.g.
    /template c:\users\fred\Application Data\Microsoft\Templates\example.dotx
  
  For more information about Office template paths, see:
  http://office.microsoft.com/en-001/word-help/about-document-template-locations-HP003082725.aspx