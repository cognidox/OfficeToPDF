Office To PDF
=============

This is the Cognidox Ltd Office To PDF tool. It can be used to convert
Microsoft Office 2007 and 2010 documents to PDF from the command line.

In order to run the tool, MS Office 2007 or 2010 must be installed.

It is distributed under the Apache License v2.0:
http://opensource.org/licenses/apache2.0

More information about installation and usage can be found on the Cognidox
project page: 
http://www.cognidox.com/opensource/officetopdf


Supported file types:
---------------------

The following file types can be converted:

 * Word       - .doc, .dot,  .docx, .dotx, .docm, .dotm
 * Excel      - .xls, .xlsx, .xlsm
 * Powerpoint - .ppt, .pptx, .pptm
 * Visio      - .vsd
 * Publisher  - .pub
 * Outlook    - .msg


Usage:
------

In order to use the tool, simply call officetopdf.exe with two arguments -
the source Office document and the destination PDF document. e.g.

officetopdf.exe somefile.docx somefile.pdf

Switches:
---------

The following optional switches can be used:

  /bookmarks - create bookmarks in the PDF when they are supported by the Office application
  /readonly  - attempts to open the source documet in read-only mode
  /hidden    - attempts to minimise the Office application when converting
