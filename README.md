# WordReader
Parses text from DOC &amp; DOCX files. 

There are two Projects: WordReader and WordReaderTest

## WordReader
The actual WordReader project. Created with MS Visual Studio 2015.
Builded in DLL.
Code is fully documented.
To use it You need to add the DLL in the using section (using WordReader).
There are two classes in DLL: docParser and docxParser.

### docParser
Parses DOC files (Word Binary). Retrieves text from the MainDocument section.
Usage (example is in WordReaderTest project):
  1. create the instance of the docParser. As a parameter use the path to the .doc document.
  2. check whether the specified document is read successfully - use docIsOK field (true when read successfully).
  3. get the text as a string - use getText method.

### docxParser
Parses DOCX files (Word extension to the Open XML). Retrieves text from the MainDocument.
Usage (example is in WordReaderTest project):
  1. create the instance of the docxParser. As a parameter use the path to the .docx document.
  2. check whether the specified document is read successfully - use docxIsOK field (true when read successfully).
  3. get the text as a string - use getText method.

## WordReaderTest
This is a simple testing project for the WordReader DLL.

## References
As information references for the WordReader project I used:
  1. [MS-CFB] - v20180912. Compound File Binary File Format. Release: September 12, 2018. Copyright © 2018 Microsoft Corporation.
  2. [MS-DOC] - v20190319. Word (.doc) Binary File Format. Release: March 19, 2019. Copyright © 2019 Microsoft Corporation.
  3. [MS-DOCX] - v20190319. Word Extensions to the Office Open XML (.docx) File Format. Release: March 19, 2019. Copyright © 2019 Microsoft Corporation.
  4. ECMA-376. Office Open XML File Formats. Part 1: Fundamentals. 1st Edition / December 2006.
  5. ECMA-376. Office Open XML File Formats. Part 2: Open Packaging Conventions. 1st Edition / December 2006.
  6. ECMA-376. Office Open XML File Formats. Part 3: Primer. 1st Edition / December 2006.
  7. ECMA-376. Office Open XML File Formats. Part 5: Markup Compatibility and Extensibility. 1st Edition / December 2006.
