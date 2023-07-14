# A php class to extract all the text from a Word DOCX document and to output it as a text array

## Description

This php class will take a DOCX type Word document and extract all the text from it. The text will include all list and paragraph numbering and also footnotes and endnotes together with their reference numbers. The text will outputted as an array, one array element per paragraph. This will make it easy to search or manipulate the text or to save it to a database. For convenience the first element [0] of the array contains the number of text array elements and the length of the longest element in the format 'Number:Length'. In normal mode the class produces no output to the screen.

A demonstration file 'textdemo.php' is included. This expects the Word docx file to be called 'sample.docx'. The demonstration file will display on screen the resultant text array, giving the number of text elements, the length of the longest one and then all the text extracted from the document along with its array element number.

# USAGE

## Include the class in your php script
```
require_once('wordtext.php');
```

## Normal mode to save all the the text to an array (no output to screen)
```
$rt = new WordTEXT(false); or $rt = new WordTEXT();
```

## Debug mode to display on screen the associated DOCX XML files and the text extracted from the document
```
$rt = new WordTEXT(true);
```

## Set output encoding (Default is ISO-8859-1)
Will alter the encoding of the resultant text - eg. 'UTF-8', 'windows-1252', etc.
```
$rt = new WordTEXT(false, 'desired encoding');
```

## Read docx file and output all the text as an array
```
$text = $rt->readDocument('FILENAME');
```

## Update Notes

Version 1.0.3 - Clearance of some bugs which prevented the script working with some dosc files. Also clearance of php warning messages

Version 1.0.2 - Updated to now work up to at least PHP 8.1

Version 1.0.1 - Original version
