# NetOdt
Easy API to create ODT documents in .NET
New feature: Easy API to read or append ODT documents (reading by lines, by tables)
## Simple API usage (short overview)
```csharp
// Needed imports
using NetOdt;
using NetOdt.Enumerations;
using System.Drawing;

// Create document (C# 8.0 syntax)
using var odtDocument = new OdtDocument("C:/MyDocument.odt");

// Set global font for the all text passages for complete document
odtDocument.SetGlobalFont("Liberation Serif", FontSize.Size12);

// Set global colors for the all text passages for the complete document
odtDocument.SetGlobalFont(Color.Red, Color.Transparent);

// Set header and footer
odtDocument.SetHeader("Extend documentation", TextStyle.Center | TextStyle.Bold);
odtDocument.SetFooter("Page 2/5", TextStyle.Right | TextStyle.Italic, Color.Gray, Color.Transparent);

// Append a title
odtDocument.AppendLine("My document", TextStyle.Title, Color.Blue, Color.Black);

// Append a centered bold text
odtDocument.AppendLine("This is a bold text", TextStyle.Center | TextStyle.Bold);

// Append ten empty lines
odtDocument.AppendEmptyLines(lines: 10);

// Append a 3x3 table and fill all cells with "Fill me"
odtDocument.AppendTable(rows: 3, columns: 3, content: "Fill me");

// Append a image with a width of 22.5 cm and a height of 14.1 cm
odtDocument.AppendImage(path: "E:/picture1.jpg", width: 22.5, height: 14.1);

// The automatic dispose call (from the using syntax) do the rest of the work
// (save document, delete temporary folder, free all used resources)
                                                                              
//*****************************************************************************************************
// Read document (C# 8.0 syntax)
using var odtDocument = new OdtDocument("C:/MyDocument.odt", true);

// Every high level paragraph is considered as a line, which can be just text or a table, or a picture.
// So, if you insert a table, the whole table is a big single line.

// Get the number of lines
int lineNumber = odtDocument.GetLineCount();

// Get the raw xml content of the line
for(int i=0;i<lineNumber;i++) { string rawLine = odtDocument.GetRawLine(i); ... }

// Get the plain text of each line
for(int i=0;i<lineNumber;i++) { string plainText = odtDocument.GetPlainLine(i); ... }

// Find the line with specific content
// the second parameter is the index of the line to start the search
int lineIndex =  odtDocument.LineIndexOf("Audio:", 0);

// Get a raw line with a table xml markup
int tableNo = 1;
string rawTable = odtDocument.GetRawTable(tableNo);

// parse table into a standard DataTable structure
// 
int bitwiseOptions = 0;
DataTable table  = odtDocument.GetDataTable(tableNo, bitwiseOptions);
// the content of each cell depends on bitwiseOptions:
// 0 - plain text
// 1 - raw xml text of each cell
// 2 - throw an exception of the number of cells is enough, otherwise any error in structure is ignored

//*****************************************************************************************************
// Append document (C# 8.0 syntax)
using var odtDocument = new OdtDocument("C:/MyDocument.odt", true);

// Append a title
odtDocument.AppendLine("My document", TextStyle.Title, Color.Blue, Color.Black);

// Append a centered bold text
odtDocument.AppendLine("This is a bold text", TextStyle.Center | TextStyle.Bold);

// Append ten empty lines
odtDocument.AppendEmptyLines(lines: 10);

// Append a 3x3 table and fill all cells with "Fill me"
odtDocument.AppendTable(rows: 3, columns: 3, content: "Fill me");

// Append a image with a width of 22.5 cm and a height of 14.1 cm
odtDocument.AppendImage(path: "E:/picture1.jpg", width: 22.5, height: 14.1);

// The automatic dispose call (from the using syntax) do the rest of the work
// (save document, delete temporary folder, free all used resources)


// thats it :-)
```

## Documentation

[All supported styles can find here](./Documentation/styles.md)

[All API calls can find here](./Documentation/api.md)

## Upcoming features
* [Version 2.x](https://github.com/TobiasSekan/NetOdt/milestone/2)
* [Version 3.x](https://github.com/TobiasSekan/NetOdt/milestone/3)

## Bugs report and feature requests
[Please use the GitHub issues tracker](https://github.com/TobiasSekan/NetOdt/issues)

## Goals
* Simple to use API
* Full document API and source-code
* Accept always primitive data-types, `string` and `StringBuilder` to write a value
* Use `enum` when only specific values are except
* Use `EnumFlag` for combinations of `enum` values (bold, italic, ...)
* Use build-in `DataTable` for tables
* No new data-types, use only build-in data-types
* .Net Standard 2.1, C# 8.0
