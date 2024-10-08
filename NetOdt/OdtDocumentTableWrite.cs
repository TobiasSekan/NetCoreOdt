using NetOdt.Enumerations;
using NetOdt.Helper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace NetOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument : IDisposable
    {
        /// <summary>
        /// Append an empty unformatted table with the given row and cell count into the document
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        public void AppendTable(in int rowCount, in int columnCount)
        {
            DocumentChanged = true;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(string.Empty),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(string.Empty)}");
                    TextContent.Append(">");

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table with the given row and cell count into the document and fill each cell with the given value
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        /// <param name="value">The value for each cell</param>
        public void AppendTable(in int rowCount, in int columnCount, in ValueType value)
        {
            DocumentChanged = true;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(value),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(value)}");
                    TextContent.Append(">");

                    AppendLine(StyleHelper.GetFormatedValue(value), style);

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table with the given row and cell count into the document and fill each cell with the given text
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        /// <param name="text">The text for each cell</param>
        public void AppendTable(in int rowCount, in int columnCount, in string text)
        {
            DocumentChanged = true;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(text),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(text)}");
                    TextContent.Append(">");

                    AppendLine(StyleHelper.GetFormatedValue(text), style);

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table with the given row and cell count into the document and fill each cell with the given content
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        /// <param name="content">The content for each cell</param>
        public void AppendTable(in int rowCount, in int columnCount, in StringBuilder content)
        {
            DocumentChanged = true;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(content),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(content)}");
                    TextContent.Append(">");


                    AppendLine(StyleHelper.GetFormatedValue(content), style);

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given data from the value array
        /// </summary>
        /// <param name="valueArray">The array that contains the values for the table</param>
        public void AppendTable(in IEnumerable<IEnumerable<ValueType>> valueArray)
        {
            DocumentChanged = true ;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{valueArray.Count()}\"/>");

            var rowNumber = 0;

            foreach(var dataRow in valueArray)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.Count();
                var columnNumber = 0;

                foreach(var column in dataRow)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(column),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(column)}");
                    TextContent.Append(">");

                    AppendLine(StyleHelper.GetFormatedValue(column), style);

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given data from the string array
        /// </summary>
        /// <param name="stringArray">The array that contains the strings for the table</param>
        public void AppendTable(in IEnumerable<IEnumerable<string>> stringArray)
        {
            DocumentChanged = true;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{stringArray.Count()}\"/>");

            var rowNumber = 0;

            foreach(var dataRow in stringArray)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.Count();
                var columnNumber = 0;

                foreach(var column in dataRow)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(column),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(column)}");
                    TextContent.Append(">");

                    AppendLine(StyleHelper.GetFormatedValue(column), style);

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given content from the content array
        /// </summary>
        /// <param name="contentArray">The array that contains the contents for the table</param>
        public void AppendTable(in IEnumerable<IEnumerable<StringBuilder>> contentArray)
        {
            DocumentChanged = true;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{contentArray.Count()}\"/>");

            var rowNumber = 0;

            foreach(var dataRow in contentArray)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.Count();
                var columnNumber = 0;

                foreach(var column in dataRow)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(column),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(column)}");
                    TextContent.Append(">");

                    AppendLine(StyleHelper.GetFormatedValue(column), style);

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given data from the <see cref="DataTable"/>
        /// </summary>
        /// <param name="dataTable">The <see cref="DataTable"/> that contains the data for the table</param>
        public void AppendTable(in DataTable dataTable)
        {
            DocumentChanged = true;
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{dataTable.Columns.Count}\"/>");

            var rowNumber = 0;

            foreach(DataRow? dataRow in dataTable.Rows)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.ItemArray.Length;
                var columnNumber = 0;

                foreach(var column in dataRow.ItemArray)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    var style = TryToAddStyle(TextStyle.None,
                                              StyleFamily.TableCell,
                                              StyleHelper.GetOfficeValueType(column),
                                              GlobalFontSize,
                                              GlobalFontName,
                                              GlobalForegroundColor,
                                              GlobalBackgroundColor,
                                              StyleHelper.GetTableCellBorderStyle(rowNumber, columnNumber, columnCount));

                    TextContent.Append("<");
                    TextContent.Append("table:table-cell");
                    TextContent.Append($" table:style-name=\"{style.Name}\"");
                    TextContent.Append($" {StyleHelper.GetOfficeValues(column)}");
                    TextContent.Append(">");

                    AppendLine(StyleHelper.GetFormatedValue(column), style);

                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }
    }
}
