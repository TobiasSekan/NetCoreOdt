using NetOdt.Helper;
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Reflection;

namespace NetOdt
{
    public sealed partial class OdtDocument
    {
        private int CacheTextCount = -1;
        private List<string> CacheLines = null;

        private void MakeLineCache()
        {
            if(CacheLines == null || CacheTextCount != TextContent.Length)
            {
                CacheLines = OdtDocumentTextParser.TextPartParser(TextContent.ToString());
                CacheTextCount = TextContent.Length;
            }
        }

        /// <summary>
        ///    Find the first line which contains a string and starting with a certain line
        /// </summary>
        /// <param name = "searchString" > The string for the search</param>
        /// <param name = "startLine">the start line for the search, indexed from 0</param>
        public int LineIndexOf(string searchString, int startLine)
        {
            MakeLineCache();
            int n = CacheLines.Count;
            for(int i = startLine; i < n; i++)
            {
                string line = CacheLines[i];
                if (line.IndexOf(searchString) >=0 )
                {
                    return i;
                }
            }
            return -1;  
        }

        /// <summary>
        ///    Get the current number of lines in the document
        /// </summary>
        public int GetLineCount()
        {
            MakeLineCache();
            return CacheLines.Count;
        }

        /// <summary>
        ///    Get line as a raw xml string
        /// </summary>
        /// <param name = "lineNo" > The sequential number of the line starting with 0.</param>
        public string GetRawLine(int lineNo)
        {
            MakeLineCache();
            return CacheLines[lineNo];
        }

        /// <summary>
        ///    Get a plain line as a text without xml markup
        /// </summary>
        /// <param name = "lineNo" > The sequential number of the line starting with 0.</param>
        public string GetPlainLine(int lineNo)
        {
            var s = GetRawLine(lineNo);
            return OdtDocumentTextParser.ExtractPlainText(s);
        }

        /// <summary>
        ///    Get table as a raw string
        /// </summary>
        /// <param name = "tableNo" > The sequential number of the table starting with 0.</param>
         public string GetRawTable(int tableNo)
        {
            MakeLineCache();
            int n = CacheLines.Count;
            for(int i = 0;i<n;i++)
            {
                if(CacheLines[i].StartsWith("<table:table"))
                {
                    if (tableNo == 0)
                    {
                        return CacheLines[i];
                    } else
                    {
                        tableNo--;
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// Get table in the form of DataTable
        /// </summary>
        /// <param name = "tableNo" > The sequential number of the table starting with 0.</param>
        /// <param name = "bitwiseOptions" > 0 - plain data, 1-full cell data, 2-throw error if missing a cell, 4-through</param>
        public DataTable GetDataTable(int tableNo, int bitwiseOptions)
        {
            var s = GetRawTable(tableNo);
            return OdtDocumentTextParser.ExtractDataTable(s, bitwiseOptions);
        }
    }
}
