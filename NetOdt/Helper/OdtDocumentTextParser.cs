using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace NetOdt.Helper
{
    internal class OdtDocumentTextParser
    {
        internal static List<string> TextPartParser(string data)
        {
            int pos = 0;
            int n = data.Length;
            List<string> result = new List<string>();
            while(pos < n)
            {
                char c = data[pos++];
                if(c == '<')
                {
                    int endPos = FindClosingTagEnd(data, pos, n, out var unused, out var unused1);
                    if(endPos < 0)
                    {
                        result.Add(data.Substring(pos-1));
                        break;
                    }
                    result.Add(data.Substring(pos - 1, endPos - pos + 1));
                    pos = endPos;
                }
            }
            return result;
        }

        internal static int FindWordEnd(string data, int pos)
        {
            int n = data.Length;
            while(pos < n)
            {
                char c = data[pos];
                if (c == ' ' || c == '\t' || c=='>')
                {
                    return pos;
                }
                pos++;
            }
            return n;
        }

        internal static int FindStartingTag(string data, int pos, string startTag, int endPos)
        {
            while(pos < endPos)
            {
                int posStart = data.IndexOf(startTag, pos);
                if(posStart < 0 || posStart >= endPos)
                {
                    return -1;
                }
                pos = posStart + startTag.Length;
                if(pos >= endPos)
                {
                    return -1;
                }
                char c = data[pos];
                if(c == ' ' || c=='>')
                {
                    return posStart;
                }
            }
            return -1;
        }
        internal static int FindClosingTagEnd(string data, int pos, int endPos, out int innerStart, out int innerEnd)
        {
            int wordEnd = FindWordEnd(data, pos);
            string tag = data.Substring(pos, wordEnd - pos);
            pos = data.IndexOf('>', wordEnd);
            innerStart = pos +1;
            innerEnd = innerStart;
            if(pos < 0 || pos >= endPos)
            {
                return -1;
            }
            if(data[pos - 1] == '/')
            {
                return pos+1; 
            }
            string startTag = "<" + tag;
            string endTag = "</" + tag + ">";
            int count = 1;
            while(true)
            {
                int posStart = FindStartingTag(data, pos,startTag, endPos);
                int posEnd = data.IndexOf(endTag, pos);
                if(posEnd < 0 || posEnd>=endPos)
                {
                    return -1;
                }
                if(posStart < 0 || posStart > posEnd)
                {
                    count--;
                    innerEnd = posEnd;
                    pos = posEnd + endTag.Length;
                    if(count == 0)
                    {
                        return pos;
                    }
                }
                else
                {
                    count++;
                    pos = posStart + startTag.Length;
                }
            }
        }

        internal static string ExtractPlainText(string s)
        {
            StringBuilder sb = new StringBuilder();
            int n = s.Length;
            int state = 0;
            for(int i = 0; i < n; i++)
            {
                char c = s[i];
                if(c == '<')
                {
                    state = 1;
                }
                else if(c == '>')
                {
                    state = 0;
                }
                else if(state == 0)
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        internal static int GetNextTagPos(string data, string tagName, int pos, int endPos, out int startInnerPos, out int endInnerPos, out int endOuterPos)
        {
            int tagPos = FindStartingTag(data, pos, "<" + tagName, endPos);
            if (tagPos == -1) { 
                endInnerPos = 0;
                endOuterPos = 0;
                startInnerPos = 0;
                return -1; 
            }
            endOuterPos = FindClosingTagEnd(data, tagPos+1, endPos, out startInnerPos, out endInnerPos);
            if( endOuterPos<0)
            {
                return -1;
            }
            return tagPos;
        }

        internal static object[] GetNextTableRow(string s, int pos, int endPos, int options, int columnNo)
        {
            object[] row = new object[columnNo];
            for(int i = 0; i < columnNo; i++)
            {
                int tagPos = GetNextTagPos(s, "table:table-cell", pos, endPos, out var startInnerPos, out var endInnerPos, out var endOuterPos);
                if (tagPos >= 0)
                {
                    row[i] = GetNextTableCell(s, startInnerPos, endInnerPos, options);
                    pos = endOuterPos;
                }
                else if ((options & 2) != 0)
                {
                    throw new Exception("missing cells in the table");
                } else
                {
                    for(;i< columnNo; i++)
                    {
                        row[i] = "";
                    }
                }
            }
            return row;
        }

        internal static string GetNextTableCell(string s, int pos, int endPos, int options)
        {
            string t = s.Substring(pos, endPos - pos);
            if ((options & 1) != 0) 
            {
                return t;
            } 
            return ExtractPlainText(t);
        }

        internal static int CalculateColumnNumber(string s, int pos, int endPos, out int columnNo)
        {
            columnNo = 0;
            string search = "<table:table-column";
            string repeater = "table:number-columns-repeated=\"";
            while(pos < endPos)
            {
                int nxtPos = s.IndexOf(search, pos);
                if (nxtPos < 0 || nxtPos>=endPos)
                {
                    return pos;
                }
                pos = nxtPos + search.Length + 1;
                nxtPos = s.IndexOf(">", pos);
                if(nxtPos < 0 || nxtPos >= endPos)
                {
                    return pos;
                }
                columnNo++;
                string t = s.Substring(pos, nxtPos - pos);
                pos = nxtPos + 1;

                nxtPos = t.IndexOf(repeater);
                if (nxtPos>=0)
                {
                    t = t.Substring(nxtPos + repeater.Length);
                    nxtPos = t.IndexOf("\"");
                    if (nxtPos>=0)
                    {
                        t = t.Substring(0, nxtPos);
                        bool success = int.TryParse(t, out int rpt);
                        if (success && rpt>0)
                        {
                            columnNo += rpt - 1;
                        }
                    }
                }
            }
            return endPos;
        }
        internal static DataTable ExtractDataTable(string s, int options)
        {
            DataTable table = new DataTable("Table");
            int zeroPos = GetNextTagPos(s, "table:table", 0, s.Length, out var pos, out var endPos, out var unused);
            if (zeroPos>=0)
            {
                pos = CalculateColumnNumber(s, pos, endPos, out var columnNo);
                for(int i = 0; i < columnNo; i++)
                {
                    table.Columns.Add(i.ToString(), typeof(string));
                }
                while (pos < endPos)
                {
                    int tagPos = GetNextTagPos(s, "table:table-row", pos, endPos, out var startInnerPos, out var endInnerPos, out var endOuterPos);
                    if (tagPos >= 0)
                    {
                        object[] row = GetNextTableRow(s, startInnerPos, endInnerPos, options, columnNo);
                        table.Rows.Add(row);
                        pos = endOuterPos;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            return table;
        }
    }
}
