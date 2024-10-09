using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities;
using NetOdt;
using NetOdt.Enumerations;
using NUnit.Framework;
using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;

namespace NetOdtTest
{
    [TestFixture]
    public class Program
    {
        [Test]
        public void CompleteTest()
        {
            var uri = new Uri("C:/tmp/testWrite.odt");

            var odtDocument = new OdtDocument(uri);

            odtDocument.SetGlobalFont("Arial", FontSize.Size12);
            odtDocument.SetGlobalColors(Color.White, Color.Black);

            odtDocument.SetHeader("Extend documentation", TextStyle.Center | TextStyle.Bold, "Liberation Serif", FontSize.Size32, Color.Pink, Color.DarkGray);
            odtDocument.SetFooter("Page 2/5", TextStyle.Right | TextStyle.Italic, "Liberation Sans", FontSize.Size20, Color.Blue, Color.LightYellow);

            odtDocument.AppendLine("My Test document", TextStyle.Title, Color.FromArgb(255, 0, 0), Color.FromArgb(0, 126, 126));

            odtDocument.AppendTable(GetTable());

            odtDocument.AppendImage("C:/tmp/picture.jpg", width: 10.5, height: 8.0);

            odtDocument.AppendLine("Unformatted", TextStyle.HeadingLevel01);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 5);
            odtDocument.AppendLine(byte.MaxValue, TextStyle.Center | TextStyle.Bold);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 2);
            odtDocument.AppendLine(uint.MaxValue, TextStyle.Right | TextStyle.Italic);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 1);
            odtDocument.AppendLine(double.NaN, TextStyle.None, "Liberation Serif", FontSize.Size60);

            odtDocument.AppendLine("This is a text", TextStyle.None, "Liberation Sans", FontSize.Size40);

            var content = new StringBuilder();
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            odtDocument.AppendLine(content, TextStyle.Justify);

            odtDocument.AppendTable(GetTable());

            odtDocument.AppendLine("Formatted", TextStyle.HeadingLevel04);

            odtDocument.AppendLine(long.MinValue, TextStyle.Bold | TextStyle.Stroke);
            odtDocument.AppendLine(byte.MaxValue, TextStyle.Italic | TextStyle.Subscript);
            odtDocument.AppendLine(uint.MaxValue, TextStyle.UnderlineSingle | TextStyle.Superscript);
            odtDocument.AppendLine(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle | TextStyle.Justify | TextStyle.Stroke);

            odtDocument.AppendLine("This\n\n\nis\n\n\na\n\n\ntext", TextStyle.Bold | TextStyle.UnderlineSingle | TextStyle.Superscript);

            var contentTwo = new StringBuilder();
            content.Append("This is a text a\n very\n\n\nvery very\n\n\nlong text");
            odtDocument.AppendLine(content, TextStyle.PageBreak | TextStyle.Italic | TextStyle.UnderlineSingle | TextStyle.Subscript);

            odtDocument.AppendLine("sub-sub-sub-sub", TextStyle.Subtitle);

            odtDocument.AppendTable(3, 3, "Fill me");
            odtDocument.AppendTable(3, 3, 0.00);
            odtDocument.Dispose();
            // on Dispose call the ODT document will automatic save and temporary working folder will delete
        }

        internal static DataTable GetTable()
        {
            var table = new DataTable();

            table.Columns.Add("Float", typeof(int));
            table.Columns.Add("Percentage", typeof(float));
            table.Columns.Add("Currency", typeof(decimal));
            table.Columns.Add("Date", typeof(DateTime));
            table.Columns.Add("Time", typeof(TimeSpan));
            table.Columns.Add("Scientific", typeof(double));
            table.Columns.Add("Fraction", typeof(float));
            table.Columns.Add("Boolean", typeof(bool));
            table.Columns.Add("String", typeof(string));
            table.Columns.Add("StringBuilder", typeof(StringBuilder));

            table.Rows.Add(25, 5.5f, 22.27m, DateTime.MinValue, TimeSpan.MinValue, 25.55, 7.8f, false, "Test", new StringBuilder("Test"));
            table.Rows.Add(-25, -5.5f, -22.27m, DateTime.Now, TimeSpan.Zero, -25.55, -7.8f, true, string.Empty, new StringBuilder());
            table.Rows.Add(1000, 5000f, 22000m, DateTime.MaxValue, TimeSpan.MaxValue, 25.55, 7.8f, false, "TestASSASSAAS", new StringBuilder("TestASDSDDASDS"));

            return table;
        }

        [Test]
        public void ReadTest()
        {
            var uri = new Uri("C:/tmp/testRead.odt");
            string message1 = "Audio:";
            string message2 = "http://sponsorschoose.org/";
            DataTable table = GetTestTableData();
            createSimpleTestReadFile(uri, message1 + message2, table);
            DataTable table1 = readSimpleTestFile(uri, message1, out var restMessage);
            if(restMessage != message2)
            {
                throw new Exception("Read cannot obtain the same result, obtained [" + restMessage + "] but expected " + message2);
            }
            compareDataTables(table, table1);
        }

        [Test]
        public void ReadTestComplex()
        {
            var uri = new Uri("C:/tmp/readTest.odt");
            string message1 = "Audio:";
            string message2 = "https://localhost:7001/music/11082024tilbedelse.mp3";
            DataTable table = readSimpleTestFile(uri, message1, out var restMessage);
            if(restMessage.Trim() != message2)
            {
                throw new Exception("Read cannot obtain the same result, obtained [" + restMessage + "] but expected " + message2);
            }
            int cols = table.Columns.Count;
            int rows = table.Rows.Count;
            if (cols!=3)
            {
                throw new Exception("Column number is " + cols + " but expected 3");
            }
            if(rows != 142)
            {
                throw new Exception("Row number is " + rows + " but expected 142");
            }

        }

        private static void createSimpleTestReadFile(Uri uri, string line, DataTable table)
        {
            OdtDocument doc = new OdtDocument(uri);
            doc.AppendLine(line);
            doc.AppendEmptyLines(1);
            doc.AppendTable(table);
            doc.Dispose();
        }

        private static DataTable readSimpleTestFile(Uri uri, string message, out string restMessage)
        {
            OdtDocument doc = new OdtDocument(uri, true);
            int pos = doc.LineIndexOf(message, 0);
            if (pos <0)
            {
                throw new Exception("Not found line with " + message);
            }
            restMessage = doc.GetPlainLine(pos);
            pos = restMessage.IndexOf(message);
            if(pos < 0)
            {
                throw new Exception("Not found " + message + " in " + restMessage);
            }
            restMessage = restMessage.Substring(pos + message.Length);
            DataTable table = doc.GetDataTable(0,0);
            return table;
        }

        private static void compareDataTables(DataTable table1, DataTable table2)
        {
            if (table1 == null || table2 == null)
            {
                throw new Exception("Null table data");
            }
            int cols = table1.Columns.Count;
            if (cols != table2.Columns.Count)
            {
                throw new Exception("Wrong number of columns, expected " + cols + " but it is " + table2.Columns.Count);
            }
            int rows = table1.Rows.Count;
            if(rows != table2.Rows.Count)
            {
                throw new Exception("Wrong number of rows, expected " + rows + " but it is " + table2.Rows.Count);
            }
            for(int col=0; col < cols; col++)
            {
                for(int row = 0; row < rows; row++)
                {
                    string s1 = table1.Rows[row].ItemArray[col].ToString();
                    string s2 = table2.Rows[row].ItemArray[col].ToString();
                    if(s1 != s2)
                    {
                        throw new Exception(s1 + " does not match " + s2 + " at " + row + ","+ col);
                    }
                }
            }
        }
        private static DataTable GetTestTableData()
        {
            DataTable table = new DataTable("Table");
            table.Columns.Add("Time", typeof(string));
            table.Columns.Add("nb", typeof(string));
            table.Columns.Add("en", typeof(string));
            object[] headers = new object[3];
            headers[0] = "";
            headers[1] = "nb";
            headers[2] = "en";
            table.Rows.Add(headers);
            object[] data1 = new object[3];
            data1[0] = "01:00";
            data1[1] = "Du rufst mich raus auf's weite Wasser";
            data1[2] = "You called me from deep waters";
            table.Rows.Add(data1);
            object[] data2 = new object[3];
            data2[0] = "02:00.23";
            data2[1] = "wo Füße nicht... mehr sicher stehn";
            data2[2] = "Where no foot, me safely placed";
            table.Rows.Add(data2);
            return table;
        }

    }
}
