using NetOdt.Class;
using NetOdt.Enumerations;
using NetOdt.Helper;
using System;
using System.Drawing;
using System.Text;

namespace NetOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// Append a content to the document (use standard style and global font and color)
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        public void AppendLine<TValue>(in TValue content)
            where TValue : notnull
            => AppendLine(content, TryToAddStyle(TextStyle.None, StyleFamily.Paragraph, OfficeValueType.String, GlobalFontSize, GlobalFontName, GlobalForegroundColor, GlobalBackgroundColor));

        /// <summary>
        /// Append a content to the document with the given style (use global font and color)
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        /// <param name="textStyle">The style for the content</param>
        public void AppendLine<TValue>(in TValue content, in TextStyle textStyle)
            where TValue : notnull
            => AppendLine(content, TryToAddStyle(textStyle, StyleFamily.Paragraph, OfficeValueType.String, GlobalFontSize, GlobalFontName, GlobalForegroundColor, GlobalBackgroundColor));

        /// <summary>
        /// Append a content to the document with the given style and color (use global font)
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        /// <param name="textStyle">The style for the content</param>
        /// <param name="foreground">The <see cref="Color"/> of foreground for the content</param>
        /// <param name="background">The <see cref="Color"/> of background for the content</param>
        public void AppendLine<TValue>(in TValue content, in TextStyle textStyle, in Color foreground, in Color background)
            where TValue : notnull
            => AppendLine(content, TryToAddStyle(textStyle, StyleFamily.Paragraph, OfficeValueType.String, GlobalFontSize, GlobalFontName, foreground, background));

        /// <summary>
        /// Append a content to the document with the given style and font (use global colors)
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        /// <param name="textStyle">The style of the value</param>
        /// <param name="fontName">The font (name) for the content</param>
        /// <param name="fontSize">The font size for the content</param>
        public void AppendLine<TValue>(in TValue content, in TextStyle textStyle, in string fontName, in FontSize fontSize)
            where TValue : notnull
            => AppendLine(content, TryToAddStyle(textStyle, StyleFamily.Paragraph, OfficeValueType.String, FontHelper.GetFontSize(fontSize), fontName, GlobalForegroundColor, GlobalBackgroundColor));

        /// <summary>
        /// Append a content to the document with the given style and font (use global colors)
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        /// <param name="textStyle">The style of the value</param>
        /// <param name="fontName">The font (name) for the content</param>
        /// <param name="fontSize">The font size for the content</param>
        public void AppendLine<TValue>(in TValue content, in TextStyle textStyle, in string fontName, in double fontSize)
            where TValue : notnull
            => AppendLine(content, TryToAddStyle(textStyle, StyleFamily.Paragraph, OfficeValueType.String, fontSize, fontName, GlobalForegroundColor, GlobalBackgroundColor));

        /// <summary>
        /// Append a content to the document with the given style, font and color
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        /// <param name="textStyle">The style of the value</param>
        /// <param name="fontName">The font (name) for the content</param>
        /// <param name="fontSize">The font size for the content</param>
        /// <param name="foreground">The <see cref="Color"/> of foreground for the content</param>
        /// <param name="background">The <see cref="Color"/> of background for the content</param>
        public void AppendLine<TValue>(in TValue content,
                                       in TextStyle textStyle,
                                       in string fontName,
                                       in FontSize fontSize,
                                       in Color foreground,
                                       in Color background)
            where TValue : notnull
            => AppendLine(content, TryToAddStyle(textStyle, StyleFamily.Paragraph, OfficeValueType.String, FontHelper.GetFontSize(fontSize), fontName, foreground, background));

        /// <summary>
        /// Append a content to the document with the given style, font and color
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        /// <param name="textStyle">The text style of the content</param>
        /// <param name="fontName">The font (name) for the content</param>
        /// <param name="fontSize">The font size for the content</param>
        /// <param name="foreground">The <see cref="Color"/> of foreground for the content</param>
        /// <param name="background">The <see cref="Color"/> of background for the content</param>
        public void AppendLine<TValue>(in TValue content,
                                       in TextStyle textStyle,
                                       in string fontName,
                                       in double fontSize,
                                       in Color foreground,
                                       in Color background)
            where TValue : notnull
        {
            AppendLine(content, TryToAddStyle(textStyle, StyleFamily.Paragraph, OfficeValueType.String, fontSize, fontName, foreground, background));
        }

        /// <summary>
        /// Append a content to the document with the given style, font and color
        /// </summary>
        /// <param name="content">The content to write into the document</param>
        /// <param name="style">The style for the content</param>
        internal void AppendLine<TValue>(in TValue content, in Style style) where TValue : notnull
        {
            DocumentChanged = true;
            if(content is StringBuilder stringBuilder)
            {
                if(stringBuilder.Length == 0 || !StringBuilderHelper.ContainsLineBreaks(stringBuilder))
                {
                    AppendXmlTextStart(style);
                    AppendValue(stringBuilder);
                    AppendXmlTextEnd(style);
                    return;
                }

                foreach(var contentBlock in content.ToString().Split('\n'))
                {
                    AppendXmlTextStart(style);
                    AppendValue(contentBlock);
                    AppendXmlTextEnd(style);
                }

                return;
            }

            if(content is string text)
            {
                if(text.Length == 0 || !text.Contains("\n"))
                {
                    AppendXmlTextStart(style);
                    AppendValue(text);
                    AppendXmlTextEnd(style);
                    return;
                }

                foreach(var textBlock in text.Split('\n'))
                {
                    AppendXmlTextStart(style);
                    AppendValue(textBlock);
                    AppendXmlTextEnd(style);
                }

                return;
            }

            AppendXmlTextStart(style);
            AppendValue(content);
            AppendXmlTextEnd(style);
        }

        /// <summary>
        /// Append the given count of empty lines
        /// </summary>
        /// <param name="countOfEmptyLines">The count of empty lines to write</param>
        public void AppendEmptyLines(in int countOfEmptyLines)
        {
            for(var count = 0; count < countOfEmptyLines; count++)
            {
                AppendLine(string.Empty, TextStyle.None);
            }
        }

        /// <summary>
        /// Append a XML start element for a text passage
        /// </summary>
        /// <param name="style">The style for the text passage</param>
        internal void AppendXmlTextStart(in Style style)
        {
            TextContent.Append("<");

            switch(style.TextStyle)
            {
                case TextStyle.HeadingLevel01:
                case TextStyle.HeadingLevel02:
                case TextStyle.HeadingLevel03:
                case TextStyle.HeadingLevel04:
                case TextStyle.HeadingLevel05:
                case TextStyle.HeadingLevel06:
                case TextStyle.HeadingLevel07:
                case TextStyle.HeadingLevel08:
                case TextStyle.HeadingLevel09:
                case TextStyle.HeadingLevel10:
                    TextContent.Append("text:h");
                    TextContent.Append($" text:style-name=\"{style.Name}\"");
                    TextContent.Append($" text:outline-level=\"{style.OutlineLevel}\"");
                    break;

                default:
                    TextContent.Append("text:p");
                    TextContent.Append($" text:style-name=\"{style.Name}\"");
                    break;
            }

            TextContent.Append(">");
        }

        /// <summary>
        /// Append the value of a text passage
        /// </summary>
        /// <typeparam name="TValue">The <see cref="Type"/> of the value</typeparam>
        /// <param name="value">The value for the text passage</param>
        internal void AppendValue<TValue>(in TValue value)
            where TValue : notnull
        {
            switch(value)
            {
                case bool boolValue:
                    TextContent.Append(boolValue);
                    break;

                case byte byteValue:
                    TextContent.Append(byteValue);
                    break;

                case sbyte sbyteValue:
                    TextContent.Append(sbyteValue);
                    break;

                case char charValue:
                    TextContent.Append(charValue);
                    break;

                case short shortValue:
                    TextContent.Append(shortValue);
                    break;

                case ushort ushortValue:
                    TextContent.Append(ushortValue);
                    break;

                case int intValue:
                    TextContent.Append(intValue);
                    break;

                case uint uintValue:
                    TextContent.Append(uintValue);
                    break;

                case long longValue:
                    TextContent.Append(longValue);
                    break;

                case ulong ulongValue:
                    TextContent.Append(ulongValue);
                    break;

                case float floatValue:
                    TextContent.Append(floatValue);
                    break;

                case double doubleValue:
                    TextContent.Append(doubleValue);
                    break;

                case decimal decimalValue:
                    TextContent.Append(decimalValue);
                    break;

                case string stringValue:
                    TextContent.Append(stringValue);
                    break;

                // there is no overload that takes a StringBuilder in .NET Standard 2.0
                // but in .Net Standard 2.1 "StringBuilder" have a overload for "StringBuilder"
                //
                case StringBuilder stringBuilder:
                    TextContent.Append(stringBuilder);
                    break;

                // Pattern matching doesn't support "ReadOnlySpan<char>",
                // but in .Net Standard 2.1 "StringBuilder" have a overload for "ReadOnlySpan<char>"
                //
                // case ReadOnlySpan<char> readOnlySpan:
                //     TextContent.Append(readOnlySpan);
                //     break;

                default:
                    TextContent.Append(value);
                    break;
            }
        }

        /// <summary>
        /// Append a XML end element for a text passage
        /// </summary>
        /// <param name="style">The style of this text passage</param>
        internal void AppendXmlTextEnd(in Style style)
        {
            switch(style.TextStyle)
            {
                case TextStyle.HeadingLevel01:
                case TextStyle.HeadingLevel02:
                case TextStyle.HeadingLevel03:
                case TextStyle.HeadingLevel04:
                case TextStyle.HeadingLevel05:
                case TextStyle.HeadingLevel06:
                case TextStyle.HeadingLevel07:
                case TextStyle.HeadingLevel08:
                case TextStyle.HeadingLevel09:
                case TextStyle.HeadingLevel10:
                    TextContent.Append("</text:h>");
                    break;

                default:
                    TextContent.Append("</text:p>");
                    break;
            }
        }
    }
}
