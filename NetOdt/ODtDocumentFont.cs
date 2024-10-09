using NetOdt.Enumerations;
using NetOdt.Helper;
using System;

namespace NetOdt
{
    /// <summary>
    /// Class to set the global font (name and size)
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// The font name for complete document
        /// </summary>
        public string GlobalFontName { get; private set; }

        /// <summary>
        /// The font size for complete document
        /// </summary>
        public double GlobalFontSize { get; private set; }

        /// <summary>
        /// Set the font (name) and font size for the complete document
        /// </summary>
        /// <param name="fontName">The name of the font</param>
        /// <param name="fontSize">The font size representation for the font</param>
        public void SetGlobalFont(string fontName, FontSize fontSize)
            => SetGlobalFont(fontName, FontHelper.GetFontSize(fontSize));

        /// <summary>
        /// Set the font (name) and font size for the complete document
        /// </summary>
        /// <param name="fontName">The name of the font</param>
        /// <param name="fontSize">The size of the font</param>
        public void SetGlobalFont(string fontName, double fontSize)
        {
            if(string.IsNullOrWhiteSpace(fontName))
            {
                throw new ArgumentOutOfRangeException(nameof(fontName), fontName, "The font name can't be a empty string");
            }

            if(fontSize < 1.0)
            {
                throw new ArgumentOutOfRangeException(nameof(fontSize), fontSize, "The font size can be smaller as 1.0");
            }
            DocumentChanged = true;
            GlobalFontName = fontName;
            GlobalFontSize = fontSize;
        }
    }
}
