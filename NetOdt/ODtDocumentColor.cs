using System.Drawing;

namespace NetOdt
{
    /// <summary>
    /// Class to set the global colors (foreground and background)
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// The global <see cref="Color"/> of the foreground for text passages of the document
        /// </summary>
        public Color GlobalForegroundColor { get; private set; }

        /// <summary>
        /// The global <see cref="Color"/> of the background for text passages of the document
        /// </summary>
        public Color GlobalBackgroundColor { get; private set; }

        /// <summary>
        /// Set the global foreground and background color for the text passages of the document
        /// </summary>
        /// <param name="foregroundColor">The global <see cref="Color"/> of the foreground for text passages of the document</param>
        /// <param name="backgroundColor">The global <see cref="Color"/> of the background for text passages of the document</param>
        public void SetGlobalColors(Color foregroundColor, Color backgroundColor)
        {
            DocumentChanged = true;
            GlobalForegroundColor = foregroundColor;
            GlobalBackgroundColor = backgroundColor;
        }
    }
}
