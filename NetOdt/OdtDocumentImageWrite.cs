using NetOdt.Constants;
using NetOdt.Helper;
using System;
using System.Globalization;
using System.IO;

namespace NetOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// Append a image to the document
        /// </summary>
        /// <param name="imagePath">The full path to image</param>
        /// <param name="width">The width of the picture in centimeter (cm)</param>
        /// <param name="height">The height of the picture in centimeter (cm)</param>
        public void AppendImage(string imagePath, double width, double height)
        {
            DocumentChanged = true;
            PictureCount++;

            var pictureExtension = Path.GetExtension(imagePath);
            var mineType         = FileHelper.GetMineType(imagePath);
            var picturePath      = $"{FolderResource.PictureFolderName}/{PictureCount}{pictureExtension}";

            FileHelper.Copy(imagePath, UriHelper.Combine(TempWorkingUri, picturePath));

            TextContent.Append($"<text:p text:style-name=\"Standard\">");
            TextContent.Append($"<draw:frame draw:style-name=\"fr1\" draw:name=\"Picture{PictureCount}\" text:anchor-type=\"paragraph\" svg:width=\"{width.ToString(CultureInfo.InvariantCulture)}cm\" svg:height=\"{height.ToString(CultureInfo.InvariantCulture)}cm\" draw:z-index=\"0\">");
            TextContent.Append($"<draw:image xlink:href=\"{picturePath}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\" loext:mime-type=\"{mineType}\"/>");
            TextContent.Append($"</draw:frame>");
            TextContent.Append($"</text:p>");

            ManifestContent.Append($"<manifest:file-entry manifest:full-path=\"{picturePath}\" manifest:media-type=\"{mineType}\"/>");
        }

        /// <summary>
        /// Append a image to the document
        /// </summary>
        /// <param name="imageUri">The <see cref="Uri"/> of the image</param>
        /// <param name="width">The width of the picture in centimeter (cm)</param>
        /// <param name="height">The height of the picture in centimeter (cm)</param>
        public void AppendImage(Uri imageUri, double width, double height)
        {
            DocumentChanged = true;
            var pictureExtension = PathHelper.GetExtension(imageUri);
            var mineType         = FileHelper.GetMineType(imageUri.AbsolutePath);
            var picturePath      = $"{FolderResource.PictureFolderName}/{PictureCount}{pictureExtension}";

            FileHelper.Copy(imageUri, UriHelper.Combine(TempWorkingUri, picturePath));

            TextContent.Append($"<text:p text:style-name=\"Standard\">");
            TextContent.Append($"<draw:frame draw:style-name=\"fr1\" draw:name=\"Picture{PictureCount}\" text:anchor-type=\"paragraph\" svg:width=\"{width}cm\" svg:height=\"{height}cm\" draw:z-index=\"0\">");
            TextContent.Append($"<draw:image xlink:href=\"{picturePath}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\" loext:mime-type=\"{mineType}\"/>");
            TextContent.Append($"</draw:frame>");
            TextContent.Append($"</text:p>");

            ManifestContent.Append($"<manifest:file-entry manifest:full-path=\"{picturePath}\" manifest:media-type=\"{mineType}\"/>");
        }
    }
}
