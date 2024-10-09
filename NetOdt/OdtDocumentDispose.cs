using NetOdt.Constants;
using System;
using System.IO;
using System.Xml;

namespace NetOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument : IDisposable
    {
        /// <summary>
        /// Save the document (override when existing), delete folder under the <see cref="TempWorkingUri"/> and free all resources
        /// </summary>
        public void Dispose()
            => Dispose(overrideExistingFile: true);

        /// <summary>
        /// Save the document, delete the folder under the <see cref="TempWorkingUri"/> and free all resources
        /// </summary>
        public void Dispose(in bool overrideExistingFile)
        {
            if (DocumentChanged)
            {
                Save(overrideExistingFile);
            }

            Directory.Delete(FolderResource.TemporaryRootFolderPath, true);

            BeforeStyleContent.Clear();
            StyleContent.Clear();
            AfterStyleContent.Clear();
            TextContent.Clear();
            AfterTextContent.Clear();

            ContentFile = new XmlDocument();

            //TempWorkingUri = null;
            //ContentFileUri = null;
        }
    }
}
