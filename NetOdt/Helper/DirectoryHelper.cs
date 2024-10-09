using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with <see cref="Directory"/>
    /// </summary>
    internal static class DirectoryHelper
    {
        /// <summary>
        /// Creates all directories and subdirectories in the specified path unless they already exist.
        /// </summary>
        /// <param name="uriLeft">The left part of the complete path</param>
        /// <param name="pathRight">The right part of the complete path</param>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void CreateDirectory(Uri uriLeft, string pathRight)
            => Directory.CreateDirectory(Path.Combine(uriLeft.AbsolutePath, pathRight));

        /// <summary>
        /// Deletes the specified directory
        /// </summary>
        /// <param name="uriLeft">The left part for the complete path</param>
        /// <param name="pathRight">The right part for the complete path</param>
        internal static void Delete(Uri uriLeft, string pathRight)
        {
            try
            {
                string path = Path.Combine(uriLeft.AbsolutePath, pathRight);
                if(Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch(Exception ex)
            {
            }
        }
            
    }
}
