using System;
using System.Collections;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using System.Collections.Generic;
using NPOI.Util;

namespace NPOI.OpenXml4Net.Util
{
    /**
     * An Interface to make getting the different bits
     *  of a Zip File easy.
     * Allows you to get at the ZipEntries, without
     *  needing to worry about ZipFile vs ZipInputStream
     *  being annoyingly very different.
     */
    public interface IZipEntrySource : ICloseable
    {
        /**
         * Returns an Enumeration of all the Entries
         */
        IEnumerator<ZipEntry> Entries { get; }

        /**
         * Returns an InputStream of the decompressed 
         *  data that makes up the entry
         */
        Stream GetInputStream(ZipEntry entry);

        /**
         * Indicates we are done with reading, and 
         *  resources may be freed
         */
        void Close();

        /**
         * Has close been called already?
         */
        bool IsClosed { get; }
    }

    public static class ZipEntrySourceExtension
    { 
        public static ZipEntry GetEntry(this IZipEntrySource zipArchive, 
            string path)
        {
            IEnumerator entries = zipArchive.Entries;
            while(entries.MoveNext())
            {
                ZipEntry entry = (ZipEntry)entries.Current;
                if(entry.Name.ToLower().Equals(path.ToLower()))
                {
                    return entry;
                }
            }
            return null;
        }
        /// <summary>
        /// true if and only if this enumeration object contains at least 
        /// one more element to provide; false otherwise.
        /// </summary>
        /// <param name="zipArchive"></param>
        /// <returns></returns>
        public static bool HasMoreElements(this IZipEntrySource zipArchive)
        {
            int numEntries = 0;
            IEnumerator entries = zipArchive.Entries;
            while(entries.MoveNext())
            {
                numEntries++;
            }
            return numEntries > 0;
        }
    }
}
