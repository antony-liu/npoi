/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using NPOI.Util;
using System;
using System.IO;
using System.Reflection;

namespace NPOI.OpenXml4Net.Util
{
    public class ZipSecureFile : ZipFile
    {
        private static POILogger _logger =
                POILogFactory.GetLogger(typeof(ZipSecureFile));
        private static double MIN_INFLATE_RATIO = 0.01d;
        private static long MAX_ENTRY_SIZE = 0xFFFFFFFFL;
        // don't alert for expanded sizes smaller than 100k
        private static long GRACE_ENTRY_SIZE = 100 * 1024;

        // The default maximum size of extracted text 
        private static long MAX_TEXT_SIZE = 10 * 1024 * 1024;
        /**
         * Sets the ratio between de- and inflated bytes to detect zipbomb.
         * It defaults to 1% (= 0.01d), i.e. when the compression is better than
         * 1% for any given read package part, the parsing will fail indicating a 
         * Zip-Bomb.
         *
         * @param ratio the ratio between de- and inflated bytes to detect zipbomb
         */
        public static double MinInflateRatio
        {
            get
            {
                return MIN_INFLATE_RATIO;
            }
            set
            {
                MIN_INFLATE_RATIO = value;
            }
        }

        /**
         * Sets the maximum file size of a single zip entry. It defaults to 4GB,
         * i.e. the 32-bit zip format maximum.
         * 
         * This can be used to limit memory consumption and protect against 
         * security vulnerabilities when documents are provided by users.
         *
         * @param maxEntrySize the max. file size of a single zip entry
         */
        public static long MaxEntrySize
        {
            get
            {
                return MAX_ENTRY_SIZE;
            }
            set
            {
                if (value < 0 || value > 0xFFFFFFFFl)
                {
                    throw new ArgumentException("Max entry size is bounded [0-4GB].");
                }
                MAX_ENTRY_SIZE = value;
            }
        }

        /**
         * get or set the maximum number of characters of text that are
         * extracted before an exception is thrown during extracting
         * text from documents.
         * 
         * This can be used to limit memory consumption and protect against 
         * security vulnerabilities when documents are provided by users.
         *
         * @param maxTextSize the max. file size of a single zip entry
         */
        public static long MaxTextSize
        {
            get
            {
                return MAX_TEXT_SIZE;
            }
            set
            {
                if (value < 0 || value > 0xFFFFFFFFL)
                {     // don't use MAX_ENTRY_SIZE here!
                    throw new ArgumentException("Max text size is bounded [0-4GB], but had " + value);
                }
                MAX_TEXT_SIZE = value;
            }
        }

        public ZipSecureFile(FileStream file, int mode)
            : base(file)
        {

        }

        public ZipSecureFile(FileStream file)
            : base(file)
        {

        }

        public ZipSecureFile(String name)
                : base(name)
        {

        }

        /**
         * Returns an input stream for reading the contents of the specified
         * zip file entry.
         *
         * <p> Closing this ZIP file will, in turn, close all input
         * streams that have been returned by invocations of this method.
         *
         * @param entry the zip file entry
         * @return the input stream for reading the contents of the specified
         * zip file entry.
         * @throws ZipException if a ZIP format error has occurred
         * @throws IOException if an I/O error has occurred
         * @throws IllegalStateException if the zip file has been closed
         */
        public new Stream GetInputStream(ZipEntry entry)
        {
            Stream zipIS = base.GetInputStream(entry);
            return AddThreshold(zipIS);
        }

        public static ThresholdInputStream AddThreshold(Stream zipIS)
        {

            ThresholdInputStream newInner = null;
            if (zipIS is InflaterInputStream)
            {
                //replace inner stream of zipIS by using a ThresholdInputStream instance??
                try
                {
                    FieldInfo f = typeof(InflaterInputStream).GetField("baseInputStream ", BindingFlags.NonPublic | BindingFlags.GetField | BindingFlags.Instance);
                    //f.SetAccessible(true);
                    Stream oldInner = (Stream)f.GetValue(zipIS);
                    newInner = new ThresholdInputStream(oldInner, null);
                    f.SetValue(zipIS, newInner);
                    //f.Set(zipIS, newInner);
                } catch (Exception ex) {
                    _logger.Log(POILogger.WARN, "SecurityManager doesn't allow manipulation via reflection for zipbomb detection - continue with original input stream", ex);
                    newInner = null;
                }
            } else {
                // the inner stream is a ZipFileInputStream, i.e. the data wasn't compressed
                newInner = null;
            }

            return new ThresholdInputStream(zipIS, newInner);
        }

        public class ThresholdInputStream : Stream
        {
            long counter = 0;
            ThresholdInputStream cis;
            Stream input;

            public override bool CanRead => true;

            public override bool CanSeek => false;

            public override bool CanWrite => false;

            public override long Length => input.Length;

            public override long Position { get => input.Position; set => input.Position = value; }

            public ThresholdInputStream(Stream is1, ThresholdInputStream cis)
            {
                this.input = is1;
                this.cis = cis;
            }

            public int Read()
            {
                int b = this.input.ReadByte();
                if (b > -1) Advance(1);
                return b;
            }

            public override int Read(byte[] b, int off, int len)
            {
                int cnt = input.Read(b, off, len);
                if (cnt > -1) Advance(cnt);
                return cnt;
            }

            public long Skip(long n)
            {

                counter = 0;
                return input.Seek(n, SeekOrigin.Current);
            }

            public void Reset()
            {
                counter = 0;
                input.Seek(0, SeekOrigin.Begin);
            }

            public void Advance(int advance)
            {
                counter += advance;

                // check the file size first, in case we are working on uncompressed streams
                if (counter > MAX_ENTRY_SIZE)
                {
                    throw new IOException("Zip bomb detected! The file would exceed the max size of the expanded data in the zip-file. "
                            + "This may indicates that the file is used to inflate memory usage and thus could pose a security risk. "
                            + "You can adjust this limit via ZipSecureFile.setMaxEntrySize() if you need to work with files which are very large. "
                            + "Counter: " + counter + ", cis.counter: " + (cis == null ? 0 : cis.counter)
                            + "Limits: MAX_ENTRY_SIZE: " + MAX_ENTRY_SIZE);
                }

                // no expanded size?
                if (cis == null)
                {
                    return;
                }

                // don't alert for small expanded size
                if (counter <= GRACE_ENTRY_SIZE)
                {
                    return;
                }

                double ratio = (double)cis.counter / (double)counter;
                if (ratio >= MIN_INFLATE_RATIO)
                {
                    return;
                }

                // one of the limits was reached, report it
                throw new IOException("Zip bomb detected! The file would exceed the max. ratio of compressed file size to the size of the expanded data. "
                        + "This may indicate that the file is used to inflate memory usage and thus could pose a security risk. "
                        + "You can adjust this limit via ZipSecureFile.setMinInflateRatio() if you need to work with files which exceed this limit. "
                        + "Counter: " + counter + ", cis.counter: " + cis.counter + ", ratio: " + (((double)cis.counter) / counter)
                        + "Limits: MIN_INFLATE_RATIO: " + MIN_INFLATE_RATIO);
            }

            public ZipEntry GetNextEntry()
            {

                if (!(input is ZipInputStream)) {
                    throw new NotSupportedException("underlying stream is not a ZipInputStream");
                }
                counter = 0;
                return ((ZipInputStream)input).GetNextEntry();
            }

            public void CloseEntry()
            {

                if (!(input is ZipInputStream)) {
                    throw new NotSupportedException("underlying stream is not a ZipInputStream");
                }
                counter = 0;
                ((ZipInputStream)input).CloseEntry();
            }

            public void Unread(int b)
            {

                if (!(input is PushbackInputStream)) {
                    throw new NotSupportedException("underlying stream is not a PushbackInputStream");
                }
                if (--counter < 0) counter = 0;
                ((PushbackInputStream)input).Unread(b);
            }

            public void Unread(byte[] b, int off, int len)
            {

                if (!(input is PushbackInputStream)) {
                    throw new NotSupportedException("underlying stream is not a PushbackInputStream");
                }
                counter -= len;
                if (--counter < 0) counter = 0;
                ((PushbackInputStream)input).Unread(b, off, len);
            }

            public int Available()
            {
                return (int)(input.Length - input.Position);
                //return input.Available();
            }

            public bool MarkSupported()
            {
                //return input.MarkSupported();
                return true;
            }

            public void Mark(int readlimit)
            {
                //input.Mark(readlimit);
            }

            public override void Flush()
            {
                throw new NotImplementedException();
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                throw new NotImplementedException();
            }

            public override void SetLength(long value)
            {
                throw new NotImplementedException();
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                throw new NotImplementedException();
            }

            public override void Close()
            {
                if (input == null)
                    return;
                input.Close();
                input = null;
            }
        }

    }
}
