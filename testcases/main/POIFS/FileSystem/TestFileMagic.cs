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


using NPOI.POIFS.FileSystem;
using NPOI.Util;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.POIFS.FileSystem
{
    [TestFixture]
    public class TestFileMagic
    {
        [Test]
        public void TestFileMagicValueOf()
        {
            ClassicAssert.AreEqual(FileMagic.XML, FileMagicContainer.ValueOf("XML"));
            ClassicAssert.AreEqual(FileMagic.XML, FileMagicContainer.ValueOf(Encoding.UTF8.GetBytes("<?xml")));

            ClassicAssert.AreEqual(FileMagic.HTML, FileMagicContainer.ValueOf("HTML"));
            ClassicAssert.AreEqual(FileMagic.HTML, FileMagicContainer.ValueOf(Encoding.UTF8.GetBytes("<!DOCTYP")));
            ClassicAssert.AreEqual(FileMagic.HTML, FileMagicContainer.ValueOf(Encoding.UTF8.GetBytes("<!DOCTYPE")));
            ClassicAssert.AreEqual(FileMagic.HTML, FileMagicContainer.ValueOf(Encoding.UTF8.GetBytes("<html")));

            try
            {
                FileMagicContainer.ValueOf("some string");
                ClassicAssert.Fail("Should catch exception here");
            }
            catch(ArgumentException e)
            {
                // expected here
            }
        }

        [Test]
        public void TestFileMagicFile()
        {
            ClassicAssert.AreEqual(FileMagic.OLE2, FileMagicContainer.ValueOf(POIDataSamples.GetSpreadSheetInstance().GetFile("SampleSS.xls")));
            ClassicAssert.AreEqual(FileMagic.OOXML, FileMagicContainer.ValueOf(POIDataSamples.GetSpreadSheetInstance().GetFile("SampleSS.xlsx")));
        }

        [Test]
        [Order(1)]
        public void TestFileMagicStream()
        {
            InputStream stream = new BufferedInputStream(new FileInputStream(POIDataSamples.GetSpreadSheetInstance().GetFile("SampleSS.xls")));
            try
            {
                ClassicAssert.AreEqual(FileMagic.OLE2, FileMagicContainer.ValueOf(stream));
            }
            finally
            { stream.Close(); }
            stream = new BufferedInputStream(new FileInputStream(POIDataSamples.GetSpreadSheetInstance().GetFile("SampleSS.xlsx")));
            try
            {
                ClassicAssert.AreEqual(FileMagic.OOXML, FileMagicContainer.ValueOf(stream));
            }
            finally
            { stream.Close(); }
        }

        [Test]
        [Order(2)]
        public void TestPrepare()
        {
            InputStream stream = new BufferedInputStream(new FileInputStream(POIDataSamples.GetSpreadSheetInstance().GetFile("SampleSS.xlsx")));
            try
            {
                ClassicAssert.AreSame(stream, FileMagicContainer.PrepareToCheckMagic(stream));
            }
            finally
            { stream.Close(); }

            stream = new InputStream1();
            try
            {
                ClassicAssert.AreNotSame(stream, FileMagicContainer.PrepareToCheckMagic(stream));
            }
            finally
            { stream.Close(); }
        }
    }

    public class InputStream1 : InputStream
    {
        public override bool CanRead => throw new NotImplementedException();

        public override bool CanSeek => throw new NotImplementedException();

        public override long Length => throw new NotImplementedException();

        public override long Position { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override int Read()
        {
            return 0;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotImplementedException();
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }
    }
}


