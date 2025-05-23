/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace TestCases.HPSF.Basic
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;

    using NPOI.HPSF;
    using NPOI.HPSF.Wellknown;
    using NPOI.POIFS.FileSystem;

    using NUnit.Framework;using NUnit.Framework.Legacy;

    /**
     * Tests HPSF's high-level writing functionality for the well-known property
     * Set "SummaryInformation" and "DocumentSummaryInformation".
     * 
     * @author Rainer Klute
     *     <a href="mailto:klute@rainer-klute.de">klute@rainer-klute.de</a>
     * @since 2006-02-01
     * @version $Id: TestWriteWellKnown.java 489730 2006-12-22 19:18:16Z bayard $
     */
    [TestFixture]
    public class TestWriteWellKnown
    {
        //static string dataDir = @"..\..\..\TestCases\HPSF\data\";
        private static String POI_FS = "TestWriteWellKnown.doc";

        /**
         * @see TestCase#SetUp()
         */
        [SetUp]
        public void SetUp()
        {
            VariantSupport.IsLogUnsupportedTypes = false;
        }

        /**
         * This Test method checks whether DocumentSummary information streams
         * can be Read. This is done by opening all "Test*" files in the directrory
         * pointed to by the "HPSF.Testdata.path" system property, trying to extract
         * the document summary information stream in the root directory and calling
         * its Get... methods.
         * @throws IOException 
         * @throws FileNotFoundException 
         * @throws MarkUnsupportedException 
         * @throws NoPropertySetStreamException 
         * @throws UnexpectedPropertySetTypeException 
         */
        [Test]
        public void TestReadDocumentSummaryInformation()
        {
            POIDataSamples _samples = POIDataSamples.GetHPSFInstance();

            string[] files = _samples.GetFiles("Test*.*");

            for (int i = 0; i < files.Length; i++)
            {
                if (!TestReadAllFiles.checkExclude(files[i]))
                    continue;
                using (FileStream doc = new FileStream(files[i], FileMode.Open, FileAccess.Read))
                {
                    Console.WriteLine("Reading file " + doc);
                    try
                    {
                        /* Read a Test document <em>doc</em> into a POI filesystem. */
                        POIFSFileSystem poifs = new POIFSFileSystem(doc);
                        DirectoryEntry dir = poifs.Root;
                        DocumentEntry dsiEntry = null;
                        try
                        {
                            dsiEntry = (DocumentEntry)dir.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
                        }
                        catch (FileNotFoundException)
                        {
                            /*
                             * A missing document summary information stream is not an error
                             * and therefore silently ignored here.
                             */
                        }
                        //catch (System.IO.IOException ex)
                        //{
                        //     // The process cannot access the file 'testcases\test-data\hpsf\TestUnicode.xls' because it is being used by another process.
                        //    Console.Error.WriteLine("Exception ignored (because some other test cases may read this file, too): " + ex.Message);
                        //}

                        /*
                         * If there is a document summry information stream, Read it from
                         * the POI filesystem.
                         */
                        if (dsiEntry != null)
                        {
                            DocumentInputStream dis = new DocumentInputStream(dsiEntry);
                            PropertySet ps = new PropertySet(dis);
                            DocumentSummaryInformation dsi = new DocumentSummaryInformation(ps);

                            /* Execute the Get... methods. */
                            Console.WriteLine(dsi.ByteCount);
                            Console.WriteLine(dsi.ByteOrder);
                            Console.WriteLine(dsi.Category);
                            Console.WriteLine(dsi.Company);
                            Console.WriteLine(dsi.CustomProperties);
                            // FIXME Console.WriteLine(dsi.Docparts);
                            // FIXME Console.WriteLine(dsi.HeadingPair);
                            Console.WriteLine(dsi.HiddenCount);
                            Console.WriteLine(dsi.LineCount);
                            Console.WriteLine(dsi.LinksDirty);
                            Console.WriteLine(dsi.Manager);
                            Console.WriteLine(dsi.MMClipCount);
                            Console.WriteLine(dsi.NoteCount);
                            Console.WriteLine(dsi.ParCount);
                            Console.WriteLine(dsi.PresentationFormat);
                            Console.WriteLine(dsi.Scale);
                            Console.WriteLine(dsi.SlideCount);
                        }
                    }
                    catch (Exception e)
                    {
                        throw new IOException("While handling file " + files[i], e);
                    }
                }
            }
        }

        [Test]
        public void TestWriteWellKnown1()
        {
            POIDataSamples _samples = POIDataSamples.GetHPSFInstance();

            using (FileStream doc1 = _samples.GetFile(POI_FS))
            {
                /* Read a Test document <em>doc1</em> into a POI filesystem. */
                POIFSFileSystem poifs = new POIFSFileSystem(doc1);
                DirectoryEntry dir = poifs.Root;
                DocumentEntry siEntry = (DocumentEntry)dir.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                DocumentEntry dsiEntry = (DocumentEntry)dir.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

                /*
                 * Read the summary information stream and the document summary
                 * information stream from the POI filesystem.
                 * 
                 * Please note that the result consists of SummaryInformation and
                 * DocumentSummaryInformation instances which are in memory only. To
                 * make them permanent they have to be written to a POI filesystem
                 * explicitly (overwriting the former contents). Then the POI filesystem
                 * should be saved to a file.
                 */
                DocumentInputStream dis = new DocumentInputStream(siEntry);
                PropertySet ps = new PropertySet(dis);
                SummaryInformation si = new SummaryInformation(ps);
                dis = new DocumentInputStream(dsiEntry);
                ps = new PropertySet(dis);
                DocumentSummaryInformation dsi = new DocumentSummaryInformation(ps);

                /*
                 * Write all properties supported by HPSF to the summary information
                 * (e.g. author, edit date, application name) and to the document
                 * summary information (e.g. company, manager).
                 */
                Calendar cal = new GregorianCalendar();
                //long time1 = (long)cal.GetMilliseconds(new DateTime(2000, 6, 6, 6, 6, 6));

                //long time2 = (long)cal.GetMilliseconds(new DateTime(2001, 7, 7, 7, 7, 7));
                //long time3 = (long)cal.GetMilliseconds(new DateTime(2002, 8, 8, 8, 8, 8));

                int nr = 4711;
                String P_APPLICATION_NAME = "Microsoft Office Word";
                String P_AUTHOR = "Rainer Klute";
                int P_CHAR_COUNT = 125;
                String P_COMMENTS = "";  //"Comments";
                DateTime P_CREATE_DATE_TIME = new DateTime(2006, 2, 1, 7, 36, 0);
                long P_EDIT_TIME = ++nr * 1000 * 10;
                String P_KEYWORDS = "Test HPSF SummaryInformation DocumentSummaryInformation Writing";
                String P_LAST_AUTHOR = "LastAuthor";
                DateTime? P_LAST_PRINTED = new DateTime(2001, 7, 7, 7, 7, 7);
                DateTime P_LAST_SAVE_DATE_TIME = new DateTime(2008, 9, 30, 9, 54, 0);
                int P_PAGE_COUNT = 1;
                String P_REV_NUMBER = "RevNumber";
                int P_SECURITY = 1;
                String P_SUBJECT = "Subject";
                String P_TEMPLATE = "Normal.dotm";
                // FIXME (byte array properties not yet implemented): byte[] P_THUMBNAIL = new byte[123];
                String P_TITLE = "This document is used for testing POI HPSF��s writing capabilities for the summary information stream and the document summary information stream";
                int P_WORD_COUNT = 21;

                int P_BYTE_COUNT = ++nr;
                String P_CATEGORY = "Category";
                String P_COMPANY = "Rainer Klute IT-Consulting GmbH";
                // FIXME (byte array properties not yet implemented): byte[]  P_DOCPARTS = new byte[123];
                // FIXME (byte array properties not yet implemented): byte[]  P_HEADING_PAIR = new byte[123];
                int P_HIDDEN_COUNT = ++nr;
                int P_LINE_COUNT = ++nr;
                bool P_LINKS_DIRTY = true;
                String P_MANAGER = "Manager";
                int P_MM_CLIP_COUNT = ++nr;
                int P_NOTE_COUNT = ++nr;
                int P_PAR_COUNT = ++nr;
                String P_PRESENTATION_FORMAT = "PresentationFormat";
                bool P_SCALE = false;
                int P_SLIDE_COUNT = ++nr;
                DateTime now = DateTime.Now;

                int POSITIVE_INTEGER = 2222;
                long POSITIVE_LONG = 3333;
                Double POSITIVE_DOUBLE = 4444;
                int NEGATIVE_INTEGER = 2222;
                long NEGATIVE_LONG = 3333;
                Double NEGATIVE_DOUBLE = 4444;

                int MAX_INTEGER = int.MaxValue;
                int MIN_INTEGER = int.MinValue;
                long MAX_LONG = long.MaxValue;
                long MIN_LONG = long.MinValue;
                Double MAX_DOUBLE = Double.MaxValue;
                Double MIN_DOUBLE = Double.MinValue;

                si.ApplicationName = P_APPLICATION_NAME;
                si.Author = P_AUTHOR;
                si.CharCount = P_CHAR_COUNT;
                si.Comments = P_COMMENTS;
                si.CreateDateTime = P_CREATE_DATE_TIME;
                si.EditTime = P_EDIT_TIME;
                si.Keywords = P_KEYWORDS;
                si.LastAuthor = P_LAST_AUTHOR;
                si.LastPrinted = P_LAST_PRINTED;
                si.LastSaveDateTime = P_LAST_SAVE_DATE_TIME;
                si.PageCount = P_PAGE_COUNT;
                si.RevNumber = P_REV_NUMBER;
                si.Security = P_SECURITY;
                si.Subject = P_SUBJECT;
                si.Template = P_TEMPLATE;
                // FIXME (byte array properties not yet implemented): si.Thumbnail=P_THUMBNAIL;
                si.Title = P_TITLE;
                si.WordCount = P_WORD_COUNT;

                dsi.ByteCount = P_BYTE_COUNT;
                dsi.Category = P_CATEGORY;
                dsi.Company = P_COMPANY;
                // FIXME (byte array properties not yet implemented): dsi.Docparts=P_DOCPARTS;
                // FIXME (byte array properties not yet implemented): dsi.HeadingPair=P_HEADING_PAIR;
                dsi.HiddenCount = P_HIDDEN_COUNT;
                dsi.LineCount = P_LINE_COUNT;
                dsi.LinksDirty = P_LINKS_DIRTY;
                dsi.Manager = P_MANAGER;
                dsi.MMClipCount = P_MM_CLIP_COUNT;
                dsi.NoteCount = P_NOTE_COUNT;
                dsi.ParCount = P_PAR_COUNT;
                dsi.PresentationFormat = P_PRESENTATION_FORMAT;
                dsi.Scale = P_SCALE;
                dsi.SlideCount = P_SLIDE_COUNT;

                CustomProperties customProperties = dsi.CustomProperties;
                if (customProperties == null)
                    customProperties = new CustomProperties();
                customProperties.Put("Schl�ssel 1", "Wert 1");
                customProperties.Put("Schl�ssel 2", "Wert 2");
                customProperties.Put("Schl�ssel 3", "Wert 3");
                customProperties.Put("Schl�ssel 4", "Wert 4");
                customProperties.Put("positive_int", POSITIVE_INTEGER);
                customProperties.Put("positive_long", POSITIVE_LONG);
                customProperties.Put("positive_Double", POSITIVE_DOUBLE);
                customProperties.Put("negative_int", NEGATIVE_INTEGER);
                customProperties.Put("negative_long", NEGATIVE_LONG);
                customProperties.Put("negative_Double", NEGATIVE_DOUBLE);
                customProperties.Put("Boolean", true);
                customProperties.Put("Date", now);
                customProperties.Put("max_int", MAX_INTEGER);
                customProperties.Put("min_int", MIN_INTEGER);
                customProperties.Put("max_long", MAX_LONG);
                customProperties.Put("min_long", MIN_LONG);
                customProperties.Put("max_Double", MAX_DOUBLE);
                customProperties.Put("min_Double", MIN_DOUBLE);
                dsi.CustomProperties = customProperties;

                /* Write the summary information stream and the document summary
                 * information stream to the POI filesystem. */
                si.Write(dir, siEntry.Name);
                dsi.Write(dir, dsiEntry.Name);

                /* Write the POI filesystem to a (temporary) file <em>doc2</em>
                 * and Close the latter. */
                using (FileStream doc2 = File.Create(@".\POI_HPSF_Test2.tmp"))
                {
                    poifs.WriteFileSystem(doc2);
                    //doc2.Flush();

                    /*
                     * Open <em>doc2</em> for Reading and check summary information and
                     * document summary information. All properties written before must be
                     * found in the property streams of <em>doc2</em> and have the correct
                     * values.
                     */
                    doc2.Flush();
                    doc2.Position = 0;
                    POIFSFileSystem poifs2 = new POIFSFileSystem(doc2);
                    dir = poifs2.Root;
                    siEntry = (DocumentEntry)dir.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                    dsiEntry = (DocumentEntry)dir.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

                    dis = new DocumentInputStream(siEntry);
                    ps = new PropertySet(dis);
                    si = new SummaryInformation(ps);
                    dis = new DocumentInputStream(dsiEntry);
                    ps = new PropertySet(dis);
                    dsi = new DocumentSummaryInformation(ps);

                    ClassicAssert.AreEqual(P_APPLICATION_NAME, si.ApplicationName);
                    ClassicAssert.AreEqual(P_AUTHOR, si.Author);
                    ClassicAssert.AreEqual(P_CHAR_COUNT, si.CharCount);
                    ClassicAssert.AreEqual(P_COMMENTS, si.Comments);
                    ClassicAssert.AreEqual(P_CREATE_DATE_TIME, si.CreateDateTime);
                    ClassicAssert.AreEqual(P_EDIT_TIME, si.EditTime);
                    ClassicAssert.AreEqual(P_KEYWORDS, si.Keywords);
                    ClassicAssert.AreEqual(P_LAST_AUTHOR, si.LastAuthor);
                    ClassicAssert.AreEqual(P_LAST_PRINTED, si.LastPrinted);
                    ClassicAssert.AreEqual(P_LAST_SAVE_DATE_TIME, si.LastSaveDateTime);
                    ClassicAssert.AreEqual(P_PAGE_COUNT, si.PageCount);
                    ClassicAssert.AreEqual(P_REV_NUMBER, si.RevNumber);
                    ClassicAssert.AreEqual(P_SECURITY, si.Security);
                    ClassicAssert.AreEqual(P_SUBJECT, si.Subject);
                    ClassicAssert.AreEqual(P_TEMPLATE, si.Template);
                    // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(P_THUMBNAIL, si.Thumbnail);
                    ClassicAssert.AreEqual(P_TITLE, si.Title);
                    ClassicAssert.AreEqual(P_WORD_COUNT, si.WordCount);

                    ClassicAssert.AreEqual(P_BYTE_COUNT, dsi.ByteCount);
                    ClassicAssert.AreEqual(P_CATEGORY, dsi.Category);
                    ClassicAssert.AreEqual(P_COMPANY, dsi.Company);
                    // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(P_, dsi.Docparts);
                    // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(P_, dsi.HeadingPair);
                    ClassicAssert.AreEqual(P_HIDDEN_COUNT, dsi.HiddenCount);
                    ClassicAssert.AreEqual(P_LINE_COUNT, dsi.LineCount);
                    ClassicAssert.AreEqual(P_LINKS_DIRTY, dsi.LinksDirty);
                    ClassicAssert.AreEqual(P_MANAGER, dsi.Manager);
                    ClassicAssert.AreEqual(P_MM_CLIP_COUNT, dsi.MMClipCount);
                    ClassicAssert.AreEqual(P_NOTE_COUNT, dsi.NoteCount);
                    ClassicAssert.AreEqual(P_PAR_COUNT, dsi.ParCount);
                    ClassicAssert.AreEqual(P_PRESENTATION_FORMAT, dsi.PresentationFormat);
                    ClassicAssert.AreEqual(P_SCALE, dsi.Scale);
                    ClassicAssert.AreEqual(P_SLIDE_COUNT, dsi.SlideCount);

                    CustomProperties cps = dsi.CustomProperties;
                    //ClassicAssert.AreEqual(customProperties, cps);
                    ClassicAssert.IsNull(cps["No value available"]);
                    ClassicAssert.AreEqual("Wert 1", cps["Schl�ssel 1"]);
                    ClassicAssert.AreEqual("Wert 2", cps["Schl�ssel 2"]);
                    ClassicAssert.AreEqual("Wert 3", cps["Schl�ssel 3"]);
                    ClassicAssert.AreEqual("Wert 4", cps["Schl�ssel 4"]);
                    ClassicAssert.AreEqual(POSITIVE_INTEGER, cps["positive_int"]);
                    ClassicAssert.AreEqual(POSITIVE_LONG, cps["positive_long"]);
                    ClassicAssert.AreEqual(POSITIVE_DOUBLE, cps["positive_Double"]);
                    ClassicAssert.AreEqual(NEGATIVE_INTEGER, cps["negative_int"]);
                    ClassicAssert.AreEqual(NEGATIVE_LONG, cps["negative_long"]);
                    ClassicAssert.AreEqual(NEGATIVE_DOUBLE, cps["negative_Double"]);
                    ClassicAssert.AreEqual(true, cps["Boolean"]);
                    ClassicAssert.AreEqual(now, cps["Date"]);
                    ClassicAssert.AreEqual(MAX_INTEGER, cps["max_int"]);
                    ClassicAssert.AreEqual(MIN_INTEGER, cps["min_int"]);
                    ClassicAssert.AreEqual(MAX_LONG, cps["max_long"]);
                    ClassicAssert.AreEqual(MIN_LONG, cps["min_long"]);
                    ClassicAssert.AreEqual(MAX_DOUBLE, cps["max_Double"]);
                    ClassicAssert.AreEqual(MIN_DOUBLE, cps["min_Double"]);

                    /* Remove all properties supported by HPSF from the summary
                     * information (e.g. author, edit date, application name) and from the
                     * document summary information (e.g. company, manager). */
                    si.RemoveApplicationName();
                    si.RemoveAuthor();
                    si.RemoveCharCount();
                    si.RemoveComments();
                    si.RemoveCreateDateTime();
                    si.RemoveEditTime();
                    si.RemoveKeywords();
                    si.RemoveLastAuthor();
                    si.RemoveLastPrinted();
                    si.RemoveLastSaveDateTime();
                    si.RemovePageCount();
                    si.RemoveRevNumber();
                    si.RemoveSecurity();
                    si.RemoveSubject();
                    si.RemoveTemplate();
                    si.RemoveThumbnail();
                    si.RemoveTitle();
                    si.RemoveWordCount();

                    dsi.RemoveByteCount();
                    dsi.RemoveCategory();
                    dsi.RemoveCompany();
                    dsi.RemoveCustomProperties();
                    dsi.RemoveDocparts();
                    dsi.RemoveHeadingPair();
                    dsi.RemoveHiddenCount();
                    dsi.RemoveLineCount();
                    dsi.RemoveLinksDirty();
                    dsi.RemoveManager();
                    dsi.RemoveMMClipCount();
                    dsi.RemoveNoteCount();
                    dsi.RemoveParCount();
                    dsi.RemovePresentationFormat();
                    dsi.RemoveScale();
                    dsi.RemoveSlideCount();

                    /* 
                     * <li>Write the summary information stream and the document summary
                     * information stream to the POI filesystem. */
                    si.Write(dir, siEntry.Name);
                    dsi.Write(dir, dsiEntry.Name);

                    /* 
                     * <li>Write the POI filesystem to a (temporary) file <em>doc3</em>
                     * and Close the latter. */
                    using (FileStream doc3 = File.Create(@".\POI_HPSF_Test3.tmp"))
                    {
                        poifs2.WriteFileSystem(doc3);
                        doc3.Position = 0;

                        /* 
                         * Open <em>doc3</em> for Reading and check summary information
                         * and document summary information. All properties Removed before must not
                         * be found in the property streams of <em>doc3</em>.
                         */
                        POIFSFileSystem poifs3 = new POIFSFileSystem(doc3);

                        dir = poifs3.Root;
                        siEntry = (DocumentEntry)dir.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                        dsiEntry = (DocumentEntry)dir.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);

                        dis = new DocumentInputStream(siEntry);
                        ps = new PropertySet(dis);
                        si = new SummaryInformation(ps);
                        dis = new DocumentInputStream(dsiEntry);
                        ps = new PropertySet(dis);
                        dsi = new DocumentSummaryInformation(ps);

                        ClassicAssert.AreEqual(null, si.ApplicationName);
                        ClassicAssert.AreEqual(null, si.Author);
                        ClassicAssert.AreEqual(0, si.CharCount);
                        ClassicAssert.IsTrue(si.WasNull);
                        ClassicAssert.AreEqual(null, si.Comments);
                        ClassicAssert.AreEqual(null, si.CreateDateTime);
                        ClassicAssert.AreEqual(0, si.EditTime);
                        ClassicAssert.IsTrue(si.WasNull);
                        ClassicAssert.AreEqual(null, si.Keywords);
                        ClassicAssert.AreEqual(null, si.LastAuthor);
                        ClassicAssert.AreEqual(null, si.LastPrinted);
                        ClassicAssert.AreEqual(null, si.LastSaveDateTime);
                        ClassicAssert.AreEqual(0, si.PageCount);
                        ClassicAssert.IsTrue(si.WasNull);
                        ClassicAssert.AreEqual(null, si.RevNumber);
                        ClassicAssert.AreEqual(0, si.Security);
                        ClassicAssert.IsTrue(si.WasNull);
                        ClassicAssert.AreEqual(null, si.Subject);
                        ClassicAssert.AreEqual(null, si.Template);
                        ClassicAssert.AreEqual(null, si.Thumbnail);
                        ClassicAssert.AreEqual(null, si.Title);
                        ClassicAssert.AreEqual(0, si.WordCount);
                        ClassicAssert.IsTrue(si.WasNull);

                        ClassicAssert.AreEqual(0, dsi.ByteCount);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(null, dsi.Category);
                        ClassicAssert.AreEqual(null, dsi.CustomProperties);
                        // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(null, dsi.Docparts);
                        // FIXME (byte array properties not yet implemented): ClassicAssert.AreEqual(null, dsi.HeadingPair);
                        ClassicAssert.AreEqual(0, dsi.HiddenCount);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(0, dsi.LineCount);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(false, dsi.LinksDirty);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(null, dsi.Manager);
                        ClassicAssert.AreEqual(0, dsi.MMClipCount);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(0, dsi.NoteCount);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(0, dsi.ParCount);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(null, dsi.PresentationFormat);
                        ClassicAssert.AreEqual(false, dsi.Scale);
                        ClassicAssert.IsTrue(dsi.WasNull);
                        ClassicAssert.AreEqual(0, dsi.SlideCount);
                        ClassicAssert.IsTrue(dsi.WasNull);
                    }
                }
            }

            if (File.Exists(@".\POI_HPSF_Test3.tmp"))
            {
                File.Delete(@".\POI_HPSF_Test3.tmp");
            }

            if (File.Exists(@".\POI_HPSF_Test2.tmp"))
            {
                File.Delete(@".\POI_HPSF_Test2.tmp");
            }
        }

        private void RunTest(FileStream file)
        {
            /* Read a Test document <em>doc</em> into a POI filesystem. */
            POIFSFileSystem poifs = new POIFSFileSystem(file);
            DirectoryEntry dir = poifs.Root;
            DocumentEntry dsiEntry = null;
            try
            {
                dsiEntry = (DocumentEntry)dir.GetEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME);
            }
            catch (FileNotFoundException)
            {
                /*
                 * A missing document summary information stream is not an error
                 * and therefore silently ignored here.
                 */
            }

            /*
             * If there is a document summry information stream, Read it from
             * the POI filesystem, else Create a new one.
             */
            DocumentSummaryInformation dsi;
            if (dsiEntry != null)
            {
                DocumentInputStream dis = new DocumentInputStream(dsiEntry);
                PropertySet ps = new PropertySet(dis);
                dsi = new DocumentSummaryInformation(ps);
            }
            else
                dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            CustomProperties cps = dsi.CustomProperties;

            if (cps == null)
                /* The document does not have custom properties. */
                return;

            foreach (var de in cps)
            {
                CustomProperty cp = (CustomProperty)de.Value;
                ClassicAssert.IsNotNull(cp.Name);
                ClassicAssert.IsNotNull(cp.Value);
            }
        }

        /**
         * Tests the simplified custom properties by Reading them from the
         * available Test files.
         *
         * @throws Exception if anything goes wrong.
         */
        [Test]
        public void TestReadCustomPropertiesFromFiles()
        {
            POIDataSamples _samples = POIDataSamples.GetHPSFInstance();
            string[] files = _samples.GetFiles("Test*.*");

            for (int i = 0; i < files.Length; i++)
            {
                if (TestReadAllFiles.checkExclude(files[i]))
                    continue;
                using (FileStream file = new FileStream(files[i], FileMode.Open, FileAccess.Read))
                {
                    try
                    {
                        RunTest(file);
                    }
                    catch (Exception e)
                    {
                        throw new IOException("While handling file " + files[i], e);
                    }
                }
            }
        }

        /**
         * Tests basic custom property features.
         */
        [Test]
        public void TestCustomerProperties()
        {
            String KEY = "Schl\u00fcssel \u00e4";
            String VALUE_1 = "Wert 1";
            String VALUE_2 = "Wert 2";

            CustomProperty cp;
            CustomProperties cps = new CustomProperties();
            ClassicAssert.AreEqual(0, cps.Count);

            /* After Adding a custom property the size must be 1 and it must be
             * possible to extract the custom property from the map. */
            cps.Put(KEY, VALUE_1);
            ClassicAssert.AreEqual(1, cps.Count);
            Object v1 = cps[KEY];
            ClassicAssert.AreEqual(VALUE_1, v1);

            /* After Adding a custom property with the same name the size must still
             * be one. */
            cps.Put(KEY, VALUE_2);
            ClassicAssert.AreEqual(1, cps.Count);
            Object v2 = cps[KEY];
            ClassicAssert.AreEqual(VALUE_2, v2);

            /* Removing the custom property must return the Remove property and
             * reduce the size to 0. */
            cp = (CustomProperty)cps.Remove(KEY);
            ClassicAssert.AreEqual(KEY, cp.Name);
            ClassicAssert.AreEqual(VALUE_2, cp.Value);
            ClassicAssert.AreEqual(0, cps.Count);
        }

        /**
         * Tests Reading custom properties from a section including Reading
         * custom properties which are not pure.
         */
        [Test]
        public void TestGetCustomerProperties()
        {
            long ID_1 = 2;
            long ID_2 = 3;
            String NAME_1 = "Schl\u00fcssel \u00e4";
            String VALUE_1 = "Wert 1";
            Dictionary<long, string> dictionary = new();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            CustomProperties cps;
            MutableSection s;

            /* A document summary information Set stream by default does have custom properties. */
            cps = dsi.CustomProperties;
            ClassicAssert.AreEqual(null, cps);

            /* Test an empty custom properties Set. */
            s = new MutableSection();
            s.SetFormatID(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID[1]);
            // s.SetCodepage(CodePageUtil.CP_UNICODE);
            dsi.AddSection(s);
            cps = dsi.CustomProperties;
            ClassicAssert.AreEqual(0, cps.Count);

            /* Add a custom property. */
            MutableProperty p = new MutableProperty();
            p.ID = ID_1;
            p.Type = Variant.VT_LPWSTR;
            p.Value = VALUE_1;
            s.SetProperty(p);
            dictionary[ID_1] = NAME_1;
            s.SetDictionary(dictionary);
            cps = dsi.CustomProperties;
            ClassicAssert.AreEqual(1, cps.Count);
            ClassicAssert.IsTrue(cps.IsPure);

            /* Add another custom property. */
            s.SetProperty((int)ID_2, Variant.VT_LPWSTR, VALUE_1);
            dictionary[ID_2] = NAME_1;
            s.SetDictionary(dictionary);
            cps = dsi.CustomProperties;
            ClassicAssert.AreEqual(1, cps.Count);
            ClassicAssert.IsFalse(cps.IsPure);
        }
    }
}