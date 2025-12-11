using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestFixture]
    public class TestContentTypeManager
    {
        /**
         * Test the properties part content parsing.
         */
        [Test]
        public void TestContentType()
        {
            String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("sample.docx");
            // Retrieves core properties part
            OPCPackage p = OPCPackage.Open(filepath, PackageAccess.READ);
            PackageRelationshipCollection rels = p.GetRelationshipsByType(PackageRelationshipTypes.CORE_PROPERTIES);
            PackageRelationship corePropertiesRelationship = rels.GetRelationship(0);
            PackagePart coreDocument = p.GetPart(corePropertiesRelationship);

            ClassicAssert.AreEqual("application/vnd.openxmlformats-package.core-properties+xml", coreDocument.ContentType);
        }


        /**
         * Test the addition of several default and override content types.
         */
        [Test]
        public void TestContentTypeAddition()
        {
            ContentTypeManager ctm = new ZipContentTypeManager(null, null);

            PackagePartName name1 = PackagingUriHelper.CreatePartName("/foo/foo.XML");
            PackagePartName name2 = PackagingUriHelper.CreatePartName("/foo/foo2.xml");
            PackagePartName name3 = PackagingUriHelper.CreatePartName("/foo/doc.rels");
            PackagePartName name4 = PackagingUriHelper.CreatePartName("/foo/doc.RELS");

            // Add content types
            ctm.AddContentType(name1, "foo-type1");
            ctm.AddContentType(name2, "foo-type2");
            ctm.AddContentType(name3, "text/xml+rel");
            ctm.AddContentType(name4, "text/xml+rel");

            ClassicAssert.AreEqual(ctm.GetContentType(name1), "foo-type1");
            ClassicAssert.AreEqual(ctm.GetContentType(name2), "foo-type2");
            ClassicAssert.AreEqual(ctm.GetContentType(name3), "text/xml+rel");
            ClassicAssert.AreEqual(ctm.GetContentType(name3), "text/xml+rel");
        }
        /**
         * Test the addition then removal of content types.
         */
        [Test]
        public void TestContentTypeRemoval()
        {
            ContentTypeManager ctm = new ZipContentTypeManager(null, null);

            PackagePartName name1 = PackagingUriHelper.CreatePartName("/foo/foo.xml");
            PackagePartName name2 = PackagingUriHelper.CreatePartName("/foo/foo2.xml");
            PackagePartName name3 = PackagingUriHelper.CreatePartName("/foo/doc.rels");
            PackagePartName name4 = PackagingUriHelper.CreatePartName("/foo/doc.RELS");

            // Add content types
            ctm.AddContentType(name1, "foo-type1");
            ctm.AddContentType(name2, "foo-type2");
            ctm.AddContentType(name3, "text/xml+rel");
            ctm.AddContentType(name4, "text/xml+rel");
            ctm.RemoveContentType(name2);
            ctm.RemoveContentType(name3);

            ClassicAssert.AreEqual(ctm.GetContentType(name1), "foo-type1");
            ClassicAssert.AreEqual(ctm.GetContentType(name2), "foo-type1");
            ClassicAssert.IsNull(ctm.GetContentType(name3));

            ctm.RemoveContentType(name1);
            ClassicAssert.IsNull(ctm.GetContentType(name1));
            ClassicAssert.IsNull(ctm.GetContentType(name2));
        }

        protected byte[] ToByteArray(IWorkbook wb)
        {
            using (MemoryStream os = new MemoryStream())
            {
                try
                {
                    wb.Write(os);
                    return os.ToArray();
                }
                catch(IOException)
                {
                    throw new RuntimeException("Assert.Failed to write excel file.");
                }
            } 
        }

        [Test]
        public void Bug62629CombinePictures()
        {
            // this file has incorrect default content-types which caused problems in Apache POI
            // we now handle this broken file more gracefully
            XSSFWorkbook book = XSSFTestDataSamples.OpenSampleWorkbook("62629_target.xlsm");
            XSSFWorkbook b = XSSFTestDataSamples.OpenSampleWorkbook("62629_toMerge.xlsx");
            for(int i = 0; i < b.NumberOfSheets; i++)
            {
                XSSFSheet sheet = book.CreateSheet(b.GetSheetName(i)) as XSSFSheet;
                CopyPictures(sheet, b.GetSheetAt(i));
            }

            XSSFWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(book);
            wbBack.Close();
            book.Close();
            b.Close();
        }

        private static void CopyPictures(ISheet newSheet, ISheet sheet)
        {
            IDrawing<IShape> drawingOld = sheet.CreateDrawingPatriarch();
            IDrawing<IShape> drawingNew = newSheet.CreateDrawingPatriarch();
            ICreationHelper helper = newSheet.Workbook.GetCreationHelper();
            if(drawingNew is XSSFDrawing)
            {
                List<XSSFShape> shapes = ((XSSFDrawing) drawingOld).GetShapes();
                for(int i = 0; i < shapes.Count; i++)
                {
                    if(shapes[i] is XSSFPicture)
                    {
                        XSSFPicture pic = (XSSFPicture) shapes[i];
                        XSSFPictureData picData = pic.PictureData as XSSFPictureData;
                        int pictureIndex = newSheet.Workbook.AddPicture(picData.Data, picData.PictureType);
                        XSSFClientAnchor anchor = null;
                        CT_TwoCellAnchor oldAnchor = ((XSSFDrawing) drawingOld).GetCTDrawing().GetTwoCellAnchorArray(i);
                        if(oldAnchor != null)
                        {
                            anchor = (XSSFClientAnchor) helper.CreateClientAnchor();
                            CT_Marker markerFrom = oldAnchor.from;
                            CT_Marker markerTo = oldAnchor.to;
                            anchor.Dx1 = (int) markerFrom.colOff;
                            anchor.Dx2 = (int) markerTo.colOff;
                            anchor.Dy1 = (int) markerFrom.rowOff;
                            anchor.Dy2 = (int) markerTo.rowOff;
                            anchor.Col1 = markerFrom.col;
                            anchor.Col2 = markerTo.col;
                            anchor.Row1 = markerFrom.row;
                            anchor.Row2 = markerTo.row;
                        }
                        drawingNew.CreatePicture(anchor, pictureIndex);
                    }
                }
            }
        }
    }
}
