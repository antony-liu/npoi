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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XWPF.UserModel
{
    using NPOI.OOXML;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using System.Xml;

    /// <summary>
    /// <para>
    /// Base class for both bottom-of-the-page footnotes <see cref="XWPFFootnote"/> and end
    /// notes <see cref="XWPFEndnote"/>).
    /// </para>
    /// <para>
    /// The only significant difference between footnotes and
    /// end notes is which part they go on. Footnotes are managed by the Footnotes part
    /// <see cref="XWPFFootnotes"/> and end notes are managed by the Endnotes part <see cref="XWPFEndnotes"/>.
    /// </para>
    /// </summary>
    /// <remarks>
    /// @since 4.0.0
    /// </remarks>
    public abstract class XWPFAbstractFootnoteEndnote : IEnumerable<XWPFParagraph>, IBody
    {
        private List<XWPFParagraph> paragraphs = [];
        private List<XWPFTable> tables = [];
        private List<XWPFPictureData> pictures = [];
        private List<IBodyElement> bodyElements = [];
        protected CT_FtnEdn ctFtnEdn;
        protected XWPFAbstractFootnotesEndnotes footnotes;
        protected XWPFDocument document;

        public XWPFAbstractFootnoteEndnote()
                : base()
        {

        }
        protected XWPFAbstractFootnoteEndnote(XWPFDocument document, CT_FtnEdn body)
        {
            ctFtnEdn = body;
            this.document = document;
            Init();
        }
        protected XWPFAbstractFootnoteEndnote(CT_FtnEdn note, XWPFAbstractFootnotesEndnotes footnotes)
        {
            this.footnotes = footnotes;
            ctFtnEdn = note;
            document = footnotes.XWPFDocument;
            Init();
        }

        protected void Init()
        {
            //XmlCursor cursor = ctFtnEdn.newCursor();
            //copied from XWPFDocument...should centralize this code
            //to avoid duplication
            foreach(object o in ctFtnEdn.Items)
            {

                if(o is CT_P ctP)
                {
                    XWPFParagraph p = new XWPFParagraph(ctP, this);
                    bodyElements.Add(p);
                    paragraphs.Add(p);
                }
                else if(o is CT_Tbl tbl)
                {
                    XWPFTable t = new XWPFTable(tbl, this);
                    bodyElements.Add(t);
                    tables.Add(t);
                }
                else if(o is CT_SdtBlock block)
                {
                    XWPFSDT c = new XWPFSDT(block, this);
                    bodyElements.Add(c);
                }
            }
            //cursor.selectPath("./*");
            //while(cursor.ToNextSelection())
            //{
            //    Xmlobject o = cursor.object;
            //    if(o is CTP)
            //    {
            //        XWPFParagraph p = new XWPFParagraph((CTP) o, this);
            //        bodyElements.add(p);
            //        paragraphs.add(p);
            //    }
            //    else if(o is CTTbl)
            //    {
            //        XWPFTable t = new XWPFTable((CTTbl) o, this);
            //        bodyElements.add(t);
            //        tables.add(t);
            //    }
            //    else if(o is CTSdtBlock)
            //    {
            //        XWPFSDT c = new XWPFSDT((CTSdtBlock) o, this);
            //        bodyElements.add(c);
            //    }

            //}
            //cursor.dispose();

        }

        /// <summary>
        /// Get the list of <see cref="XWPFParagraph"/>s in the footnote.
        /// </summary>
        /// <returns>List of paragraphs</returns>
        public IList<XWPFParagraph> Paragraphs => paragraphs;

        /// <summary>
        /// Get an iterator over the <see cref="XWPFParagraph"/>s in the footnote.
        /// </summary>
        /// <returns>Iterator over the paragraph list.</returns>
        public IEnumerator<XWPFParagraph> GetEnumerator()
        {
            return paragraphs.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// Get the list of <see cref="XWPFTable"/>s in the footnote.
        /// </summary>
        /// <returns>List of tables</returns>
        public IList<XWPFTable> Tables => tables;

        /// <summary>
        /// Gets the list of <see cref="XWPFPictureData"/>s in the footnote.
        /// </summary>
        /// <returns>List of pictures</returns>
        public List<XWPFPictureData> GetPictures()
        {
            return pictures;
        }

        /// <summary>
        /// Gets the body elements (<see cref="IBodyElement"/>) of the footnote.
        /// </summary>
        /// <returns>List of body elements.</returns>
        public IList<IBodyElement> BodyElements => bodyElements;

        /// <summary>
        /// Gets the underlying CTFtnEdn object for the footnote.
        /// </summary>
        /// <returns>CTFtnEdn object</returns>
        public CT_FtnEdn GetCTFtnEdn()
        {
            return ctFtnEdn;
        }

        /// <summary>
        /// <para>
        /// Set the underlying CTFtnEdn for the footnote.
        /// </para>
        /// <para>
        /// Use <see cref="XWPFDocument.CreateFootnote()" /> to create new footnotes.
        /// </para>
        /// </summary>
        /// <param name="footnote">The CTFtnEdn object that will underly the footnote.</param>
        public void SetCTFtnEdn(CT_FtnEdn footnote)
        {
            ctFtnEdn = footnote;
        }

        /// <summary>
        /// Gets the <see cref="XWPFTable"/> at the specified position from the footnote's table array.
        /// </summary>
        /// <param name="pos">in table array</param>
        /// <returns>The <see cref="XWPFTable"/> at position pos, or null if there is no table at position pos.</returns>
        /// @see NPOI.xwpf.UserModel.IBody#getTableArray(int)
        public XWPFTable GetTableArray(int pos)
        {
            if(pos >= 0 && pos < tables.Count)
            {
                return tables[pos];
            }
            return null;
        }

        /// <summary>
        /// Inserts an existing {@link XWPFTable) into the arrays bodyElements and tables.
        /// </summary>
        /// <param name="pos">Position, in the bodyElements array, to insert the table</param>
        /// <param name="table">{@link XWPFTable) to be inserted</param>
        /// @see NPOI.xwpf.UserModel.IBody#insertTable(int pos, XWPFTable table)
        public void InsertTable(int pos, XWPFTable table)
        {
            bodyElements.Insert(pos, table);
            int i = 0;
            foreach(CT_Tbl tbl in ctFtnEdn.GetTblList())
            {
                if(tbl == table.GetCTTbl())
                {
                    break;
                }
                i++;
            }
            tables.Insert(i, table);

        }

        /// <summary>
        /// if there is a corresponding <see cref="XWPFTable"/> of the parameter
        /// ctTable in the tableList of this header
        /// the method will return this table, or null if there is no
        /// corresponding <see cref="XWPFTable"/>.
        /// </summary>
        /// <param name="ctTable"></param>
        /// @see NPOI.xwpf.UserModel.IBody#getTable(CTTbl ctTable)
        public XWPFTable GetTable(CT_Tbl ctTable)
        {
            foreach(XWPFTable table in tables)
            {
                if(table == null)
                    return null;
                if(table.GetCTTbl().Equals(ctTable))
                    return table;
            }
            return null;
        }

        /// <summary>
        /// if there is a corresponding <see cref="XWPFParagraph"/> of the parameter p in the paragraphList of this header or footer
        /// the method will return that paragraph, otherwise the method will return null.
        /// </summary>
        /// <param name="p">The CTP paragraph to find the corresponding <see cref="XWPFParagraph"/> for.</param>
        /// <returns>The <see cref="XWPFParagraph"/> that corresponds to the CTP paragraph in the paragraph
        /// list of this footnote or null if no paragraph is found.
        /// </returns>
        /// @see NPOI.xwpf.UserModel.IBody#getParagraph(CTP p)
        public XWPFParagraph GetParagraph(CT_P p)
        {
            foreach(XWPFParagraph paragraph in paragraphs)
            {
                if(paragraph.GetCTP().Equals(p))
                    return paragraph;
            }
            return null;
        }

        /// <summary>
        /// Returns the <see cref="XWPFParagraph"/> at position pos in footnote's paragraph array.
        /// </summary>
        /// <param name="pos">Array position of the paragraph to Get.</param>
        /// <returns>the <see cref="XWPFParagraph"/> at position pos, or null if there is no paragraph at that position.</returns>
        /// 
        /// @see NPOI.xwpf.UserModel.IBody#getParagraphArray(int pos)
        public XWPFParagraph GetParagraphArray(int pos)
        {
            if(pos >=0 && pos < paragraphs.Count)
            {
                return paragraphs[pos];
            }
            return null;
        }

        /// <summary>
        /// Get the <see cref="XWPFTableCell"/> that belongs to the CTTc cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns><see cref="XWPFTableCell"/> that corresponds to the CTTc cell, if there is one, otherwise null.</returns>
        /// @see NPOI.xwpf.UserModel.IBody#getTableCell(CTTc cell)
        public XWPFTableCell GetTableCell(CT_Tc cell)
        {
            //XmlCursor cursor = cell.newCursor();
            //cursor.ToParent();
            //Xmlobject o = cursor.object;
            //if(!(o is CTRow))
            //{
            //    return null;
            //}
            //CTRow row = (CTRow) o;
            //cursor.ToParent();
            //o = cursor.object;
            //cursor.dispose();
            //if(!(o is CTTbl))
            //{
            //    return null;
            //}
            //CTTbl tbl = (CTTbl) o;
            //XWPFTable table = GetTable(tbl);
            //if(table == null)
            //{
            //    return null;
            //}
            //XWPFTableRow tableRow = table.GetRow(row);
            //if(tableRow == null)
            //{
            //    return null;
            //}
            //return tableRow.GetTableCell(cell);
            throw new NotImplementedException();
        }

        /// <summary>
        /// Verifies that cursor is on the right position.
        /// </summary>
        /// <param name="cursor"></param>
        /// <returns>true if the cursor is within a CTFtnEdn element.</returns>
        private static bool IsCursorInFtn(/*XmlCursor*/ object cursor)
        {
            //XmlCursor verify = cursor.newCursor();
            //verify.ToParent();
            //if(verify.object == this.ctFtnEdn) {
            //    return true;
            //}
            return false;
        }

        /// <summary>
        /// The owning object for this footnote
        /// </summary>
        /// <returns>The <see cref="XWPFFootnotes"/> object that contains this footnote.</returns>
        public POIXMLDocumentPart GetOwner()
        {
            return footnotes;
        }

        /// <summary>
        /// Insert a table constructed from OOXML table markup.
        /// </summary>
        /// <param name="cursor"></param>
        /// <returns>the inserted <see cref="XWPFTable"/></returns>
        /// @see NPOI.xwpf.UserModel.IBody#insertNewTbl(XmlCursor cursor)
        public XWPFTable InsertNewTbl(/*XmlCursor*/ XmlDocument cursor)
        {
            //if(isCursorInFtn(cursor))
            //{
            //    string uri = CTTbl.type.Name.NamespaceURI;
            //    string localPart = "tbl";
            //    cursor.beginElement(localPart, uri);
            //    cursor.ToParent();
            //    CTTbl t = (CTTbl) cursor.object;
            //    XWPFTable newT = new XWPFTable(t, this);
            //    cursor.removeXmlContents();
            //    Xmlobject o = null;
            //    while(!(o is CTTbl) && (cursor.ToPrevSibling()))
            //    {
            //        o = cursor.object;
            //    }
            //    if(!(o is CTTbl))
            //    {
            //        tables.add(0, newT);
            //    }
            //    else
            //    {
            //        int pos = tables.indexOf(getTable((CTTbl) o)) + 1;
            //        tables.add(pos, newT);
            //    }
            //    int i = 0;
            //    cursor = t.newCursor();
            //    while(cursor.ToPrevSibling())
            //    {
            //        o = cursor.object;
            //        if(o is CTP || o is CTTbl)
            //            i++;
            //    }
            //    bodyElements.add(i, newT);
            //    XmlCursor c2 = t.newCursor();
            //    cursor.ToCursor(c2);
            //    cursor.ToEndToken();
            //    c2.dispose();
            //    return newT;
            //}
            //return null;
            throw new NotImplementedException();
        }

        /// <summary>
        /// Add a new <see cref="XWPFParagraph"/> at position of the cursor.
        /// </summary>
        /// <param name="cursor"></param>
        /// <returns>The inserted <see cref="XWPFParagraph"/></returns>
        /// @see NPOI.xwpf.UserModel.IBody#insertNewParagraph(XmlCursor cursor)
        public XWPFParagraph InsertNewParagraph(/*XmlCursor*/ XmlDocument cursor)
        {
            //if(isCursorInFtn(cursor))
            //{
            //    string uri = CTP.type.Name.NamespaceURI;
            //    string localPart = "p";
            //    cursor.beginElement(localPart, uri);
            //    cursor.ToParent();
            //    CTP p = (CTP) cursor.object;
            //    XWPFParagraph newP = new XWPFParagraph(p, this);
            //    Xmlobject o = null;
            //    while(!(o is CTP) && (cursor.ToPrevSibling()))
            //    {
            //        o = cursor.object;
            //    }
            //    if((!(o is CTP)) || o == p)
            //    {
            //        paragraphs.add(0, newP);
            //    }
            //    else
            //    {
            //        int pos = paragraphs.indexOf(getParagraph((CTP) o)) + 1;
            //        paragraphs.add(pos, newP);
            //    }
            //    int i = 0;
            //    XmlCursor p2 = p.newCursor();
            //    cursor.ToCursor(p2);
            //    p2.dispose();
            //    while(cursor.ToPrevSibling())
            //    {
            //        o = cursor.object;
            //        if(o is CTP || o is CTTbl)
            //            i++;
            //    }
            //    bodyElements.add(i, newP);
            //    p2 = p.newCursor();
            //    cursor.ToCursor(p2);
            //    cursor.ToEndToken();
            //    p2.dispose();
            //    return newP;
            //}
            //return null;
            throw new NotImplementedException();
        }

        /// <summary>
        /// Add a new <see cref="XWPFTable"/> to the end of the footnote.
        /// </summary>
        /// <param name="table">CTTbl object from which to construct the <see cref="XWPFTable"/></param>
        /// <returns>The added <see cref="XWPFTable"/></returns>
        public XWPFTable AddNewTbl(CT_Tbl table)
        {
            CT_Tbl newTable = ctFtnEdn.AddNewTbl();
            newTable.Set(table);
            XWPFTable xTable = new XWPFTable(newTable, this);
            tables.Add(xTable);
            return xTable;
        }

        /// <summary>
        /// Add a new <see cref="XWPFParagraph"/> to the end of the footnote.
        /// </summary>
        /// <param name="paragraph">CTP paragraph from which to construct the <see cref="XWPFParagraph"/></param>
        /// <returns>The added <see cref="XWPFParagraph"/></returns>
        public XWPFParagraph AddNewParagraph(CT_P paragraph)
        {
            CT_P newPara = ctFtnEdn.AddNewP(paragraph);
            //newPara.Set(paragraph);
            XWPFParagraph xPara = new XWPFParagraph(newPara, this);
            paragraphs.Add(xPara);
            return xPara;
        }

        /// <summary>
        /// Get the <see cref="XWPFDocument"/> the footnote is part of.
        /// <see cref="NPOI.XWPF.UserModel.IBody.XWPFDocument" />
        /// </summary>
        public XWPFDocument GetXWPFDocument()
        {
            return document;
        }

        /// <summary>
        /// Get the Part to which the footnote belongs, which you need for adding relationships to other parts
        /// </summary>
        /// <returns><see cref="POIXMLDocumentPart"/> that contains the footnote.</returns>
        /// 
        /// @see NPOI.xwpf.UserModel.IBody#getPart()
        public POIXMLDocumentPart Part
        {
            get
            {
                return footnotes;
            }
        }

        /// <summary>
        /// Get the part type  <see cref="BodyType"/> of the footnote.
        /// </summary>
        /// <returns>The <see cref="BodyType"/> value.</returns>
        /// 
        /// @see NPOI.xwpf.UserModel.IBody#getPartType()
        public BodyType PartType => BodyType.FOOTNOTE;

        /// <summary>
        /// <para>
        /// Get the ID of the footnote.
        /// </para>
        /// <para>
        /// Footnote IDs are unique across all bottom-of-the-page and
        /// end note footnotes.
        /// </para>
        /// </summary>
        /// <returns>Footnote ID</returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public int Id
        {
            get
            {
                return this.ctFtnEdn.id;
            }
        }

        /// <summary>
        /// Appends a new <see cref="XWPFParagraph"/> to this footnote.
        /// </summary>
        /// <returns>The new <see cref="XWPFParagraph"/></returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public XWPFParagraph CreateParagraph()
        {
            XWPFParagraph p = new XWPFParagraph(this.ctFtnEdn.AddNewP(), this);
            paragraphs.Add(p);
            bodyElements.Add(p);

            // If the paragraph is the first paragraph in the footnote, 
            // ensure that it has a footnote reference run.

            if(p.Equals(Paragraphs[0]))
            {
                EnsureFootnoteRef(p);
            }
            return p;
        }

        /// <summary>
        /// <para>
        /// Ensure that the specified paragraph has a reference marker for this
        /// footnote by adding a footnote reference if one is not found.
        /// </para>
        /// <para>
        /// This method is for the first paragraph in the footnote, not
        /// paragraphs that will refer to the footnote. For references to
        /// the footnote, use <see cref="XWPFParagraph.addFootnoteReference(XWPFFootnote)" />.
        /// </para>
        /// <para>
        /// The first run of the first paragraph in a footnote should
        /// contain a <see cref="CTFtnEdnRef"/> object.
        /// </para>
        /// </summary>
        /// <param name="p">The <see cref="XWPFParagraph"/> to ensure</param>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public abstract void EnsureFootnoteRef(XWPFParagraph p);

        /// <summary>
        /// Appends a new <see cref="XWPFTable"/> to this footnote
        /// </summary>
        /// <returns>The new <see cref="XWPFTable"/></returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public XWPFTable CreateTable()
        {
            XWPFTable table = new XWPFTable(ctFtnEdn.AddNewTbl(), this);
            if(bodyElements.Count == 0)
            {
                XWPFParagraph p = CreateParagraph();
                EnsureFootnoteRef(p);
            }
            bodyElements.Add(table);
            tables.Add(table);
            return table;
        }

        /// <summary>
        /// Appends a new <see cref="XWPFTable"/> to this footnote
        /// </summary>
        /// <param name="rows">Number of rows to initialize the table with</param>
        /// <param name="cols">Number of columns to initialize the table with</param>
        /// <returns>the new <see cref="XWPFTable"/> with the specified number of rows and columns</returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public XWPFTable CreateTable(int rows, int cols)
        {
            XWPFTable table = new XWPFTable(ctFtnEdn.AddNewTbl(), this, rows, cols);
            bodyElements.Add(table);
            tables.Add(table);
            return table;
        }
    }
}
