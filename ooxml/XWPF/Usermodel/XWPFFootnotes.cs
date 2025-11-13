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
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using System.Xml;

    /// <summary>
    /// Looks After the collection of Footnotes for a document.
    /// Manages bottom-of-the-page footnotes (<see cref="XWPFFootnote"/>).
    /// </summary>
    public class XWPFFootnotes : XWPFAbstractFootnotesEndnotes
    {
        protected CT_Footnotes ctFootnotes;
        private readonly List<XWPFHyperlink> hyperlinks = [];
        /// <summary>
        /// Construct XWPFFootnotes from a package part
        /// </summary>
        /// <param name="part">the package part holding the data of the footnotes,</param>
        ///
        /// <remarks>
        /// @since POI 3.14-Beta1
        /// </remarks>
        public XWPFFootnotes(PackagePart part) : base(part)
        {

        }

        /// <summary>
        /// Construct XWPFFootnotes from scratch for a new document.
        /// </summary>
        public XWPFFootnotes()
        {
        }

        /// <summary>
        /// Sets the ctFootnotes
        /// </summary>
        /// <param name="footnotes">Collection of CTFntEdn objects.</param>
        public void SetFootnotes(CT_Footnotes footnotes)
        {
            ctFootnotes = footnotes;
        }

        /// <summary>
        /// Create a new footnote and add it to the document.
        /// </summary>
        /// <returns>New <see cref="XWPFFootnote"/></returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public XWPFFootnote CreateFootnote()
        {
            CT_FtnEdn newNote = new();
            newNote.type = ST_FtnEdn.normal;

            XWPFFootnote footnote = AddFootnote(newNote);
            footnote.GetCTFtnEdn().id = IdManager.NextId();
            return footnote;

        }

        /// <summary>
        /// Remove the specified footnote if present.
        /// </summary>
        /// <param name="pos">Array position of the footnote to be removed</param>
        /// <returns>True if the footnote was removed.</returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public bool RemoveFootnote(int pos)
        {
            if(ctFootnotes.SizeOfFootnoteArray() >= pos - 1)
            {
                ctFootnotes.RemoveFootnote(pos);
                listFootnote.RemoveAt(pos);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Read document
        /// </summary>
        internal override void OnDocumentRead()
        {
            FootnotesDocument notesDoc;
            Stream is1 = null;
            try
            {
                is1 = GetPackagePart().GetInputStream();
                XmlDocument xmldoc = ConvertStreamToXml(is1);
                notesDoc = FootnotesDocument.Parse(xmldoc, NamespaceManager);
                ctFootnotes = notesDoc.Footnotes;
            }
            catch(XmlException)
            {
                throw new POIXMLException();
            }
            finally
            {
                if(is1 != null)
                {
                is1.Close();
                }
            }
            //get any Footnote
            if(ctFootnotes.footnote != null)
            {
                foreach(CT_FtnEdn note in ctFootnotes.footnote)
                {
                    listFootnote.Add(new XWPFFootnote(note, this));
                }
            }
            InitHyperlinks();
        }
        private void InitHyperlinks()
        {
            try
            {
                IEnumerator<PackageRelationship> relIter =
                    GetPackagePart().GetRelationshipsByType(XWPFRelation.HYPERLINK.Relation).GetEnumerator();
                while(relIter.MoveNext())
                {
                    PackageRelationship rel = relIter.Current;
                    hyperlinks.Add(new XWPFHyperlink(rel.Id, rel.TargetUri.OriginalString));
                }
            }
            catch(InvalidDataException e)
            {
                throw new POIXMLException(e);
            }
        }
        protected internal override void Commit()
        {
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //xmlOptions.SetSaveSyntheticDocumentElement(new QName(CTFootnotes.type.Name.NamespaceURI, "footnotes"));
            PackagePart part = GetPackagePart();
            using(Stream out1 = part.GetOutputStream())
            {
                FootnotesDocument notesDoc = new FootnotesDocument(ctFootnotes);
                notesDoc.Save(out1);
            }
        }

        /// <summary>
        /// Add an <see cref="XWPFFootnote"/> to the document
        /// </summary>
        /// <param name="footnote">Footnote to add</param>
        /// <exception cref="IOException"></exception>
        public void AddFootnote(XWPFFootnote footnote)
        {
            listFootnote.Add(footnote);
            ctFootnotes.AddNewFootnote().Set(footnote.GetCTFtnEdn());
        }

        /// <summary>
        /// Add a CT footnote to the document
        /// </summary>
        /// <param name="note">CTFtnEdn to add.</param>
        /// <exception cref="IOException"></exception>
        public XWPFFootnote AddFootnote(CT_FtnEdn note)
        {
            CT_FtnEdn newNote = ctFootnotes.AddNewFootnote();
            newNote.Set(note);
            XWPFFootnote xNote = new XWPFFootnote(newNote, this);
            listFootnote.Add(xNote);
            return xNote;
        }

        /// <summary>
        /// Get the list of <see cref="XWPFFootnote"/> in the Footnotes part.
        /// </summary>
        /// <returns>List, possibly empty, of footnotes.</returns>
        public List<XWPFFootnote> GetFootnotesList()
        {
            List<XWPFFootnote> resultList = new List<XWPFFootnote>();
            foreach(XWPFAbstractFootnoteEndnote note in listFootnote)
            {
                resultList.Add((XWPFFootnote) note);
            }
            return resultList;
        }

        public XWPFHyperlink GetHyperlinkByID(string id)
        {
            foreach(XWPFHyperlink link in hyperlinks)
            {
                if(link.Id.Equals(id))
                {
                    return link;
                }
            }

            return null;
        }

        public List<XWPFHyperlink> GetHyperlinks()
        {
            return hyperlinks;
        }
    }
}
