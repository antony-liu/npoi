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
    using System.Xml;
    /// <summary>
    /// Looks After the collection of end notes for a document.
    /// Managed end notes (<see cref="XWPFEndnote"/>).
    /// </summary>
    /// <remarks>
    /// @since 4.0.0
    /// </remarks>
    public class XWPFEndnotes : XWPFAbstractFootnotesEndnotes
    {
        protected CT_Endnotes ctEndnotes;

        public XWPFEndnotes() : base()
        {

        }

        /// <summary>
        /// Construct XWPFEndnotes from a package part
        /// </summary>
        /// <param name="part">the package part holding the data of the footnotes,</param>
        ///
        /// <remarks>
        /// @since POI 3.14-Beta1
        /// </remarks>
        public XWPFEndnotes(PackagePart part) : base(part)
        {

        }

        /// <summary>
        /// Set the end notes for this part.
        /// </summary>
        /// <param name="endnotes">The endnotes to be added.</param>
        public void SetEndnotes(CT_Endnotes endnotes)
        {
            ctEndnotes = endnotes;
        }

        /// <summary>
        /// Create a new end note and add it to the document.
        /// </summary>
        /// <returns>New XWPFEndnote</returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public XWPFEndnote CreateEndnote()
        {
            CT_FtnEdn newNote = new();
            newNote.type = ST_FtnEdn.normal;

            XWPFEndnote footnote = AddEndnote(newNote);
            footnote.GetCTFtnEdn().id = IdManager.NextId();
            return footnote;

        }

        /// <summary>
        /// Remove the specified footnote if present.
        /// </summary>
        /// <param name="pos"></param>
        /// <returns>True if the footnote was removed.</returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public bool RemoveFootnote(int pos)
        {
            if(ctEndnotes.SizeOfEndnoteArray() >= pos - 1)
            {
                ctEndnotes.RemoveEndnote(pos);
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
            EndnotesDocument notesDoc;
            Stream is1 = null;
            try
            {
                is1 = GetPackagePart().GetInputStream();
                var doc = ConvertStreamToXml(is1);
                notesDoc = EndnotesDocument.Parse(doc, NamespaceManager);
                ctEndnotes = notesDoc.Endnotes;
            }
            catch(XmlException e)
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

            foreach(CT_FtnEdn note in ctEndnotes.endnote)
            {
                listFootnote.Add(new XWPFEndnote(note, this));
            }
        }
        protected internal override void Commit()
        {
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //xmlOptions.SetSaveSyntheticDocumentElement(new QName(CTEndnotes.type.Name.NamespaceURI, "endnotes"));
            PackagePart part = GetPackagePart();
            using(Stream out1 = part.GetOutputStream())
            {
                EndnotesDocument notesDoc = new EndnotesDocument(ctEndnotes);
                notesDoc.Save(out1);
            }
        }

        /// <summary>
        /// add an <see cref="XWPFEndnote"/> to the document
        /// </summary>
        /// <param name="endnote"></param>
        /// <exception cref="IOException"></exception>
        public void AddEndnote(XWPFEndnote endnote)
        {
            listFootnote.Add(endnote);
            ctEndnotes.AddNewEndnote().Set(endnote.GetCTFtnEdn());
        }

        /// <summary>
        /// Add an endnote to the document
        /// </summary>
        /// <param name="note">Note to add</param>
        /// <returns>New <see cref="XWPFEndnote"/></returns>
        /// <exception cref="IOException"></exception>
        public XWPFEndnote AddEndnote(CT_FtnEdn note)
        {
            CT_FtnEdn newNote = ctEndnotes.AddNewEndnote();
            newNote.Set(note);
            XWPFEndnote xNote = new XWPFEndnote(newNote, this);
            listFootnote.Add(xNote);
            return xNote;
        }

        /// <summary>
        /// Get the end note with the specified ID, if any.
        /// </summary>
        /// <param name="id">End note ID.</param>
        /// <returns>The end note or null if not found.</returns>
        public XWPFEndnote GetFootnoteById(int id)
        {
            return (XWPFEndnote) base.GetFootnoteById(id);
        }

        /// <summary>
        /// Get the list of <see cref="XWPFEndnote"/> in the Endnotes part.
        /// </summary>
        /// <returns>List, possibly empty, of end notes.</returns>
        public List<XWPFEndnote> GetEndnotesList()
        {
            List<XWPFEndnote> resultList = new List<XWPFEndnote>();
            foreach(XWPFAbstractFootnoteEndnote note in listFootnote)
            {
                resultList.Add((XWPFEndnote) note);
            }
            return resultList;
        }

        /// <summary>
        /// Remove the specified end note if present.
        /// </summary>
        /// <param name="pos">Array position of the endnote to be removed</param>
        /// <returns>True if the end note was removed.</returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        public bool RemoveEndnote(int pos)
        {
            if(ctEndnotes.SizeOfEndnoteArray() >= pos - 1)
            {
                ctEndnotes.RemoveEndnote(pos);
                listFootnote.RemoveAt(pos);
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}


