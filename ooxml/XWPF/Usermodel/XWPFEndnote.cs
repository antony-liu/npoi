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
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;





    /// <summary>
    /// <para>
    /// Represents an end note footnote.
    /// </para>
    /// <para>
    /// End notes are collected at the end of a document or section rather than
    /// at the bottom of a page.
    /// </para>
    /// <para>
    /// Create a new footnote using <see cref="XWPFDocument.CreateEndnote()" /> or
    /// <see cref="XWPFEndnotes.CreateFootnote()" />.
    /// </para>
    /// <para>
    /// The first body element of a footnote should (or possibly must) be a paragraph
    /// with the first run containing a CTFtnEdnRef object. The <see cref="XWPFFootnote.CreateParagraph()" />
    /// and <see cref="XWPFFootnote.CreateTable()" /> methods do this for you.
    /// </para>
    /// <para>
    /// Footnotes have IDs that are unique across all footnotes in the document. You use
    /// the footnote ID to create a reference to a footnote from within a paragraph.
    /// </para>
    /// <para>
    /// To create a reference to a footnote within a paragraph you create a run
    /// with a CTFtnEdnRef that specifies the ID of the target paragraph.
    /// The <see cref="XWPFParagraph.addFootnoteReference(XWPFAbstractFootnoteEndnote)" />
    /// method does this for you.
    /// </para>
    /// </summary>
    /// <remarks>
    /// @since 4.0.0
    /// </remarks>
    public class XWPFEndnote : XWPFAbstractFootnoteEndnote
    {

        public XWPFEndnote() { }
        public XWPFEndnote(XWPFDocument document, CT_FtnEdn body)
                : base(document, body)
        {

        }
        public XWPFEndnote(CT_FtnEdn note, XWPFAbstractFootnotesEndnotes footnotes) : base(note, footnotes)
        {

        }

        /// <summary>
        /// <para>
        /// Ensure that the specified paragraph has a reference marker for this
        /// end note by adding a footnote reference if one is not found.
        /// </para>
        /// <para>
        /// This method is for the first paragraph in the footnote, not
        /// paragraphs that will refer to the footnote. For references to
        /// the footnote, use <see cref="XWPFParagraph.addFootnoteReference(XWPFAbstractFootnoteEndnote))" />.
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
        public override void EnsureFootnoteRef(XWPFParagraph p)
        {
            XWPFRun r = null;
            if(p.Runs.Count > 0)
            {
                r = p.Runs[0];
            }
            if(r == null)
            {
                r = p.CreateRun();
            }
            CT_R ctr = r.GetCTR();
            bool foundRef = false;
            foreach(CT_FtnEdnRef ref1 in ctr.GetEndnoteReferenceList())
            {
                if(Id.Equals(ref1.id))
                {
                    foundRef = true;
                    break;
                }
            }
            if(!foundRef)
            {
                ctr.AddNewRPr().AddNewRStyle().val = "FootnoteReference";
                ctr.AddNewEndnoteRef();
            }

        }
    }
}
