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
    /// <summary>
    /// Base class for the Footnotes and Endnotes part implementations.
    /// </summary>
    /// <remarks>
    /// @since 4.0.0
    /// </remarks>
    public abstract class XWPFAbstractFootnotesEndnotes : POIXMLDocumentPart
    {
        protected XWPFDocument document;
        protected List<XWPFAbstractFootnoteEndnote> listFootnote = [];
        private FootnoteEndnoteIdManager idManager;

        public XWPFAbstractFootnotesEndnotes(OPCPackage pkg)
            : base(pkg)
        {

        }

        public XWPFAbstractFootnotesEndnotes(OPCPackage pkg,
            string coreDocumentRel)
            : base(pkg, coreDocumentRel)
        {

        }

        public XWPFAbstractFootnotesEndnotes()
            : base()
        {
        }

        public XWPFAbstractFootnotesEndnotes(PackagePart part)
            : base(part)
        {

        }

        public XWPFAbstractFootnotesEndnotes(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {

        }


        public XWPFAbstractFootnoteEndnote GetFootnoteById(int id)
        {
            foreach(XWPFAbstractFootnoteEndnote note in listFootnote)
            {
                if(note.GetCTFtnEdn().id == id)
                    return note;
            }
            return null;
        }

        /// <summary>
        /// <see cref="IBody.Part" />
        /// </summary>
        public XWPFDocument XWPFDocument
        {
            get
            {
                if(document != null)
                {
                    return document;
                }
                else
                {
                    return (XWPFDocument) GetParent();
                }
            }
            set
            {
                document = value;
            }
        }

        public FootnoteEndnoteIdManager IdManager
        {
            get
            {
                return this.idManager;
            }
            set
            {
                this.idManager = value;
            }
        }

    }
}


