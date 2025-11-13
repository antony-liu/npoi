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

using NPOI.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XWPF.UserModel
{




    /// <summary>
    /// <para>
    /// Manages IDs for footnotes and endnotes.
    /// </para>
    /// <para>
    /// Footnotes and endnotes are managed in separate parts but
    /// represent a single namespace of IDs.
    /// </para>
    /// </summary>
    public class FootnoteEndnoteIdManager
    {

        private XWPFDocument document;

        public FootnoteEndnoteIdManager(XWPFDocument document)
        {
            this.document = document;
        }

        /// <summary>
        /// Gets the next ID number.
        /// </summary>
        /// <returns>ID number to use.</returns>
        public int NextId()
        {

            List<int> ids = new List<int>();
            foreach(XWPFAbstractFootnoteEndnote note in document.GetFootnotes())
            {
                ids.Add(note.Id);
            }
            foreach(XWPFAbstractFootnoteEndnote note in document.GetEndnotes())
            {
                ids.Add(note.Id);
            }
            int cand = ids.Count;
            while(ids.Contains(cand))
            {
                cand++;
            }

            return cand;
        }


    }
}


