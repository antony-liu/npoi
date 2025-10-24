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
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{

    using NPOI.OpenXml4Net.Exceptions;
    using System.Collections;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Text.RegularExpressions;

    /// <summary>
    /// A package part collection.
    /// </summary>
    public sealed class PackagePartCollection
    {
        /// <summary>
        /// HashSet use to store this collection part names as string for rule
        /// M1.11 optimized checking.
        /// </summary>
        private HashSet<String> registerPartNameStr = new HashSet<String>();

        private SortedDictionary<String, PackagePart> packagePartLookup =
            new SortedDictionary<String, PackagePart>(new PackagePartNameComparator());

        /// <summary>
        /// Check rule [M1.11]: a package implementer shall neither create nor
        /// recognize a part with a part name derived from another part name by
        /// appending segments to it.
        /// </summary>
        /// <param name="partName">name of part</param>
        /// <param name="part">part to Put</param>
        /// <returns>the previous value associated with <c>partName</c>, or
        /// <c>null</c> if there was no mapping for <c>partName</c>.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Throws if you try to add a part with a name derived from
        /// another part name.
        /// </exception>
        public PackagePart Put(PackagePartName partName, PackagePart part)
        {
            string ppName = partName.Name;
            StringBuilder concatSeg = new StringBuilder();
            // split at slash, but keep leading slash
            string delim = "(?=["+PackagingUriHelper.FORWARD_SLASH_STRING+".])";
            foreach(string seg in Regex.Split(ppName, delim))
            {
                concatSeg.Append(seg);
                if(registerPartNameStr.Contains(concatSeg.ToString()))
                {
                    throw new InvalidOperationException(
                        "You can't add a part with a part name derived from another part ! [M1.11]");
                }
            }
            registerPartNameStr.Add(ppName);
            return packagePartLookup[ppName] = part;
        }

        public PackagePart Remove(PackagePartName key)
        {
            if(key == null)
            {
                return null;
            }
            string ppName = key.Name;

            if(packagePartLookup.TryGetValue(ppName, out PackagePart pp))
            {
                this.packagePartLookup.Remove(ppName);
                this.registerPartNameStr.Remove(ppName);
            }
            return pp;
        }


        /// <summary>
        /// The values themselves should be returned in sorted order. Doing it here
        /// avoids paying the high cost of Natural Ordering per insertion.
        /// </summary>
        /// <returns>unmodifiable collection of parts</returns>
        public List<PackagePart> Values
        {
            get
            {
                return packagePartLookup.Values.ToList();
            }

        }
        public List<PackagePart> SortedValues
        {
            get
            {
                return packagePartLookup.Values.ToList();
            }

        }
        public bool ContainsKey(PackagePartName partName)
        {
            return partName != null && packagePartLookup.ContainsKey(partName.Name);
        }

        public PackagePart Get(PackagePartName partName)
        {
            if(partName == null)
            {
                return null;
            }
            packagePartLookup.TryGetValue(partName.Name, out PackagePart pp);
            return pp;
        }

        public int Size
        {
            get
            {
                return packagePartLookup.Count;
            }
        }



        /// <summary>
        /// Get an unused part index based on the namePattern, which doesn't exist yet
        /// and has the lowest positive index
        /// </summary>
        /// <param name="nameTemplate">
        /// The template for new part names containing a <c>'#'</c> for the index,
        /// e.g. "/ppt/slides/slide#.xml"
        /// </param>
        /// <returns>the next available part name index</returns>
        /// <exception cref="InvalidFormatException">if the nameTemplate is null or doesn't contain
        /// the index char (#) or results in an invalid part name
        /// </exception>
        public int GetUnusedPartIndex(string nameTemplate)
        {
            if(nameTemplate == null || !nameTemplate.Contains('#'))
            {
                throw new InvalidFormatException("name template must not be null and contain an index char (#)");
            }

            Regex pattern = new Regex(nameTemplate.Replace("#", "([0-9]+)"));

            Func<String, int> indexFromName = (name) => {
                Match m = pattern.Match(name);
                return m.Success ? Int32.Parse(m.Groups[1].Value) : 0;
            };

            //return packagePartLookup.keySet().stream()
            //    .mapToInt(indexFromName)
            //    .collect(BitSet::new, BitSet::set, BitSet::or).nextClearBit(1);

            var indexArray = packagePartLookup.Keys.Select(name => indexFromName(name));
            BitArray usedIndexes = new BitArray(indexArray.Max() + 1);
            foreach (int x in indexArray)
            {
                usedIndexes.Set(x, true);
            }
            int index = 1;

            for(; index < usedIndexes.Length; index++)
            {
                if (!usedIndexes.Get(index))
                {
                    break;
                }
            }
            return index;
        }
    }

    public class PackagePartNameComparator : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            return PackagePartName.Compare(x, y);
        }
    }
}
