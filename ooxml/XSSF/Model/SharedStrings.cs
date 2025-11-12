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

namespace NPOI.XSSF.Model
{
    using NPOI.SS.UserModel;


    /// <summary>
    /// <para>
    /// Table of strings shared across all sheets in a workbook.
    /// </para>
    /// <para>
    /// A workbook may contain thousands of cells containing string (non-numeric) data. Furthermore this data is very
    /// likely to be repeated across many rows or columns. The goal of implementing a single string table that is shared
    /// across the workbook is to improve performance in opening and saving the file by only reading and writing the
    /// repetitive information once.
    /// </para>
    /// <para>
    /// </para>
    /// <para>
    /// Consider for example a workbook summarizing information for cities within various countries. There may be a
    /// column for the name of the country, a column for the name of each city in that country, and a column
    /// containing the data for each city. In this case the country name is repetitive, being duplicated in many cells.
    /// In many cases the repetition is extensive, and a tremendous savings is realized by making use of a shared string
    /// table when saving the workbook. When displaying text in the spreadsheet, the cell table will just contain an
    /// index into the string table as the value of a cell, instead of the full string.
    /// </para>
    /// <para>
    /// </para>
    /// <para>
    /// The shared string table contains all the necessary information for displaying the string: the text, formatting
    /// properties, and phonetic properties (for East Asian languages).
    /// </para>
    /// </summary>
    public interface ISharedStrings
    {

        /// <summary>
        /// Return a string item by index
        /// </summary>
        /// <param name="idx">index of item to return.</param>
        /// <returns>the item at the specified position in this Shared string table.</returns>
        IRichTextString GetItemAt(int idx);

        /// <summary>
        /// Return an integer representing the total count of strings in the workbook. This count does not
        /// include any numbers, it counts only the total of text strings in the workbook.
        /// </summary>
        /// <returns>the total count of strings in the workbook</returns>
        int Count { get; }

        /// <summary>
        /// Returns an integer representing the total count of unique strings in the Shared string Table.
        /// A string is unique even if it is a copy of another string, but has different formatting applied
        /// at the character level.
        /// </summary>
        /// <returns>the total count of unique strings in the workbook</returns>
        int UniqueCount { get; }
    }
}
