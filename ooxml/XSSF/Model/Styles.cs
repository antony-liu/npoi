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
    using NPOI.XSSF.UserModel;
    using NPOI.XSSF.UserModel.Extensions;

    public interface IStyles
    {

        /// <summary>
        /// Get number format string given its id
        /// </summary>
        /// <param name="fmtId">number format id</param>
        /// <returns>number format code</returns>
        string GetNumberFormatAt(short fmtId);

        /// <summary>
        /// Puts <c>fmt</c> in the numberFormats map if the format is not
        /// already in the the number format style table.
        /// Does nothing if <c>fmt</c> is already in number format style table.
        /// </summary>
        /// <param name="fmt">the number format to add to number format style table</param>
        /// <returns>the index of <c>fmt</c> in the number format style table</returns>
        /// <exception cref="InvalidOperationException">if adding the number format to the styles table
        /// would exceed the max allowed.
        /// </exception>
        int PutNumberFormat(string fmt);

        /// <summary>
        /// Add a number format with a specific ID into the numberFormats map.
        /// If a format with the same ID already exists, overwrite the format code
        /// with <c>fmt</c>
        /// This may be used to override built-in number formats.
        /// </summary>
        /// <param name="index">the number format ID</param>
        /// <param name="fmt">the number format code</param>
        void PutNumberFormat(short index, string fmt);

        /// <summary>
        /// Remove a number format from the style table if it exists.
        /// All cell styles with this number format will be modified to use the default number format.
        /// </summary>
        /// <param name="index">the number format id to remove</param>
        /// <returns>true if the number format was removed</returns>
        bool RemoveNumberFormat(short index);

        /// <summary>
        /// Remove a number format from the style table if it exists
        /// All cell styles with this number format will be modified to use the default number format
        /// </summary>
        /// <param name="fmt">the number format to remove</param>
        /// <returns>true if the number format was removed</returns>
        bool RemoveNumberFormat(string fmt);

        XSSFFont GetFontAt(int idx);

        /// <summary>
        /// Records the given font in the font table.
        /// Will re-use an existing font index if this
        ///  font matches another, EXCEPT if forced
        ///  registration is requested.
        /// This allows people to create several fonts
        ///  then customise them later.
        /// Note - End Users probably want to call
        ///  <see cref="XSSFFont.registerTo(StylesTable)" />
        /// </summary>
        int PutFont(XSSFFont font, bool forceRegistration);

        /// <summary>
        /// Records the given font in the font table.
        /// Will re-use an existing font index if this
        /// font matches another.
        /// </summary>
        int PutFont(XSSFFont font);

        /// <summary>
        /// </summary>
        /// <param name="idx">style index</param>
        /// <returns>XSSFCellStyle or null if idx is out of bounds for xfs array</returns>
        XSSFCellStyle GetStyleAt(int idx);

        int PutStyle(XSSFCellStyle style);

        XSSFCellBorder GetBorderAt(int idx);

        /// <summary>
        /// Adds a border to the border style table if it isn't already in the style table
        /// Does nothing if border is already in borders style table
        /// </summary>
        /// <param name="border">border to add</param>
        /// <returns>the index of the added border</returns>
        int PutBorder(XSSFCellBorder border);

        XSSFCellFill GetFillAt(int idx);

        /// <summary>
        /// Adds a fill to the fill style table if it isn't already in the style table
        /// Does nothing if fill is already in fill style table
        /// </summary>
        /// <param name="fill">fill to add</param>
        /// <returns>the index of the added fill</returns>
        int PutFill(XSSFCellFill fill);

        int NumCellStyles { get; }

        int NumDataFormats { get; }
    }
}
