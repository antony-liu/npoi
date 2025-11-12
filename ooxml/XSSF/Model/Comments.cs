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

using System.Collections.Generic;

namespace NPOI.XSSF.Model
{
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// An interface exposing useful functions for dealing with Excel Workbook Comments.
    /// It is intended that this interface should support low level access and not expose
    /// all the comments in memory
    /// </summary>
    public interface IComments
    {
        int NumberOfComments { get; }

        int NumberOfAuthors {  get; }

        string GetAuthor(long authorId);

        int FindAuthor(string author);

        /// <summary>
        /// Finds the cell comment at cellAddress, if one exists
        /// </summary>
        /// <param name="cellAddress">the address of the cell to find a comment</param>
        /// <returns>cell comment if one exists, otherwise returns null</returns>
        XSSFComment FindCellComment(CellAddress cellAddress);

        /// <summary>
        /// Remove the comment at cellRef location, if one exists
        /// </summary>
        /// <param name="cellRef">the location of the comment to remove</param>
        /// <returns>returns true if a comment was removed</returns>
        bool RemoveComment(CellAddress cellRef);

        /// <summary>
        /// Returns all cell addresses that have comments.
        /// </summary>
        /// <returns>An iterator to traverse all cell addresses that have comments.</returns>
        /// <remarks>
        /// @since 4.0.0
        /// </remarks>
        List<CellAddress> GetCellAddresses();
    }
}


