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
    /// <summary>
    /// The width types for tables and table cells. Table width can be specified as "auto" (AUTO),
    /// an absolute value in 20ths of a point (DXA), or as a percentage (PCT).
    /// </summary>
    /// <remarks>
    /// @since 4.0.0
    /// </remarks>
    public enum TableWidthType
    {
        Auto,
        Dxa,
        Nil,
        Pct
    }
    public static class TableWidthTypeExtensions
    {
        private static Dictionary<ST_TblWidth, TableWidthType> reverse = new Dictionary<ST_TblWidth, TableWidthType>()
        {
            { ST_TblWidth.auto, TableWidthType.Auto },
            { ST_TblWidth.dxa, TableWidthType.Dxa },
            { ST_TblWidth.nil, TableWidthType.Nil },
            { ST_TblWidth.pct, TableWidthType.Pct },
        };
        public static TableWidthType ValueOf(ST_TblWidth value)
        {
            if(reverse.TryGetValue(value, out var result))
            {
                return result;
            }

            throw new ArgumentException("Invalid ST_TblWidth", nameof(value));
        }
        public static ST_TblWidth ToST_TblWidth(this TableWidthType value)
        {
            switch(value)
            {
                case TableWidthType.Auto:
                    return ST_TblWidth.auto;
                case TableWidthType.Dxa:
                    return ST_TblWidth.dxa;
                case TableWidthType.Nil:
                    return ST_TblWidth.nil;
                case TableWidthType.Pct:
                    return ST_TblWidth.pct;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }
}

