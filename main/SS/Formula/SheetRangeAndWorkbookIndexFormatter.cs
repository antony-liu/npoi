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

namespace NPOI.SS.Formula
{
    public class SheetRangeAndWorkbookIndexFormatter
    {
        private SheetRangeAndWorkbookIndexFormatter()
        {
        }

        public static string Format(StringBuilder sb, int workbookIndex, string firstSheetName, string lastSheetName)
        {
            if(AnySheetNameNeedsEscaping(firstSheetName, lastSheetName))
            {
                return FormatWithDelimiting(sb, workbookIndex, firstSheetName, lastSheetName);
            }
            else
            {
                return FormatWithoutDelimiting(sb, workbookIndex, firstSheetName, lastSheetName);
            }
        }

        private static string FormatWithDelimiting(StringBuilder sb, int workbookIndex, string firstSheetName, string lastSheetName)
        {
            sb.Append('\'');
            if(workbookIndex >= 0)
            {
                sb.Append('[');
                sb.Append(workbookIndex);
                sb.Append(']');
            }

            SheetNameFormatter.AppendAndEscape(sb, firstSheetName);

            if(lastSheetName != null)
            {
                sb.Append(':');
                SheetNameFormatter.AppendAndEscape(sb, lastSheetName);
            }

            sb.Append('\'');
            return sb.ToString();
        }

        private static string FormatWithoutDelimiting(StringBuilder sb, int workbookIndex, string firstSheetName, string lastSheetName)
        {
            if(workbookIndex >= 0)
            {
                sb.Append('[');
                sb.Append(workbookIndex);
                sb.Append(']');
            }

            sb.Append(firstSheetName);

            if(lastSheetName != null)
            {
                sb.Append(':');
                sb.Append(lastSheetName);
            }

            return sb.ToString();
        }

        private static bool AnySheetNameNeedsEscaping(string firstSheetName, string lastSheetName)
        {
            bool anySheetNameNeedsDelimiting = SheetNameFormatter.NeedsDelimiting(firstSheetName);
            anySheetNameNeedsDelimiting |= SheetNameFormatter.NeedsDelimiting(lastSheetName);
            return anySheetNameNeedsDelimiting;
        }
    }
}


