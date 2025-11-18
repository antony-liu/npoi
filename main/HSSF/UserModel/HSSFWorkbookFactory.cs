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

namespace NPOI.HSSF.UserModel
{
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    /// <summary>
    /// Helper class which is instantiated by reflection from
    /// <see cref="WorkbookFactory.Create(FileInfo)" /> and similar
    /// </summary>
    public class HSSFWorkbookFactory : WorkbookFactory
    {
        /**
         * Create a new empty Workbook
         *
         * @return The created workbook
         */
        public static HSSFWorkbook CreateWorkbook()
        {
            return new HSSFWorkbook();
        }
        /// <summary>
        /// <para>
        /// Creates a HSSFWorkbook from the given NPOIFSFileSystem
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        public static HSSFWorkbook CreateWorkbook(NPOIFSFileSystem fs)
        {
            return new HSSFWorkbook(fs);
        }

        /// <summary>
        /// <para>
        /// Creates a HSSFWorkbook from the given DirectoryNode
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        public static HSSFWorkbook CreateWorkbook(DirectoryNode root)
        {
            return new HSSFWorkbook(root, true);
        }
    }
}


