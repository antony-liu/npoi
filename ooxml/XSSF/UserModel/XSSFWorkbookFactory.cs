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

namespace NPOI.XSSF.UserModel
{
    using NPOI;
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    public class XSSFWorkbookFactory : WorkbookFactory
    {
        /**
         * Create a new empty Workbook
         *
         * @return The created workbook
         */
        public static XSSFWorkbook CreateWorkbook()
        {
            return new XSSFWorkbook();
        }
        /// <summary>
        /// <para>
        /// Creates a XSSFWorkbook from the given OOXML Package.
        /// This is a convenience method to go along the create-methods of the super class.
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="pkg">The <see cref="OPCPackage"/> opened for reading data.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="InvalidFormatException"></exception>
        public static XSSFWorkbook Create(OPCPackage pkg)
        {
            return CreateWorkbook(pkg);
        }

        /// <summary>
        /// <para>
        /// Creates a XSSFWorkbook from the given OOXML Package
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="pkg">The <see cref="ZipPackage"/> opened for reading data.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="InvalidFormatException"></exception>
        public static XSSFWorkbook CreateWorkbook(ZipPackage pkg)
        {
            return CreateWorkbook((OPCPackage) pkg);
        }

        /// <summary>
        /// <para>
        /// Creates a XSSFWorkbook from the given OOXML Package
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="pkg">The <see cref="OPCPackage"/> opened for reading data.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="InvalidFormatException"></exception>
        public static XSSFWorkbook CreateWorkbook(OPCPackage pkg)
        {
            try
            {
                return new XSSFWorkbook(pkg);
            }
            catch(RuntimeException ioe)
            {
                // ensure that file handles are closed (use revert() to not re-write the file)
                pkg.Revert();
                //pkg.Close();

                // rethrow exception
                throw;
            }
        }

        /// <summary>
        /// <para>
        /// Creates the XSSFWorkbook from the given File, which must exist and be readable.
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="file">The file to read data from.</param>
        /// <param name="readOnly">If the Workbook should be opened in read-only mode to avoid writing back
        /// changes when the document is closed.
        /// </param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="EncryptedDocumentException">If the wrong password is given for a protected file</exception>
        public static XSSFWorkbook CreateWorkbook(FileInfo file, bool readOnly)
        {
            OPCPackage pkg = OPCPackage.Open(file, readOnly ? PackageAccess.READ : PackageAccess.READ_WRITE);
            return CreateWorkbook(pkg);
        }

        /// <summary>
        /// <para>
        /// Creates a XSSFWorkbook from the given InputStream
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="stream">The <see cref="Stream"/> to read data from.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="InvalidFormatException"></exception>
        public static XSSFWorkbook CreateWorkbook(Stream stream)
        {
            OPCPackage pkg = OPCPackage.Open(stream);
            return CreateWorkbook(pkg);
        }
    }
}