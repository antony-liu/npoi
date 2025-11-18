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

using MathNet.Numerics.LinearAlgebra.Factorization;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Security.Claims;
using System.Text;

namespace NPOI.SS.UserModel
{
    using NPOI;
    using NPOI.HSSF.Record.Crypto;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.Crypt;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using System.Reflection;
    using System.Threading;



    /// <summary>
    /// Factory for creating the appropriate kind of Workbook
    /// (be it <see cref="HSSFWorkbook"/> or XSSFWorkbook),
    /// by auto-detecting from the supplied input.
    /// </summary>
    public class WorkbookFactory
    {
        /// <summary>
        /// Create a new empty Workbook, either XSSF or HSSF depending
        /// on the parameter
        /// </summary>
        /// <param name="xssf">If an XSSFWorkbook or a HSSFWorkbook should be created</param>
        /// <returns>The created workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        public static IWorkbook Create(bool xssf)
        {
            if(xssf)
            {
                return CreateXSSFWorkbook();
            }
            else
            {
                return CreateHSSFWorkbook();
            }
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
        /// <param name="fs">The <see cref="NPOIFSFileSystem"/> to read the document from</param>
        /// <returns>The created workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        public static IWorkbook Create(NPOIFSFileSystem fs)
        {
            return Create(fs, null);
        }

        /// <summary>
        /// Creates a Workbook from the given NPOIFSFileSystem, which may
        /// be password protected
        /// </summary>
        /// <param name="fs">The <see cref="NPOIFSFileSystem"/> to read the document from</param>
        /// <param name="password">The password that should be used or null if no password is necessary.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        private static IWorkbook Create(NPOIFSFileSystem fs, string password)
        {
            return Create(fs.Root, password);
        }


        /// <summary>
        /// Creates a Workbook from the given NPOIFSFileSystem.
        /// </summary>
        /// <param name="root">The <see cref="DirectoryNode"/> to start reading the document from</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        public static IWorkbook Create(DirectoryNode root)
        {
            return Create(root, null);
        }


        /// <summary>
        /// Creates a Workbook from the given NPOIFSFileSystem, which may
        /// be password protected
        /// </summary>
        /// <param name="root">The <see cref="DirectoryNode"/> to start reading the document from</param>
        /// <param name="password">The password that should be used or null if no password is necessary.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        public static IWorkbook Create(DirectoryNode root, string password)
        {
            // Encrypted OOXML files go inside OLE2 containers, is this one?
            if(root.HasEntry(Decryptor.DEFAULT_POIFS_ENTRY))
            {
                InputStream stream = null;
                try
                {
                    stream = DocumentFactoryHelper.GetDecryptedStream(root, password);

                    return CreateXSSFWorkbook(stream);
                }
                finally
                {
                    IOUtils.CloseQuietly(stream);
                    // as we processed the full stream already, we can close the filesystem here
                    // otherwise file handles are leaked
                    root.FileSystem.Close();
                }
            }

            // If we Get here, it isn't an encrypted PPTX file
            // So, treat it as a regular HSLF PPT one
            bool passwordSet = false;
            if(password != null)
            {
                Biff8EncryptionKey.CurrentUserPassword = password;
                passwordSet = true;
            }
            try
            {
                return CreateHSSFWorkbook(root);
            }
            finally
            {
                if(passwordSet)
                {
                    Biff8EncryptionKey.CurrentUserPassword = null;
                }
            }
        }

        /// <summary>
        /// <para>
        /// Creates a XSSFWorkbook from the given OOXML Package.
        /// As the WorkbookFactory is located in the POI module, which doesn't know about the OOXML formats,
        /// this can be only achieved by using an object reference to the OPCPackage.
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="pkg">The <see cref="OPCPackage"/> opened for reading data.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <remarks>
        /// @deprecated use XSSFWorkbookFactory.Create
        /// </remarks>
        [Obsolete("Use XSSFWorkbookFactory.Create")]
        [Removal(Version = "4.2")]
        public static IWorkbook Create(object pkg)
        {
            return CreateXSSFWorkbook(pkg);
        }

        /// <summary>
        /// <para>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        /// the given InputStream.
        /// </para>
        /// <para>
        /// Your input stream MUST either support mark/reset, or
        /// be wrapped as a <see cref="BufferedInputStream"/>!
        /// Note that using an <see cref="InputStream"/> has a higher memory footprint
        /// than using a <see cref="File"/>.
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use. Note also that loading
        /// from an InputStream requires more memory than loading
        /// from a File, so prefer <see cref="create(File)" /> where possible.
        /// </para>
        /// </summary>
        /// <param name="inp">The <see cref="InputStream"/> to read data from.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="EncryptedDocumentException">If the Workbook given is password protected</exception>
        public static IWorkbook Create(Stream inp)
        {
            return Create(inp, null);
        }

        /// <summary>
        /// <para>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        /// the given InputStream, which may be password protected.
        /// </para>
        /// <para>
        /// Your input stream MUST either support mark/reset, or
        /// be wrapped as a <see cref="BufferedInputStream"/>!
        /// Note that using an <see cref="InputStream"/> has a higher memory footprint
        /// than using a <see cref="File"/>.
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use. Note also that loading
        /// from an InputStream requires more memory than loading
        /// from a File, so prefer <see cref="create(File)" /> where possible.
        /// </para>
        /// </summary>
        /// <param name="inp">The <see cref="InputStream"/> to read data from.</param>
        /// <param name="password">The password that should be used or null if no password is necessary.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="EncryptedDocumentException">If the wrong password is given for a protected file</exception>
        public static IWorkbook Create(Stream inp, string password)
        {
            //InputStream is1 = FileMagicContainer.PrepareToCheckMagic(inp);
            FileMagic fm = FileMagicContainer.ValueOf(inp);

            switch(fm)
            {
                case FileMagic.OLE2:
                    NPOIFSFileSystem fs = new NPOIFSFileSystem(inp);
                    return Create(fs, password);
                case FileMagic.OOXML:
                    return CreateXSSFWorkbook(inp);
                default:
                    throw new IOException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
            }
        }

        /// <summary>
        /// <para>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        /// the given File, which must exist and be readable.
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="file">The file to read data from.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="EncryptedDocumentException">If the Workbook given is password protected</exception>
        public static IWorkbook Create(FileInfo file)
        {
            return Create(file, null);
        }

        /// <summary>
        /// <para>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        /// the given File, which must exist and be readable, and
        /// may be password protected
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="file">The file to read data from.</param>
        /// <param name="password">The password that should be used or null if no password is necessary.</param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="EncryptedDocumentException">If the wrong password is given for a protected file</exception>
        public static IWorkbook Create(FileInfo file, string password)
        {
            return Create(file, password, false);
        }

        /// <summary>
        /// <para>
        /// Creates the appropriate HSSFWorkbook / XSSFWorkbook from
        /// the given File, which must exist and be readable, and
        /// may be password protected
        /// </para>
        /// <para>
        /// Note that in order to properly release resources the
        /// Workbook should be closed After use.
        /// </para>
        /// </summary>
        /// <param name="file">The file to read data from.</param>
        /// <param name="password">The password that should be used or null if no password is necessary.</param>
        /// <param name="readOnly">If the Workbook should be opened in read-only mode to avoid writing back
        /// changes when the document is closed.
        /// </param>
        /// <returns>The created Workbook</returns>
        /// <exception cref="IOException">if an error occurs while reading the data</exception>
        /// <exception cref="EncryptedDocumentException">If the wrong password is given for a protected file</exception>
        public static IWorkbook Create(FileInfo file, string password, bool readOnly)
        {
            if(!file.Exists)
            {
                throw new FileNotFoundException(file.ToString());
            }

            NPOIFSFileSystem fs = null;
            try
            {
                fs = new NPOIFSFileSystem(file, readOnly);
                return Create(fs, password);
            }
            catch(OfficeXmlFileException e)
            {
                IOUtils.CloseQuietly(fs);
                return CreateXSSFWorkbook(file, readOnly);
            }
            catch(RuntimeException e)
            {
                IOUtils.CloseQuietly(fs);
                throw;
            }
        }

        private static IWorkbook CreateHSSFWorkbook(params object[] args)
        {
            return CreateWorkbook("NPOI.Core.dll", "NPOI.HSSF.UserModel.HSSFWorkbookFactory", args);
        }

        private static IWorkbook CreateXSSFWorkbook(params object[] args)
        {
            return CreateWorkbook("NPOI.OOXML.dll", "NPOI.XSSF.UserModel.XSSFWorkbookFactory", args);
        }

        private static IWorkbook CreateWorkbook(string assemblyName, string factoryClass, object[] args)
        {
            try
            {
                Assembly assembly = Assembly.LoadFrom(assemblyName);
                Type clazz = assembly.GetType(factoryClass);
                Type[] argsClz = new Type[args.Length];
                int i=0;
                foreach(object o in args)
                {
                    Type c = o.GetType();
                    //if(Boolean.class.IsAssignableFrom(c)) {
                    //    c = bool.class;
                    //} 
                    //else if (InputStream.class.IsAssignableFrom(c)) {
                    //                    c = InputStream.class;
                    //}
                    argsClz[i++] = c;
                }
                MethodInfo m = clazz.GetMethod("CreateWorkbook", argsClz);
                return (IWorkbook) m.Invoke(null, args);
            }
            catch(TargetInvocationException e)
            {
                Exception t = e.InnerException;
                if(t is IOException)
                {
                    throw (IOException) t;
                }
                else if(t is EncryptedDocumentException)
                {
                    throw (EncryptedDocumentException) t;
                }
                else if(t is OldFileFormatException)
                {
                    throw (OldFileFormatException) t;
                }
                else if(t is UnsupportedFileFormatException)
                {
                    throw (UnsupportedFileFormatException) t;
                }
                else if(t is RuntimeException)
                {
                    throw (RuntimeException) t;
                }
                else
                {
                    throw new IOException(t.Message, t);
                }
            }
            catch(Exception e)
            {
                throw new IOException(e.Message, e);
            }
        }
    }
}
