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

namespace NPOI.POIFS.Macros
{
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using SixLabors.ImageSharp.Memory;

    public static class RecordTypeExtension
    {
        public static int GetConstantLength(this VBAMacroReader.RecordType type)
        {
            int constantLength = -1;
            switch(type)
            {
                case VBAMacroReader.RecordType.PROJECT_HELP_CONTEXT:
                    constantLength = 8;
                    break;
                case VBAMacroReader.RecordType.PROJECT_VERSION:
                    constantLength = 10;
                    break;
                case VBAMacroReader.RecordType.MODULE_TYPE_PROCEDURAL:
                    constantLength = 4;
                    break;
                case VBAMacroReader.RecordType.MODULE_TYPE_OTHER:
                    constantLength = 4;
                    break;
                case VBAMacroReader.RecordType.MODULE_PRIVATE:
                    constantLength = 4;
                    break;
                default:
                    constantLength = -1;
                    break;
            }
            return constantLength;
        }

        public static VBAMacroReader.RecordType Lookup(int id)
        {
            foreach(VBAMacroReader.RecordType type in Enum.GetValues(typeof(VBAMacroReader.RecordType)))
            {
                if((int) type == id)
                {
                    return type;
                }
            }
            return VBAMacroReader.RecordType.UNKNOWN;
        }
    }
    /// <summary>
    /// <para>
    /// Finds all VBA Macros in an office file (OLE2/POIFS and OOXML/OPC),
    ///  and returns them.
    /// </para>
    /// <para>
    /// </para>
    /// <para>
    /// <b>NOTE:</b> This does not read macros from .ppt files.
    /// See NPOI.HSLF.UserModel.TestBugs.MacrosFromHSLF in the scratchpad
    /// module for an example of how to do this. Patches that make macro
    /// extraction from .ppt more elegant are welcomed!
    /// </para>
    /// </summary>
    /// <remarks>
    /// @since 3.15-beta2
    /// </remarks>
    public class VBAMacroReader : ICloseable
    {
        private static POILogger LOGGER = POILogFactory.GetLogger(typeof(VBAMacroReader));

        //arbitrary limit on size of strings to read, etc.
        private static int MAX_STRING_LENGTH = 20000;
        protected static string VBA_PROJECT_OOXML = "vbaProject.bin";
        protected static string VBA_PROJECT_POIFS = "VBA";

        private NPOIFSFileSystem fs;

        public VBAMacroReader(InputStream rstream)
        {
            InputStream is1 = FileMagicContainer.PrepareToCheckMagic(rstream);
            FileMagic fm = FileMagicContainer.ValueOf(is1);
            if(fm == FileMagic.OLE2)
            {
                fs = new POIFSFileSystem(is1);
            }
            else
            {
                OpenOOXML(is1);
            }
        }

        public VBAMacroReader(FileInfo file)
        {
            try
            {
                this.fs = new POIFSFileSystem(file);
            }
            catch(OfficeXmlFileException e)
            {
                OpenOOXML(new FileInputStream(file.OpenRead()));
            }
        }
        public VBAMacroReader(NPOIFSFileSystem fs)
        {
            this.fs = fs;
        }

        private void OpenOOXML(InputStream zipFile)
        {
            ZipInputStream zis = new ZipInputStream(zipFile);
            try
            {
                ZipEntry zipEntry;
                while((zipEntry = zis.GetNextEntry()) != null)
                {
                    if(zipEntry.Name.EndsWith(VBA_PROJECT_OOXML, StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            // Make a NPOIFS from the contents, and close the stream
                            this.fs = new NPOIFSFileSystem(zis);
                            return;
                        }
                        catch(IOException e)
                        {
                            // Tidy up
                            zis.Close();

                            // Pass on
                            throw;
                        }
                    }
                }
            }
            finally { zis.Close(); }
            throw new ArgumentException("No VBA project found");
        }

        public void Close()
        {
            fs.Close();
            fs = null;
        }

        public Dictionary<String, IModule> ReadMacroModules()
        {
            ModuleMap modules = new ModuleMap();
            //ascii -> unicode mapping for module names
            //preserve insertion order
            Dictionary<String, String> moduleNameMap = new ();

            FindMacros(fs.Root, modules);
            FindModuleNameMap(fs.Root, moduleNameMap, modules);
            FindProjectProperties(fs.Root, moduleNameMap, modules);

            Dictionary<String, IModule> moduleSources = [];
            foreach(KeyValuePair<String, ModuleImpl> entry in modules)
            {
                ModuleImpl module = entry.Value;
                module.charset = modules.charset;
                moduleSources[entry.Key] = module;
            }
            return moduleSources;
        }

        /// <summary>
        /// Reads all macros from all modules of the opened office file.
        /// </summary>
        /// <returns>All the macros and their contents</returns>
        ///
        /// <remarks>
        /// @since 3.15-beta2
        /// </remarks>
        public Dictionary<String, String> ReadMacros()
        {
            Dictionary<String, IModule> modules = ReadMacroModules();
            Dictionary<String, String> moduleSources = [];
            foreach(KeyValuePair<String, IModule> entry in modules)
            {
                moduleSources[entry.Key] = entry.Value.GetContent();
            }
            return moduleSources;
        }

        public class ModuleImpl : IModule
        {
            internal int? offset;
            internal byte[] buf;
            internal ModuleType moduleType;
            internal Encoding charset;
            internal void Read(InputStream in1)
            {
                MemoryStream out1 = new MemoryStream();
                IOUtils.Copy(in1, out1);
                out1.Close();
                buf = out1.ToArray();
            }
            public string GetContent()
            {
                return charset.GetString(buf);
            }
            public ModuleType GetModuleType()
            {
                return moduleType;
            }
        }

        public class ModuleMap : Dictionary<String, ModuleImpl>
        {
            internal Encoding charset = StringUtil.WIN_1252; // default charset
        }


        /// <summary>
        /// Recursively traverses directory structure rooted at <tt>dir</tt>.
        /// For each macro module that is found, the module's name and code are
        /// added to <tt>modules<tt>.
        /// </summary>
        /// <param name="dir">The directory of entries to look at</param>
        /// <param name="modules">The resulting map of modules</param>
        /// <exception cref="IOException">If reading the VBA module Assert.Fails</exception>
        /// <remarks>
        /// @since 3.15-beta2
        /// </remarks>
        protected void FindMacros(DirectoryNode dir, ModuleMap modules)
        {
            if(VBA_PROJECT_POIFS.Equals(dir.Name, StringComparison.OrdinalIgnoreCase))
            {
                // VBA project directory, process
                ReadMacros(dir, modules);
            }
            else
            {
                // Check children
                foreach(Entry child in dir)
                {
                    if(child is DirectoryNode)
                    {
                        FindMacros((DirectoryNode) child, modules);
                    }
                }
            }
        }



        /// <summary>
        /// <para>
        /// reads module from DIR node in input stream and adds it to the modules map for decompression later
        /// on the second pass through this function, the module will be decompressed
        /// </para>
        /// <para>
        /// Side-effects: adds a new module to the module map or Sets the buf field on the module
        /// to the decompressed stream contents (the VBA code for one module)
        /// </para>
        /// </summary>
        /// <param name="in">the run-length encoded input stream to read from</param>
        /// <param name="streamName">the stream name of the module</param>
        /// <param name="modules">a map to store the modules</param>
        /// <exception cref="IOException">If reading data from the stream or from modules Assert.Fails</exception>
        private static void ReadModuleMetadataFromDirStream(RLEDecompressingInputStream in1, string streamName, ModuleMap modules)
        {
            int moduleOffset = in1.ReadInt();
            modules.TryGetValue(streamName, out ModuleImpl module);
            if(module == null)
            {
                // First time we've seen the module. Add it to the ModuleMap and decompress it later
                module = new ModuleImpl();
                module.offset = moduleOffset;
                modules[streamName] = module;
                // Would adding module.read(in) here be correct?
            }
            else
            {
                // Decompress a previously found module and store the decompressed result into module.buf
                RLEDecompressingInputStream stream = new RLEDecompressingInputStream(
                    new MemoryStream(module.buf, moduleOffset, module.buf.Length - moduleOffset)
                );
                module.Read(stream);
                stream.Close();
            }
        }

        private static void ReadModuleFromDocumentStream(DocumentNode documentNode, string name, ModuleMap modules)
        {
            modules.TryGetValue(name, out ModuleImpl module);
            // TODO Refactor this to fetch dir then do the rest
            if(module == null)
            {
                // no DIR stream with offsets yet, so store the compressed bytes for later
                module = new ModuleImpl();
                modules.Add(name, module);
                DocumentInputStream dis = new DocumentInputStream(documentNode);
                try
                {
                    module.Read(dis);
                }
                finally { dis.Close(); }
            }
            else if(module.buf == null)
            { //if we haven't already read the bytes for the module keyed off this name...

                if(module.offset == null)
                {
                    //This should not happen. bug 59858
                    throw new IOException("Module offset for '" + name + "' was never read.");
                }

                //try the general case, where module.offset is accurate
                InputStream decompressed = null;
                DocumentInputStream compressed = new DocumentInputStream(documentNode);
                try
                {
                    // we know the offset already, so decompress immediately on-the-fly
                    long skippedBytes = compressed.Skip(module.offset.Value);
                    if(skippedBytes != module.offset)
                    {
                        throw new IOException("tried to skip " + module.offset + " bytes, but actually skipped " + skippedBytes + " bytes");
                    }
                    decompressed = new RLEDecompressingInputStream(compressed);
                    module.Read(decompressed);
                    return;
                }
                catch(ArgumentException)
                {
                }
                catch(InvalidOperationException) { }
                finally
                {
                    IOUtils.CloseQuietly(compressed);
                    IOUtils.CloseQuietly(decompressed);
                }

                //bad module.offset, try brute force
                compressed = new DocumentInputStream(documentNode);
                byte[] decompressedBytes;
                try
                {
                    decompressedBytes = FindCompressedStreamWBruteForce(compressed);
                }
                finally
                {
                    IOUtils.CloseQuietly(compressed);
                }

                if(decompressedBytes != null)
                {

                    module.Read(new ByteArrayInputStream(decompressedBytes));
                }
            }

        }

        /// <summary>
        /// Skips <tt>n</tt> bytes in an input stream, throwing IOException if the
        /// number of bytes skipped is different than requested.
        /// </summary>
        /// <exception cref="IOException">If skipping would exceed the available data or skipping did not work.</exception>
        private static void TrySkip(InputStream in1, long n)
        {
            long skippedBytes = IOUtils.SkipFully(in1, n);
            if(skippedBytes != n)
            {
                if(skippedBytes < 0)
                {
                    throw new IOException(
                        "Tried skipping " + n + " bytes, but no bytes were skipped. "
                        + "The end of the stream has been reached or the stream is closed.");
                }
                else
                {
                    throw new IOException(
                        "Tried skipping " + n + " bytes, but only " + skippedBytes + " bytes were skipped. "
                        + "This should never happen with a non-corrupt file.");
                }
            }
        }
        // Constants from MS-OVBA: https://msdn.microsoft.com/en-us/library/office/cc313094(v=office.12).aspx
        private static  int STREAMNAME_RESERVED = 0x0032;
        private static  int PROJECT_CONSTANTS_RESERVED = 0x003C;
        private static  int HELP_FILE_PATH_RESERVED = 0x003D;
        private static  int REFERENCE_NAME_RESERVED = 0x003E;
        private static  int DOC_STRING_RESERVED = 0x0040;
        private static  int MODULE_DOCSTRING_RESERVED = 0x0048;

        /**
         * Reads VBA Project modules from a VBA Project directory located at
         * <tt>macroDir</tt> into <tt>modules</tt>.
         *
         * @since 3.15-beta2
         */
        protected void ReadMacros(DirectoryNode macroDir, ModuleMap modules)
        {
            //bug59858 shows that dirstream may not be in this directory (\MBD00082648\_VBA_PROJECT_CUR\VBA ENTRY NAME)
            //but may be in another directory (\_VBA_PROJECT_CUR\VBA ENTRY NAME)
            //process the dirstream first -- "dir" is case insensitive
            foreach(String entryName in macroDir.EntryNames)
            {
                if("dir".Equals(entryName, StringComparison.OrdinalIgnoreCase))
                {
                    ProcessDirStream(macroDir.GetEntry(entryName), modules);
                    break;
                }
            }

            foreach(Entry entry in macroDir)
            {
                if(!(entry is DocumentNode))
                { continue; }

                String name = entry.Name;
                DocumentNode document = (DocumentNode)entry;

                if(!"dir".Equals(name, StringComparison.OrdinalIgnoreCase)
                                    && !name.StartsWith("__SRP", StringComparison.OrdinalIgnoreCase)
                            && !name.StartsWith("_VBA_PROJECT", StringComparison.OrdinalIgnoreCase))
                {
                    // process module, skip __SRP and _VBA_PROJECT since these do not contain macros
                    ReadModuleFromDocumentStream(document, name, modules);
                }
            }
        }

        protected void FindProjectProperties(DirectoryNode node, Dictionary<String, String> moduleNameMap, ModuleMap modules)
        {
            foreach(Entry entry in node)
            {
                if("project".Equals(entry.Name, StringComparison.OrdinalIgnoreCase))
                {
                    DocumentNode document = (DocumentNode)entry;
                    DocumentInputStream dis = new DocumentInputStream(document);
                    try
                    {
                        ReadProjectProperties(dis, moduleNameMap, modules);
                        return;
                    }
                    finally { dis.Close(); }
                }
                else if(entry is DirectoryNode)
                {
                    FindProjectProperties((DirectoryNode) entry, moduleNameMap, modules);
                }
            }
        }

        protected void FindModuleNameMap(DirectoryNode node, Dictionary<String, String> moduleNameMap, ModuleMap modules)
        {
            foreach(Entry entry in node)
            {
                if("projectwm".Equals(entry.Name, StringComparison.OrdinalIgnoreCase))
                {
                    DocumentNode document = (DocumentNode)entry;
                    DocumentInputStream dis = new DocumentInputStream(document);
                    try
                    {
                        ReadNameMapRecords(dis, moduleNameMap, modules.charset);
                        return;
                    }
                    finally { dis.Close(); }
                }
                else if(entry.IsDirectoryEntry)
                {
                    FindModuleNameMap((DirectoryNode) entry, moduleNameMap, modules);
                }
            }
        }

        public enum RecordType
        {
            // Constants from MS-OVBA: https://msdn.microsoft.com/en-us/library/office/cc313094(v=office.12).aspx
            MODULE_OFFSET = 0x0031,
            PROJECT_SYS_KIND = 0x01,
            PROJECT_LCID = 0x0002,
            PROJECT_LCID_INVOKE=0x14,
            PROJECT_CODEPAGE = 0x0003,
            PROJECT_NAME = 0x04,
            PROJECT_DOC_STRING = 0x05,
            PROJECT_HELP_FILE_PATH = 0x06,
            PROJECT_HELP_CONTEXT = 0x07, //(0x07, 8),
            PROJECT_LIB_FLAGS=0x08,
            PROJECT_VERSION = 0x09, //(0x09, 10),
            PROJECT_CONSTANTS=0x0C,
            PROJECT_MODULES=0x0F,
            DIR_STREAM_TERMINATOR=0x10,
            PROJECT_COOKIE=0x13,
            MODULE_NAME=0x19,
            MODULE_NAME_UNICODE=0x47,
            MODULE_STREAM_NAME=0x1A,
            MODULE_DOC_STRING=0x1C,
            MODULE_HELP_CONTEXT=0x1E,
            MODULE_COOKIE=0x2c,
            MODULE_TYPE_PROCEDURAL = 0x21, //(0x21, 4),
            MODULE_TYPE_OTHER = 0x22, //(0x22, 4),
            MODULE_PRIVATE = 0x28, //(0x28, 4),
            REFERENCE_NAME=(0x16),
            REFERENCE_REGISTERED=(0x0D),
            REFERENCE_PROJECT=(0x0E),
            REFERENCE_CONTROL_A=(0x2F),

            //according to the spec, REFERENCE_CONTROL_B(0x33) should have the
            //same structure as REFERENCE_CONTROL_A(0x2F).
            //However, it seems to have the int(length) record structure that most others do.
            //See 59830.xls for this record.
            REFERENCE_CONTROL_B=(0x33),
            //REFERENCE_ORIGINAL(0x33),


            MODULE_TERMINATOR=(0x002B),
            EOF=(-1),
            UNKNOWN=(-2)
        }

        public enum DIR_STATE
        {
            INFORMATION_RECORD,
            REFERENCES_RECORD,
            MODULES_RECORD
        }

        private class ASCIIUnicodeStringPair
        {
            private  String ascii;
            private  String unicode;
            private  int pushbackRecordId;

            internal ASCIIUnicodeStringPair(String ascii, int pushbackRecordId)
            {
                this.ascii = ascii;
                this.unicode = "";
                this.pushbackRecordId = pushbackRecordId;
            }

            internal ASCIIUnicodeStringPair(String ascii, String unicode)
            {
                this.ascii = ascii;
                this.unicode = unicode;
                pushbackRecordId = -1;
            }

            internal String Ascii
            {
                get { return ascii; }
            }

            internal String Unicode
            {
                get { return unicode; }
            }

            internal int PushbackRecordId
            {
                get { return pushbackRecordId; }
            }
        }

        private static void ProcessDirStream(Entry dir, ModuleMap modules)
        {
            DocumentNode dirDocumentNode = (DocumentNode)dir;
            DIR_STATE dirState = DIR_STATE.INFORMATION_RECORD;
            DocumentInputStream dis = new DocumentInputStream(dirDocumentNode);
            try
            {
                String streamName = null;
                int recordId = 0;
                RLEDecompressingInputStream in1 = new RLEDecompressingInputStream(dis);
                try
                {
                    while(true)
                    {
                        recordId = in1.ReadShort();
                        if(recordId == -1)
                        {
                            break;
                        }
                        RecordType type = RecordTypeExtension.Lookup(recordId);

                        if(type.Equals(RecordType.EOF) || type.Equals(RecordType.DIR_STREAM_TERMINATOR))
                        {
                            break;
                        }
                        switch(type)
                        {
                            case RecordType.PROJECT_VERSION:
                                TrySkip(in1, RecordType.PROJECT_VERSION.GetConstantLength());
                                break;
                            case RecordType.PROJECT_CODEPAGE:
                                in1.ReadInt();//record size must == 4
                                int codepage = in1.ReadShort();
                                modules.charset = Encoding.GetEncoding(CodePageUtil.CodepageToEncoding(codepage));
                                break;
                            case RecordType.MODULE_STREAM_NAME:
                                ASCIIUnicodeStringPair pair = ReadStringPair(in1, modules.charset, STREAMNAME_RESERVED);
                                streamName = pair.Ascii;
                                break;
                            case RecordType.PROJECT_DOC_STRING:
                                ReadStringPair(in1, modules.charset, DOC_STRING_RESERVED);
                                break;
                            case RecordType.PROJECT_HELP_FILE_PATH:
                                ReadStringPair(in1, modules.charset, HELP_FILE_PATH_RESERVED);
                                break;
                            case RecordType.PROJECT_CONSTANTS:
                                ReadStringPair(in1, modules.charset, PROJECT_CONSTANTS_RESERVED);
                                break;
                            case RecordType.REFERENCE_NAME:
                                if(dirState.Equals(DIR_STATE.INFORMATION_RECORD))
                                {
                                    dirState = DIR_STATE.REFERENCES_RECORD;
                                }
                                ASCIIUnicodeStringPair stringPair = ReadStringPair(in1,
                                    modules.charset, REFERENCE_NAME_RESERVED, false);
                                if(stringPair.PushbackRecordId == -1)
                                {
                                    break;
                                }
                                //Special handling for when there's only an ascii string and a REFERENCED_REGISTERED
                                //record that follows.
                                //See https://github.com/decalage2/oletools/blob/master/oletools/olevba.py#L1516
                                //and https://github.com/decalage2/oletools/pull/135 from (@c1fe)
                                if(stringPair.PushbackRecordId != (int) RecordType.REFERENCE_REGISTERED)
                                {
                                    throw new ArgumentException("Unexpected reserved character. "+
                                            "Expected "+HexDump.ToHex(REFERENCE_NAME_RESERVED)
                                            + " or "+HexDump.ToHex((int) RecordType.REFERENCE_REGISTERED)+
                                            " not: "+HexDump.ToHex(stringPair.PushbackRecordId));
                                }
                                int recLength = in1.ReadInt();
                                TrySkip(in1, recLength);
                                break;
                            //fall through!
                            case RecordType.REFERENCE_REGISTERED:
                                //REFERENCE_REGISTERED must come immediately after
                                //REFERENCE_NAME to allow for fall through in special case of bug 62625
                                recLength = in1.ReadInt();
                                TrySkip(in1, recLength);
                                break;
                            case RecordType.MODULE_DOC_STRING:
                                int modDocStringLength = in1.ReadInt();
                                ReadString(in1, modDocStringLength, modules.charset);
                                int modDocStringReserved = in1.ReadShort();
                                if(modDocStringReserved != MODULE_DOCSTRING_RESERVED)
                                {
                                    throw new IOException("Expected x003C after stream name before Unicode stream name, but found: " +
                                            HexDump.ToHex(modDocStringReserved));
                                }
                                int unicodeModDocStringLength = in1.ReadInt();
                                ReadUnicodeString(in1, unicodeModDocStringLength);
                                // do something with this at some point
                                break;
                            case RecordType.MODULE_OFFSET:
                                int modOffsetSz = in1.ReadInt();
                                //should be 4
                                ReadModuleMetadataFromDirStream(in1, streamName, modules);
                                break;
                            case RecordType.PROJECT_MODULES:
                                dirState = DIR_STATE.MODULES_RECORD;
                                in1.ReadInt();//size must == 2
                                in1.ReadShort();//number of modules
                                break;
                            case RecordType.REFERENCE_CONTROL_A:
                                int szTwiddled = in1.ReadInt();
                                TrySkip(in1, szTwiddled);
                                int nextRecord = in1.ReadShort();
                                //reference name is optional!
                                if(nextRecord == (int) RecordType.REFERENCE_NAME)
                                {
                                    ReadStringPair(in1, modules.charset, REFERENCE_NAME_RESERVED);
                                    nextRecord = in1.ReadShort();
                                }
                                if(nextRecord != 0x30)
                                {
                                    throw new IOException("Expected 0x30 as Reserved3 in a ReferenceControl record");
                                }
                                int szExtended = in1.ReadInt();
                                TrySkip(in1, szExtended);
                                break;
                            case RecordType.MODULE_TERMINATOR:
                                int endOfModulesReserved = in1.ReadInt();
                                //must be 0;
                                break;
                            default:
                                if(type.GetConstantLength() > -1)
                                {
                                    TrySkip(in1, type.GetConstantLength());
                                }
                                else
                                {
                                    int recordLength = in1.ReadInt();
                                    TrySkip(in1, recordLength);
                                }
                                break;
                        }
                    }
                }
                catch(IOException e)
                {
                    throw new IOException(
                            "Error occurred while reading macros at section id "
                                    + recordId + " (" + HexDump.ShortToHex(recordId) + ")", e);
                }
            }
            finally
            {
                dis.Close();
            }
        }



        private static ASCIIUnicodeStringPair ReadStringPair(
            RLEDecompressingInputStream in1,
            Encoding charset, int reservedByte)
        {
            return ReadStringPair(in1, charset, reservedByte, true);
        }

        private static ASCIIUnicodeStringPair ReadStringPair(
            RLEDecompressingInputStream in1,
            Encoding charset, int reservedByte,
            bool throwOnUnexpectedReservedByte)
        {
            int nameLength = in1.ReadInt();
            String ascii = ReadString(in1, nameLength, charset);
            int reserved = in1.ReadShort();

            if(reserved != reservedByte)
            {
                if(throwOnUnexpectedReservedByte)
                {
                    throw new IOException("Expected " + HexDump.ToHex(reservedByte) +
                            "after name before Unicode name, but found: " +
                            HexDump.ToHex(reserved));
                }
                else
                {
                    return new ASCIIUnicodeStringPair(ascii, reserved);
                }
            }
            int unicodeNameRecordLength = in1.ReadInt();
            String unicode = ReadUnicodeString(in1, unicodeNameRecordLength);
            return new ASCIIUnicodeStringPair(ascii, unicode);
        }

        protected void ReadNameMapRecords(InputStream is1,
            Dictionary<String, String> moduleNames, Encoding charset)
        {
            //see 2.3.3 PROJECTwm Stream: Module Name Information
            //multibytecharstring
            String mbcs = null;
            String unicode = null;
            //arbitrary sanity threshold
            int maxNameRecords = 10000;
            int records = 0;
            while(++records < maxNameRecords)
            {
                try
                {
                    int b = IOUtils.ReadByte(is1);
                    //check for two 0x00 that mark end of record
                    if(b == 0)
                    {
                        b = IOUtils.ReadByte(is1);
                        if(b == 0)
                        {
                            return;
                        }
                    }
                    mbcs = readMBCS(b, is1, charset, MAX_STRING_LENGTH);
                }
                catch(EOFException)
                {
                    return;
                }

                try
                {
                    unicode = readUnicode(is1, MAX_STRING_LENGTH);
                }
                catch(EOFException)
                {
                    return;
                }
                if(mbcs.Trim().Length > 0 && unicode.Trim().Length > 0)
                {
                    moduleNames[mbcs] = unicode;
                }

            }
            if(records >= maxNameRecords)
            {
                //LOGGER.log(POILogger.WARN, "Hit max name records to read ("+maxNameRecords+"). Stopped early.");
            }
        }

        private static String readUnicode(InputStream is1, int maxLength)
        {
            //reads null-terminated unicode string
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            int b0 = IOUtils.ReadByte(is1);
            int b1 = IOUtils.ReadByte(is1);

            int read = 2;
            while((b0 + b1) != 0 && read<maxLength)
            {

                bos.Write(b0);
                bos.Write(b1);
                b0 = IOUtils.ReadByte(is1);
                b1 = IOUtils.ReadByte(is1);
                read += 2;
            }
            if(read >= maxLength)
            {
                //LOGGER.log(POILogger.WARN, "stopped reading unicode name after "+read+" bytes");
            }
            return Encoding.Unicode.GetString(bos.ToByteArray());
        }

        private static String readMBCS(int firstByte, InputStream is1, Encoding charset, int maxLength)
        {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            int len = 0;
            int b = firstByte;
            while(b > 0 && len < maxLength)
            {
                ++len;
                bos.Write(b);
                b = IOUtils.ReadByte(is1);
            }
            return charset.GetString(bos.ToByteArray());
        }

        /**
         * Read <tt>length</tt> bytes of MBCS (multi-byte character set) characters from the stream
         *
         * @param stream the inputstream to read from
         * @param length number of bytes to read from stream
         * @param charset the character set encoding of the bytes in the stream
         * @return a java String in the supplied character set
         * @throws IOException If reading from the stream fails
         */
        private static String ReadString(InputStream stream, int length, Encoding charset)
        {
            byte[] buffer = IOUtils.SafelyAllocate(length, MAX_STRING_LENGTH);
            int bytesRead = IOUtils.ReadFully(stream, buffer);
            if(bytesRead != length)
            {
                throw new IOException("Tried to read: "+length +
                        ", but could only read: "+bytesRead);
            }
            return charset.GetString(buffer, 0, length);
        }

        protected void ReadProjectProperties(DocumentInputStream dis,
                                             Dictionary<String, String> moduleNameMap, ModuleMap modules)
        {
            StreamReader reader = new StreamReader(dis, modules.charset);
            StringBuilder builder = new StringBuilder();
            char[] buffer = new char[512];
            int read;
            while((read = reader.Read(buffer, 0, 512)) > 0)
            {
                builder.Append(buffer, 0, read);
            }
            String properties = builder.ToString();
            //the module name map names should be in exactly the same order
            //as the module names here. See 2.3.3 PROJECTwm Stream.
            //At some point, we might want to enforce that.
#pragma warning disable CA1866
            foreach(String line in properties.Split(["\r\n", "\n\r"], StringSplitOptions.None))
            {
                if(!line.StartsWith("["))
                {
                    String[] tokens = line.Split(['=']);
                    if(tokens.Length > 1 && tokens[1].Length > 1
                            && tokens[1].StartsWith("\"") && tokens[1].EndsWith("\""))
                    {
                        // Remove any double quotes
                        tokens[1] = tokens[1].Substring(1, tokens[1].Length - 2);
                    }
                    if("Document".Equals(tokens[0]) && tokens.Length > 1)
                    {
                        String mn = tokens[1].Substring(0, tokens[1].IndexOf("/&H"));
                        ModuleImpl module = GetModule(mn, moduleNameMap, modules);
                        if(module != null)
                        {
                            module.moduleType = ModuleType.Document;
                        }
                        else
                        {
                            //LOGGER.log(POILogger.WARN, "couldn't find module with name: "+mn);
                        }
                    }
                    else if("Module".Equals(tokens[0]) && tokens.Length > 1)
                    {
                        ModuleImpl module = GetModule(tokens[1], moduleNameMap, modules);
                        if(module != null)
                        {
                            module.moduleType = ModuleType.Module;
                        }
                        else
                        {
                            //LOGGER.log(POILogger.WARN, "couldn't find module with name: "+tokens[1]);
                        }
                    }
                    else if("Class".Equals(tokens[0]) && tokens.Length > 1)
                    {
                        ModuleImpl module = GetModule(tokens[1], moduleNameMap, modules);
                        if(module != null)
                        {
                            module.moduleType = ModuleType.Class;
                        }
                        else
                        {
                            //LOGGER.log(POILogger.WARN, "couldn't find module with name: "+tokens[1]);
                        }
                    }
                }
            }
        }
        //can return null!
        private static ModuleImpl GetModule(String moduleName, Dictionary<String, String> moduleNameMap, ModuleMap moduleMap)
        {
            if(moduleNameMap.TryGetValue(moduleName, out string value))
            {
                return moduleMap[value];
            }
            return moduleMap[moduleName];
        }

        private static String ReadUnicodeString(RLEDecompressingInputStream in1, int unicodeNameRecordLength)
        {
            byte[] buffer = IOUtils.SafelyAllocate(unicodeNameRecordLength, MAX_STRING_LENGTH);
            int bytesRead = IOUtils.ReadFully(in1, buffer);
            if(bytesRead != unicodeNameRecordLength)
            {
                throw new EOFException();
            }
            return Encoding.Unicode.GetString(buffer);
        }

        /**
         * Sometimes the offset record in the dirstream is incorrect, but the macro can still be found.
         * This will try to find the the first RLEDecompressing stream that starts with "Attribute".
         * This relies on some, er, heuristics, admittedly.
         *
         * @param is full module inputstream to read
         * @return uncompressed bytes if found, <code>null</code> otherwise
         * @throws IOException for a true IOException copying the is to a byte array
         */
        private static byte[] FindCompressedStreamWBruteForce(InputStream is1)
        {
            //buffer to memory for multiple tries
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            IOUtils.Copy(is1, bos);
            byte[] compressed = bos.ToByteArray();
            byte[] decompressed = null;
            for(int i = 0; i < compressed.Length; i++)
            {
                if(compressed[i] == 0x01 && i < compressed.Length-1)
                {
                    int w = LittleEndian.GetUShort(compressed, i+1);
                    if(w <= 0 || (w & 0x7000) != 0x3000)
                    {
                        continue;
                    }
                    decompressed = TryToDecompress(new ByteArrayInputStream(compressed, i, compressed.Length - i));
                    if(decompressed != null)
                    {
                        if(decompressed.Length > 9)
                        {
                            //this is a complete hack.  The challenge is that there
                            //can be many 0 length or junk streams that are uncompressed
                            //look in the first 20 characters for "Attribute"
                            int firstX = Math.Min(20, decompressed.Length);
                            String start = StringUtil.WIN_1252.GetString(decompressed, 0, firstX);
                            if(start.Contains("Attribute"))
                            {
                                return decompressed;
                            }
                        }
                    }
                }
            }
            return decompressed;
        }

        private static byte[] TryToDecompress(InputStream is1)
        {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try
            {
                IOUtils.copy(new RLEDecompressingInputStream(is1), bos);
            }
            catch(ArgumentException)
            {
                return null;
            }
            catch(IOException)
            {
                return null;
            }
            catch(InvalidOperationException)
            {
                return null;
            }
            return bos.ToByteArray();
        }

    }
}
