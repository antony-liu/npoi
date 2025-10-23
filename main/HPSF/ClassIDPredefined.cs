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

namespace NPOI.HPSF
{

    public enum ClassIDPredefinedEnum
    {
        /** OLE 1.0 package manager */
        OLE_V1_PACKAGE,
        /** Excel V3 - document */
        EXCEL_V3,
        /** Excel V3 - chart */
        EXCEL_V3_CHART,
        /** Excel V3 - macro */
        EXCEL_V3_MACRO,
        /** Excel V7 / 95 - document */
        EXCEL_V7,
        /** Excel V7 / 95 - workbook */
        EXCEL_V7_WORKBOOK,
        /** Excel V7 / 95 - chart */
        EXCEL_V7_CHART,
        /** Excel V8 / 97 - document */
        EXCEL_V8,
        /** Excel V8 / 97 - chart */
        EXCEL_V8_CHART,
        /** Excel V11 / 2003 - document */
        EXCEL_V11,
        /** Excel V12 / 2007 - document */
        EXCEL_V12,
        /** Excel V12 / 2007 - macro */
        EXCEL_V12_MACRO,
        /** Excel V12 / 2007 - xlsb document */
        EXCEL_V12_XLSB,
        /* Excel V14 / 2010 - document */
        EXCEL_V14,
        /* Excel V14 / 2010 - workbook */
        EXCEL_V14_WORKBOOK,
        /* Excel V14 / 2010 - chart */
        EXCEL_V14_CHART,
        /** Excel V14 / 2010 - OpenDocument spreadsheet */
        EXCEL_V14_ODS,
        /** Word V7 / 95 - document */
        WORD_V7,
        /** Word V8 / 97 - document */
        WORD_V8,
        /** Word V12 / 2007 - document */
        WORD_V12,
        /** Word V12 / 2007 - macro */
        WORD_V12_MACRO,
        /** Powerpoint V7 / 95 - document */
        POWERPOINT_V7,
        /** Powerpoint V7 / 95 - slide */
        POWERPOINT_V7_SLIDE,
        /** Powerpoint V8 / 97 - document */
        POWERPOINT_V8,
        /** Powerpoint V8 / 97 - template */
        POWERPOINT_V8_TPL,
        /** Powerpoint V12 / 2007 - document */
        POWERPOINT_V12,
        /** Powerpoint V12 / 2007 - macro */
        POWERPOINT_V12_MACRO,
        /** Publisher V12 */
        PUBLISHER_V12,
        /** Visio 2000 (V6) / 2002 (V10) - Drawing */
        VISIO_V10,
        /** Equation Editor 3.0 */
        EQUATION_V3,
        /** AcroExch.Document */
        PDF,
        /** Plain Text Persistent Handler **/
        TXT_ONLY

    }

    public class ClassIDPredefined
    {
        /** OLE 1.0 package manager */
        public static ClassIDPredefined OLE_V1_PACKAGE = new ClassIDPredefined("{0003000C-0000-0000-C000-000000000046}", ".bin", null);
        /** Excel V3 - document */
        public static ClassIDPredefined EXCEL_V3 = new ClassIDPredefined("{00030000-0000-0000-C000-000000000046}", ".xls", "application/vnd.ms-excel");
        /** Excel V3 - chart */
        public static ClassIDPredefined EXCEL_V3_CHART = new ClassIDPredefined("{00030001-0000-0000-C000-000000000046}", null, null);
        /** Excel V3 - macro */
        public static ClassIDPredefined EXCEL_V3_MACRO = new ClassIDPredefined("{00030002-0000-0000-C000-000000000046}", null, null);
        /** Excel V7 / 95 - document */
        public static ClassIDPredefined EXCEL_V7 = new ClassIDPredefined("{00020810-0000-0000-C000-000000000046}", ".xls", "application/vnd.ms-excel");
        /** Excel V7 / 95 - workbook */
        public static ClassIDPredefined EXCEL_V7_WORKBOOK = new ClassIDPredefined("{00020841-0000-0000-C000-000000000046}", null, null);
        /** Excel V7 / 95 - chart */
        public static ClassIDPredefined EXCEL_V7_CHART = new ClassIDPredefined("{00020811-0000-0000-C000-000000000046}", null, null);
        /** Excel V8 / 97 - document */
        public static ClassIDPredefined EXCEL_V8 = new ClassIDPredefined("{00020820-0000-0000-C000-000000000046}", ".xls", "application/vnd.ms-excel");
        /** Excel V8 / 97 - chart */
        public static ClassIDPredefined EXCEL_V8_CHART = new ClassIDPredefined("{00020821-0000-0000-C000-000000000046}", null, null);
        /** Excel V11 / 2003 - document */
        public static ClassIDPredefined EXCEL_V11 = new ClassIDPredefined("{00020812-0000-0000-C000-000000000046}", ".xlsx", "application/vnd.Openxmlformats-officedocument.spreadsheetml.sheet");
        /** Excel V12 / 2007 - document */
        public static ClassIDPredefined EXCEL_V12 = new ClassIDPredefined("{00020830-0000-0000-C000-000000000046}", ".xlsx", "application/vnd.Openxmlformats-officedocument.spreadsheetml.sheet");
        /** Excel V12 / 2007 - macro */
        public static ClassIDPredefined EXCEL_V12_MACRO = new ClassIDPredefined("{00020832-0000-0000-C000-000000000046}", ".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12");
        /** Excel V12 / 2007 - xlsb document */
        public static ClassIDPredefined EXCEL_V12_XLSB = new ClassIDPredefined("{00020833-0000-0000-C000-000000000046}", ".xlsb", "application/vnd.ms-excel.sheet.binary.macroEnabled.12");
        /* Excel V14 / 2010 - document */
        public static ClassIDPredefined EXCEL_V14 = new ClassIDPredefined("{00024500-0000-0000-C000-000000000046}", ".xlsx", "application/vnd.Openxmlformats-officedocument.spreadsheetml.sheet");
        /* Excel V14 / 2010 - workbook */
        public static ClassIDPredefined EXCEL_V14_WORKBOOK = new ClassIDPredefined("{000208D5-0000-0000-C000-000000000046}", null, null);
        /* Excel V14 / 2010 - chart */
        public static ClassIDPredefined EXCEL_V14_CHART = new ClassIDPredefined("{00024505-0014-0000-C000-000000000046}", null, null);
        /** Excel V14 / 2010 - OpenDocument spreadsheet */
        public static ClassIDPredefined EXCEL_V14_ODS = new ClassIDPredefined("{EABCECDB-CC1C-4A6F-B4E3-7F888A5ADFC8}", ".ods", "application/vnd.oasis.Opendocument.spreadsheet");
        /** Word V7 / 95 - document */
        public static ClassIDPredefined WORD_V7 = new ClassIDPredefined("{00020900-0000-0000-C000-000000000046}", ".doc", "application/msword");
        /** Word V8 / 97 - document */
        public static ClassIDPredefined WORD_V8 = new ClassIDPredefined("{00020906-0000-0000-C000-000000000046}", ".doc", "application/msword");
        /** Word V12 / 2007 - document */
        public static ClassIDPredefined WORD_V12 = new ClassIDPredefined("{F4754C9B-64F5-4B40-8AF4-679732AC0607}", ".docx", "application/vnd.Openxmlformats-officedocument.wordprocessingml.document");
        /** Word V12 / 2007 - macro */
        public static ClassIDPredefined WORD_V12_MACRO = new ClassIDPredefined("{18A06B6B-2F3F-4E2B-A611-52BE631B2D22}", ".docm", "application/vnd.ms-word.document.macroEnabled.12");
        /** Powerpoint V7 / 95 - document */
        public static ClassIDPredefined POWERPOINT_V7 = new ClassIDPredefined("{EA7BAE70-FB3B-11CD-A903-00AA00510EA3}", ".ppt", "application/vnd.ms-powerpoint");
        /** Powerpoint V7 / 95 - slide */
        public static ClassIDPredefined POWERPOINT_V7_SLIDE = new ClassIDPredefined("{EA7BAE71-FB3B-11CD-A903-00AA00510EA3}", null, null);
        /** Powerpoint V8 / 97 - document */
        public static ClassIDPredefined POWERPOINT_V8 = new ClassIDPredefined("{64818D10-4F9B-11CF-86EA-00AA00B929E8}", ".ppt", "application/vnd.ms-powerpoint");
        /** Powerpoint V8 / 97 - template */
        public static ClassIDPredefined POWERPOINT_V8_TPL = new ClassIDPredefined("{64818D11-4F9B-11CF-86EA-00AA00B929E8}", ".pot", "application/vnd.ms-powerpoint");
        /** Powerpoint V12 / 2007 - document */
        public static ClassIDPredefined POWERPOINT_V12 = new ClassIDPredefined("{CF4F55F4-8F87-4D47-80BB-5808164BB3F8}", ".pptx", "application/vnd.Openxmlformats-officedocument.presentationml.presentation");
        /** Powerpoint V12 / 2007 - macro */
        public static ClassIDPredefined POWERPOINT_V12_MACRO = new ClassIDPredefined("{DC020317-E6E2-4A62-B9FA-B3EFE16626F4}", ".pptm", "application/vnd.ms-powerpoint.presentation.macroEnabled.12");
        /** Publisher V12 */
        public static ClassIDPredefined PUBLISHER_V12 = new ClassIDPredefined("{0002123D-0000-0000-C000-000000000046}", ".pub", "application/x-mspublisher");
        /** Visio 2000 (V6) / 2002 (V10) - Drawing */
        public static ClassIDPredefined VISIO_V10 = new ClassIDPredefined("{00021A14-0000-0000-C000-000000000046}", ".vsd", "application/vnd.visio");
        /** Equation Editor 3.0 */
        public static ClassIDPredefined EQUATION_V3 = new ClassIDPredefined("{0002CE02-0000-0000-C000-000000000046}", null, null);
        /** AcroExch.Document */
        public static ClassIDPredefined PDF = new ClassIDPredefined("{B801CA65-A1FC-11D0-85AD-444553540000}", ".pdf", "application/pdf");
        /** Plain Text Persistent Handler **/
        public static ClassIDPredefined TXT_ONLY = new ClassIDPredefined("{5e941d80-bf96-11cd-b579-08002b30bfeb}", ".txt", "text/plain");

        private static Dictionary<String,ClassIDPredefined> LOOKUP = new Dictionary<String,ClassIDPredefined>();

        static ClassIDPredefined()
        {
            LOOKUP.Add(OLE_V1_PACKAGE.externalForm, OLE_V1_PACKAGE);
            LOOKUP.Add(EXCEL_V3.externalForm, EXCEL_V3);
            LOOKUP.Add(EXCEL_V3_CHART.externalForm, EXCEL_V3_CHART);
            LOOKUP.Add(EXCEL_V3_MACRO.externalForm, EXCEL_V3_MACRO);
            LOOKUP.Add(EXCEL_V7.externalForm, EXCEL_V7);
            LOOKUP.Add(EXCEL_V7_WORKBOOK.externalForm, EXCEL_V7_WORKBOOK);
            LOOKUP.Add(EXCEL_V7_CHART.externalForm, EXCEL_V7_CHART);
            LOOKUP.Add(EXCEL_V8.externalForm, EXCEL_V8);
            LOOKUP.Add(EXCEL_V8_CHART.externalForm, EXCEL_V8_CHART);
            LOOKUP.Add(EXCEL_V11.externalForm, EXCEL_V11);
            LOOKUP.Add(EXCEL_V12.externalForm, EXCEL_V12);
            LOOKUP.Add(EXCEL_V12_MACRO.externalForm, EXCEL_V12_MACRO);
            LOOKUP.Add(EXCEL_V12_XLSB.externalForm, EXCEL_V12_XLSB);
            LOOKUP.Add(EXCEL_V14.externalForm, EXCEL_V14);
            LOOKUP.Add(EXCEL_V14_WORKBOOK.externalForm, EXCEL_V14_WORKBOOK);
            LOOKUP.Add(EXCEL_V14_CHART.externalForm, EXCEL_V14_CHART);
            LOOKUP.Add(EXCEL_V14_ODS.externalForm, EXCEL_V14_ODS);
            LOOKUP.Add(WORD_V7.externalForm, WORD_V7);
            LOOKUP.Add(WORD_V8.externalForm, WORD_V8);
            LOOKUP.Add(WORD_V12.externalForm, WORD_V12);
            LOOKUP.Add(WORD_V12_MACRO.externalForm, WORD_V12_MACRO);
            LOOKUP.Add(POWERPOINT_V7.externalForm, POWERPOINT_V7);
            LOOKUP.Add(POWERPOINT_V7_SLIDE.externalForm, POWERPOINT_V7_SLIDE);
            LOOKUP.Add(POWERPOINT_V8.externalForm, POWERPOINT_V8);
            LOOKUP.Add(POWERPOINT_V8_TPL.externalForm, POWERPOINT_V8_TPL);
            LOOKUP.Add(POWERPOINT_V12.externalForm, POWERPOINT_V12);
            LOOKUP.Add(POWERPOINT_V12_MACRO.externalForm, POWERPOINT_V12_MACRO);
            LOOKUP.Add(PUBLISHER_V12.externalForm, PUBLISHER_V12);
            LOOKUP.Add(VISIO_V10.externalForm, VISIO_V10);
            LOOKUP.Add(EQUATION_V3.externalForm, EQUATION_V3);
            LOOKUP.Add(PDF.externalForm, PDF);
            LOOKUP.Add(TXT_ONLY.externalForm, TXT_ONLY);
        }

        private string externalForm;
        private ClassID classId;
        private string fileExtension;
        private string contentType;

        private ClassIDPredefined(string externalForm, string fileExtension, string contentType)
        {
            this.externalForm = externalForm;
            this.fileExtension = fileExtension;
            this.contentType = contentType;
        }

        public ClassID GetClassID()
        {
            // TODO: init classId directly in the constructor when old statics have been removed from ClassID
            if(classId == null)
            {
                classId = new ClassID(externalForm);
            }

            return classId;
        }

        public string FileExtension
        {
            get
            {
                return fileExtension;
            }
        }

        public string ContentType
        {
            get
            {
                return contentType;
            }
        }

        public static ClassIDPredefined Lookup(string externalForm)
        {
            if(LOOKUP.TryGetValue(externalForm, out ClassIDPredefined result))
            {
                return result;
            }
            return null;
        }

        public static ClassIDPredefined Lookup(ClassID classID)
        {
            return (classID == null) ? null : Lookup(classID.ToString());
        }

        public override string ToString()
        {
            return externalForm;
        }
        public override bool Equals(object obj)
        {
            if(obj == null)
            {
                return false;
            }
            if(obj.GetType() != GetType())
            {
                return false;
            }
            ClassIDPredefined other = (ClassIDPredefined)obj;
            return externalForm.Equals(other.externalForm);
        }
        public override int GetHashCode()
        {
            return externalForm.GetHashCode();
        }
    }
}


