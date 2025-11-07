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

namespace NPOI.OpenXml4Net.Util
{
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI.Util;

    using NPOI.XSSF;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    [TestFixture]
    public class TestZipSecureFile
    {
        [Test]
        public void TestThresholdInputStream()
        {
            // This Assert.Fails in Java 10 because our reflection injection of the ThresholdInputStream causes a
            // ClassCastException in ZipFile now
            // The relevant change in the JDK is http://hg.Openjdk.java.net/jdk/jdk10/rev/85ea7e83af30#l5.66
            using(ZipFile thresholdInputStream = new ZipFile(XSSFTestDataSamples.GetSampleFile("template.xlsx").FullName))
            {
                using(ZipSecureFile secureFile = new ZipSecureFile(XSSFTestDataSamples.GetSampleFile("template.xlsx").FullName))
                {
                    IEnumerator entries = thresholdInputStream.GetEnumerator();
                    while(entries.MoveNext())
                    {
                        ZipEntry entry = entries.Current as ZipEntry;
                        
                        using(Stream inputStream = secureFile.GetInputStream(entry))
                        {
                            ClassicAssert.IsTrue(IOUtils.ToByteArray(inputStream).Length > 0);
                        }
                    }
                }

            }
        }
    }
}


