using System;
using System.IO;
using NPOI.Util;

namespace NPOI.HPSF
{
    public class Filetime
    {
        /**
         * The Windows FILETIME structure holds a date and time associated with a
         * file. The structure identifies a 64-bit integer specifying the
         * number of 100-nanosecond intervals which have passed since
         * January 1, 1601, Coordinated Universal Time (UTC).
         */
        private static BigInteger EPOCH_DIFF = BigInteger.ValueOf(-11644473600000L);
        /** Factor between filetime long and date milliseconds */
        private static BigInteger NANO_100 = BigInteger.ValueOf(1000L * 10L);
        private long fileTime;

        internal Filetime(DateTime date)
        {
            fileTime = date.ToFileTime();
        }
        internal Filetime() { }

        internal void Read(LittleEndianByteArrayInputStream lei)
        {
            fileTime = lei.ReadLong();
        }

        public byte[] ToByteArray()
        {
            byte[] result = new byte[LittleEndianConsts.LONG_SIZE];
            LittleEndian.PutLong(result, 0, fileTime);
            return result;
        }

        public int Write(Stream out1)
        {
            out1.Write(ToByteArray(), 0, LittleEndianConsts.LONG_SIZE);
            return LittleEndianConsts.LONG_SIZE;
        }

        internal DateTime GetJavaValue()
        {
            return DateTime.FromFileTime(fileTime);
        }

        /// <summary>
        /// Converts a Windows FILETIME (in UTC) into a {@link Date} (in UTC).
        /// </summary>
        /// <param name="filetime"></param>
        /// <returns></returns>
        public static DateTime FiletimeToDate(long filetime)
        {
            //long ms_since_16010101 = filetime / NANO_100;
            //long ms_since_19700101 = ms_since_16010101 + EPOCH_DIFF;
            //return new DateTime(ms_since_19700101);
            return DateTime.FromFileTime(filetime);
        }

        /**
         * Converts a {@link Date} into a filetime.
         *
         * @param date The date to be converted
         * @return The filetime
         *
         * @see #filetimeToDate(long)
         */
        public static long DateToFileTime(DateTime date)
        {
            //long ms_since_19700101 = date.getTime();
            //long ms_since_16010101 = ms_since_19700101 - EPOCH_DIFF;
            //return ms_since_16010101 * NANO_100;
            return date.ToFileTime();
        }

        /**
         * Return {@code true} if the date is undefined
         *
         * @param date the date
         * @return {@code true} if the date is undefined
         */
        public static bool IsUndefined(DateTime? date)
        {
            return !date.HasValue || DateToFileTime(date.Value) == 0;
        }

        private static BigInteger twoComplement(long val)
        {
            // for negative BigInteger, top byte is negative
            byte[] contents = {
                (byte)(val < 0 ? 0 : -1),
                (byte)((val >> 56) & 0xFF),
                (byte)((val >> 48) & 0xFF),
                (byte)((val >> 40) & 0xFF),
                (byte)((val >> 32) & 0xFF),
                (byte)((val >> 24) & 0xFF),
                (byte)((val >> 16) & 0xFF),
                (byte)((val >> 8) & 0xFF),
                (byte)(val & 0xFF),
            };

            return new BigInteger(contents);
        }
    }
}