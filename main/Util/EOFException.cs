using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Util
{
    public class EOFException : IOException
    {
        public EOFException()
            : base()
        {
        }
        public EOFException(string message)
            : base(message)
        {
        }       
    }
}
