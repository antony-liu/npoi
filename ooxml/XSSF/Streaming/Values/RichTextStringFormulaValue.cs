using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.XSSF.Streaming.Values
{
    internal class RichTextStringFormulaValue : FormulaValue
    {
        IRichTextString _preEvaluatedValue;
        public override CellType GetFormulaType()
        {
            return CellType.String;
        }
        public void SetPreEvaluatedValue(IRichTextString value)
        {
            _preEvaluatedValue=value;
        }
        public IRichTextString GetPreEvaluatedValue()
        {
            return _preEvaluatedValue;
        }
    }
}
