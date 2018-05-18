using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADODB
{
    public class RecordsetItem
    {
        public RecordsetItem(string pName, object pValue)
        {
            Name = pName;
            Value = pValue;
        }

        public string Name { get; private set; }
        public object Value { get; private set; }
    }
}
