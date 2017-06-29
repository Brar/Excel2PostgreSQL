using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2PostgreSQL
{
    internal class TransferRow
    {
        private readonly object[] _data;

        public TransferRow(object[] data)
        {
            _data = data;
        }

        public object this[ColumnInfo col]
        {
            get => _data[col.Index];
            set => _data[col.Index] = value;
        }
    }
}
