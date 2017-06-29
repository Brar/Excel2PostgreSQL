using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NpgsqlTypes;

namespace Excel2PostgreSQL
{
    internal class ColumnInfo
    {
        public ColumnInfo(int index, string name)
        {
            Index = index;
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }
        public int Index { get; }
        public string Name { get; }
        public NpgsqlDbType Type { get; set; } = NpgsqlDbType.Text;
        public bool Nullable { get; set; } = false;
    }
}
