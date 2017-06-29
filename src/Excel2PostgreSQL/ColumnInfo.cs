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
        public string Name { get; set; }
        public NpgsqlDbType Type { get; set; }
        public bool Nullable { get; set; } = false;
    }
}
