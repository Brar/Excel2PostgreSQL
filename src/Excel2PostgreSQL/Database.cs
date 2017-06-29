using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Npgsql;
using NpgsqlTypes;

namespace Excel2PostgreSQL
{
    internal static class Database
    {
        private static NpgsqlConnection _connection;
        private static NpgsqlCommandBuilder cb = new NpgsqlCommandBuilder();

        public static void Connect(string dbName = null, string host = "::1", int port = 5432, string user = null,
            string password = null)
        {
            Disconnect();
            var cs = CreateConnectionString(dbName, host, port, user, password);
            _connection = new NpgsqlConnection(cs);
            _connection.Open();
        }

        public static void Disconnect()
        {
            _connection?.Close();
        }

        public static void CreateDb(string dbName, bool dropIfExists = false)
        {
            if (dropIfExists)
            {
                DropDb(dbName);
            }

            using (var command = _connection.CreateCommand())
            {
                command.CommandText = $"CREATE DATABASE {cb.QuoteIdentifier(dbName)};";
                command.ExecuteNonQuery();
            }
        }

        public static void AddTableWithData(Table table)
        {
            CreateTable(table);
            string cmd = CreateCopyCommand(table);
            using (var writer =
                _connection.BeginBinaryImport(cmd))
            {
                foreach (object[] values in table.Data)
                {
                    writer.StartRow();
                    for (int i = 0; i < values.Length; i++)
                    {
                        var type = table.ColumnInfos[i].Type;
                        var val = ConvertValue(values[i], type);
                        writer.Write(val, type);
                    }
                }
            }
        }

        private static object ConvertValue(object value, NpgsqlDbType type)
        {
            if (value == null)
            {
                return null;
            }
            switch (type)
            {
                case NpgsqlDbType.Money:
                case NpgsqlDbType.Double:
                case NpgsqlDbType.Boolean:
                    return value;
                case NpgsqlDbType.Smallint:
                    return Convert.ToInt16((double)value);
                case NpgsqlDbType.Integer:
                    return Convert.ToInt32((double)value);
                case NpgsqlDbType.Bigint:
                    return Convert.ToInt64((double)value);
                case NpgsqlDbType.Text:
                    return value.ToString();
                case NpgsqlDbType.Date:
                case NpgsqlDbType.Timestamp:
                case NpgsqlDbType.Time:
                    return DateTime.FromOADate((double)value);
            }
            throw new NotSupportedException();
        }

        private static void CreateTable(Table table)
        {
            using (var command = _connection.CreateCommand())
            {
                StringBuilder sb = new StringBuilder($"CREATE TABLE {cb.QuoteIdentifier(table.Name)} (");
                foreach (var columnInfo in table.ColumnInfos)
                {
                    sb.Append(cb.QuoteIdentifier(columnInfo.Name));
                    sb.Append(" ");
                    sb.Append(GetTypeString(columnInfo.Type));
                    sb.Append(columnInfo.Nullable ? " NULL," : " NOT NULL,");
                }
                sb.Length -= 1;
                sb.Append(");");
                command.CommandText = sb.ToString();
                command.ExecuteNonQuery();
            }
        }

        private static string GetTypeString(NpgsqlDbType type)
        {
            switch (type)
            {
                case NpgsqlDbType.Boolean:
                    return "bool";
                case NpgsqlDbType.Smallint:
                    return "int2";
                case NpgsqlDbType.Integer:
                    return "int4";
                case NpgsqlDbType.Bigint:
                    return "int8";
                case NpgsqlDbType.Double:
                    return "float8";
                case NpgsqlDbType.Money:
                    return "money";
                case NpgsqlDbType.Text:
                    return "text";
                case NpgsqlDbType.Date:
                    return "date";
                case NpgsqlDbType.Timestamp:
                    return "timestamp";
                case NpgsqlDbType.Time:
                    return "time";
            }
            throw new NotSupportedException();
        }

        private static string CreateCopyCommand(Table table)
        {
            StringBuilder sb = new StringBuilder("COPY ");
            sb.Append(cb.QuoteIdentifier(table.Name));
            sb.Append(" (");
            foreach (var columnInfo in table.ColumnInfos)
            {
                sb.Append(cb.QuoteIdentifier(columnInfo.Name));
                sb.Append(", ");
            }
            sb.Length -= 2;
            sb.Append(") FROM STDIN (FORMAT BINARY);");

            return sb.ToString();
        }

        private static string CreateConnectionString(string dbName, string host, int port, string user,
            string password)
        {
            NpgsqlConnectionStringBuilder csb =
                new NpgsqlConnectionStringBuilder
                {
                    ApplicationName = "Excel2PostgreSQL",
                    Pooling = false,
                    Host = host ?? throw new ArgumentNullException(nameof(host)),
                    Port = port,
                    Username = user ?? Environment.UserName
                };
            if (dbName != null)
            {
                csb.Database = dbName;
            }
            if (password == null)
            {
                csb.IntegratedSecurity = true;
            }
            else
            {
                csb.Password = password;
            }
            return csb.ToString();
        }

        private static void DropDb(string dbName)
        {
            using (var command = _connection.CreateCommand())
            {
                command.CommandText = $"DROP DATABASE IF EXISTS {cb.QuoteIdentifier(dbName)};";
                command.ExecuteNonQuery();
            }
        }
    }
}
