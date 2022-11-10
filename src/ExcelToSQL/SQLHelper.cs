using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToSQL
{
    static class SQLHelper
    {
        const string NULL = "NULL";
        public static string BuildInsertAdhocSQL(DataTable table)
        {
            StringBuilder rows = new StringBuilder();
            foreach(DataRow row in table.Rows)
            {
                StringBuilder sql = new StringBuilder("INSERT INTO " + table.TableName + " (");
                StringBuilder values = new StringBuilder("VALUES (");
                bool bFirst = true;

                foreach(DataColumn column in table.Columns)
                {
                    //자동 생성은 만들지 않음.
                    if(column.AutoIncrement)
                    {
                        string identityType = null;
                        switch(column.DataType.Name)
                        {
                            case "Int16":
                                identityType = "smallint";
                                break;
                            case "SByte":
                                identityType = "tinyint";
                                break;
                            case "Int64":
                                identityType = "bigint";
                                break;
                            case "Decimal":
                                identityType = "decimal";
                                break;
                            default:
                                identityType = "int";
                                break;
                        }
                    }
                    else
                    {
                        if(bFirst)
                            bFirst = false;
                        else
                        {
                            sql.Append(", ");
                            values.Append(", ");
                        }
                                                                     
                        sql.Append(column.ColumnName);

                        string dataType = null;
                        switch(column.DataType.Name)
                        {
                            case "Int16":
                                dataType = "smallint";
                                break;
                            case "SByte":
                                dataType = "tinyint";
                                break;
                            case "Int32":
                                dataType = "int";
                                break;
                            case "Int64":
                                dataType = "bigint";
                                break;
                            case "Decimal":
                            case "Double":
                            case "Single":
                                dataType = "decimal";
                                break;
                            case "Boolean":
                                dataType = "bool";
                                break;
                            case "String":
                            case "Char":
                            case "DateTime":
                                dataType = "string";
                                break;
                            default:
                                dataType = "string";
                                break;
                        }

                        if(dataType == "string")
                        {
                            string value = row[column.ColumnName].ToString();
                            if (value.ToUpper() ==NULL)
                                values.Append(value);
                            else
                                values.Append($"'{value}'");

                        }
                        else
                            values.Append(row[column.ColumnName]);
                    }
                }
                sql.Append(") ");
                sql.Append(values.ToString());
                sql.Append(")");

                rows.AppendLine(sql.ToString());
            }

            return rows.ToString(); ;
        }

        public static string BuildAllFieldsSQL(DataTable table)
        {
            string sql = "";
            foreach(DataColumn column in table.Columns)
            {
                if(sql.Length > 0)
                    sql += ", ";
                sql += column.ColumnName;
            }
            return sql;
        }

        public static string BuildInsertSQL(DataTable table)
        {
            StringBuilder sql = new StringBuilder("INSERT INTO " + table.TableName + " (");
            StringBuilder values = new StringBuilder("VALUES (");
            bool bFirst = true;
            bool bIdentity = false;
            string identityType = null;

            foreach(DataColumn column in table.Columns)
            {
                if(column.AutoIncrement)
                {
                    bIdentity = true;

                    switch(column.DataType.Name)
                    {
                        case "Int16":
                            identityType = "smallint";
                            break;
                        case "SByte":
                            identityType = "tinyint";
                            break;
                        case "Int64":
                            identityType = "bigint";
                            break;
                        case "Decimal":
                            identityType = "decimal";
                            break;
                        default:
                            identityType = "int";
                            break;
                    }
                }
                else
                {
                    if(bFirst)
                        bFirst = false;
                    else
                    {
                        sql.Append(", ");
                        values.Append(", ");
                    }

                    sql.Append(column.ColumnName);
                    values.Append("@");
                    values.Append(column.ColumnName);
                }
            }
            sql.Append(") ");
            sql.Append(values.ToString());
            sql.Append(")");

            if(bIdentity)
            {
                sql.Append("; SELECT CAST(scope_identity() AS ");
                sql.Append(identityType);
                sql.Append(")");
            }

            return sql.ToString(); ;
        }

        // Creates a SqlParameter and adds it to the command

        public static void InsertParameter(SqlCommand command,
                                             string parameterName,
                                             string sourceColumn,
                                             object value)
        {
            SqlParameter parameter = new SqlParameter(parameterName, value);

            parameter.Direction = ParameterDirection.Input;
            parameter.ParameterName = parameterName;
            parameter.SourceColumn = sourceColumn;
            parameter.SourceVersion = DataRowVersion.Current;

            command.Parameters.Add(parameter);
        }

        // Creates a SqlCommand for inserting a DataRow
        public static SqlCommand CreateInsertCommand(DataRow row)
        {
            DataTable table = row.Table;
            string sql = BuildInsertSQL(table);
            SqlCommand command = new SqlCommand(sql);
            command.CommandType = System.Data.CommandType.Text;

            foreach(DataColumn column in table.Columns)
            {
                if(!column.AutoIncrement)
                {
                    string parameterName = "@" + column.ColumnName;
                    InsertParameter(command, parameterName,
                                      column.ColumnName,
                                      row[column.ColumnName]);
                }
            }
            return command;
        }

        // Inserts the DataRow for the connection, returning the identity
        public static object InsertDataRow(DataRow row, string connectionString)
        {
            SqlCommand command = CreateInsertCommand(row);

            using(SqlConnection connection = new SqlConnection(connectionString))
            {
                command.Connection = connection;
                command.CommandType = System.Data.CommandType.Text;
                connection.Open();
                return command.ExecuteScalar();
            }
        }

        /// <summary>
        /// convert sp_executesql to nomarl query
        /// </summary>
        /// <param name="origSql">sp_executesql</param>
        /// <returns>query text</returns>
        public static string ConvertExecuteSqlToNomarlSql(string origSql)
        {
            var re = new Regex(@"exec*\s*sp_executesql\s+N'([\s\S]*)',\s*N'(@[\s\S]*?)',\s*([\s\S]*)", RegexOptions.IgnoreCase); // 1: the sql, 2: the declare, 3: the setting
            var match = re.Match(origSql);
            if(match.Success)
            {
                var sql = match.Groups[1].Value.Replace("''", "'");
                //var declare = match.Groups[2].Value;
                var setting = match.Groups[3].Value + ',';

                // to deal with comma or single quote in variable values, we can use the variable name to split
                var re2 = new Regex(@"@[^',]*?\s*=");
                var variables = re2.Matches(setting).Cast<Match>().Select(m => m.Value).ToArray();
                var values = re2.Split(setting).Where(s => !string.IsNullOrWhiteSpace(s)).Select(m => m.Trim(',').Trim().Trim(';')).ToArray();
                for(int i = variables.Length - 1; i >= 0; i--)
                {
                    sql = Regex.Replace(sql, "(" + variables[i].Replace("=", "") + ")", values[i], RegexOptions.Singleline);
                }
                return sql;
            }

            return @"Unknown sql query format.";

        }
    }
}
