using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Data.OleDb;

namespace PowerExtensions
{
    public static class PowerExtensions
    {
        // Mean
        public static float Mean(this IEnumerable<float> input)
        {
            return input.Average();
        }
        public static double Mean(this IEnumerable<double> input)
        {
            return input.Average();
        }
        public static double Mean(this IEnumerable<int> input)
        {
            return input.Average();
        }
        public static decimal Mean(this IEnumerable<decimal> input)
        {
            return input.Average();
        }

        // Median
        public static float Median(this IEnumerable<float> input)
        {
            return Convert.ToSingle(Median(input.Select(x => (double)x)));
        }
        public static double Median (this IEnumerable<double> input)
        {
            var sortedInput = input.OrderBy(n => n).ToArray();
            var size = sortedInput.Length;
            var mid = size / 2;

            if (size % 2 == 0)
            {
                
                return ((sortedInput.ElementAt(mid - 1) +
                         sortedInput.ElementAt(mid)) / 2);
            }
            else
            {
                return sortedInput.ElementAt(mid);
            }
        }
        public static double Median(this IEnumerable<int> input)
        {
            return Median(input.Select(x => (double)x));
        }
        public static decimal Median(this IEnumerable<decimal> input)
        {
            return Convert.ToDecimal(Median(input.Select(x => (double)x)));
        }

        // Mode
        public static int[] Modes(this IEnumerable<int> input)
        {
            var modes = input
                        .GroupBy(val => val)
                        .Select(kvp => new
                        {
                            Key = kvp.Key,
                            Count = kvp.Count()
                        })
                        .ToList();
            var maxCount = modes.Max(m => m.Count);

            return modes
                    .Where(x => x.Count == maxCount && maxCount > 1)
                    .Select(x => x.Key)
                    .ToArray();
                    
        }
        public static T To<T>(this IConvertible obj)
        {
            return (T)Convert.ChangeType(obj, typeof(T));
        }


        #region System

            #region String

        /// <summary>
        ///     Indicates whether this string is numeric.
        /// </summary>
        public static bool IsNumeric(this string value)
        {
            long retNum;
            return long.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
        }

        /// <summary>
        ///     Returns characters from right of specified length.
        /// </summary>
        /// <param name="length">Max number of charaters to return.</param>
        /// <returns>Returns string from right.</returns>
        public static string Right(this string value, int length)
        {
            return value != null && value.Length > length ? value.Substring(value.Length - length) : value;
        }

        /// <summary>
        ///     Returns characters from left of specified length.
        /// </summary>
        /// <param name="length">Max number of charaters to return.</param>
        /// <returns>Returns string from left.</returns>
        public static string Left(this string value, int length)
        {
            return value != null && value.Length > length ? value.Substring(0, length) : value;
        }

        /// <summary>
        ///     Removes the array of characters from the current string.
        /// </summary>
        /// <param name="chars"></param>
        /// <returns></returns>
        public static string RemoveChars(this string input, params char[] chars)
        {
            return new string(input.Where((ch) => !chars.Contains(ch)).ToArray());
        }

            #endregion

        #endregion

        #region System.IO

            #region Stream

        /// <summary>
        ///     Returns a byte array of the current stream
        /// </summary>
        public static byte[] ConvertToByteArray(this Stream stream)
        {
            var len = Convert.ToInt32(stream.Length);
            byte[] data = new byte[len + 1];

            // Convert to a byte array
            stream.Read(data, 0, len);
            stream.Close();

            return data;
        }
        
            #endregion

        #endregion

        #region System.Text

            #region Stringbuilder

        /// <summary>
        ///     Appends the string returned by a processing a composite format string followed by
        ///     the default line terminator, which contains zero or more format items, to this instance.
        ///     Each format item is replaced by the string representation of a corresponding
        ///     argument in a parameter array using a specified format provider.
        /// </summary>
        /// <param name="format">A composite format string.</param>
        public static StringBuilder AppendLineFormat(this StringBuilder sb, string format, params object[] arguments)
        {
            var strFormat = String.Format(format, arguments);
            sb.AppendLine(strFormat);
            return sb;
        }

            #endregion

        #endregion

        #region System.Data

            #region DataTable

        public static XDocument ToXml(this DataTable dt, string rootName)
        {
            var xdoc = new XDocument
            {
                Declaration = new XDeclaration("1.0", "utf-8", "")
            };
            xdoc.Add(new XElement(rootName));
            foreach (DataRow row in dt.Rows)
            {
                var element = new XElement(dt.TableName);
                foreach (DataColumn col in dt.Columns)
                {
                    element.Add(new XElement(col.ColumnName, row[col].ToString().Trim(' ')));
                }
                if (xdoc.Root != null) xdoc.Root.Add(element);
            }

            return xdoc;
        }

        public static void ToExcel(this DataTable dt, string filePath)
        {
            if (CreateTemplate(filePath))
            {
                // Create the connection to the Excel template
                using (var xlConnection = GetExcelConnection(filePath))
                {
                    // Open the connection
                    xlConnection.Open();

                    // Create the table
                    var xlCommand = GetCreateTableCommand(dt, xlConnection);
                    xlCommand.ExecuteNonQuery();

                    // Insert the rows
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        // Build the insert command
                        xlCommand = GetInsertCommand(dt, i, xlConnection);

                        // Excecute the command
                        xlCommand.ExecuteNonQuery();
                    }
                }
            }
        }

        private static bool CreateTemplate(string filePath)
        {
            // Get the Excel format template to use
            byte[] templateBytes = null;
            string ext = Path.GetExtension(filePath.ToLower());
            switch (ext)
            {
                case ".xls":
                    templateBytes = Properties.Resources.XLS_Template;
                    break;
                case ".xlsx":
                    templateBytes = Properties.Resources.XLSX_Template;
                    break;
                default:
                    return false;
            }

            if (templateBytes == null)
                return false;

            // Write the template file
            using (FileStream templateFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                try
                {
                    templateFile.Write(templateBytes, 0, templateBytes.Length);
                    return true;
                }
                catch (System.IO.IOException)
                {
                    return false;
                }
            }
        }

        private static OleDbConnection GetExcelConnection(string xlPath)
        {
            string _dataProvider = "Microsoft.ACE.OLEDB.12.0";
            string _exProperties = string.Empty;
            string _tblHeaders = "HDR=YES;";
            string _mode = string.Empty;

            // Set the correct extended properties
            switch (Path.GetExtension(xlPath))
            {
                case ".xls":
                    _exProperties = "Excel 8.0";
                    break;
                case ".xlsx":
                    _exProperties = "Excel 12.0 Xml";
                    break;
                case ".xlsb":
                    _exProperties = "Excel 12.0";
                    break;
                case ".xlsm":
                    _exProperties = "Excel 12.0 Macro";
                    break;
                default:
                    return null;
            }

            // Build the connection string
            string returnFormat = "Provider={0}{4};Data Source={1};Extended Properties=\"{2};{3}\"";

            // Return the connection
            return new OleDbConnection(string.Format(returnFormat, _dataProvider, xlPath, _exProperties, _tblHeaders, _mode));
        }
        private static OleDbCommand GetDropTableCommand(string TableName, OleDbConnection connection)
        {
            var returnCommand = connection.CreateCommand();

            // Create a new command
            StringBuilder sb = new StringBuilder();

            // Append the table name
            sb.AppendFormat("DROP TABLE [{0}$]", TableName);

            // Set the command text
            returnCommand.CommandText = sb.ToString();

            // Return the completed command
            return returnCommand;
        }
        private static OleDbCommand GetCreateTableCommand(DataTable dataTable, OleDbConnection connection)
        {
            // Create the command
            var returnCommand = connection.CreateCommand();

            // Create a new command string
            StringBuilder sb = new StringBuilder("CREATE TABLE");

            // Append the table name
            sb.AppendFormat(" [{0}] (", dataTable.TableName);

            // Iterate through each column in the datatable
            foreach (DataColumn col in dataTable.Columns)
            {
                // Append the column name and data type
                sb.AppendFormat("[{0}] {1}, ", col.Caption, GetOleDbType(col.DataType).ToString().ToLower());
            }

            // Replace the last comma with a parenthesis close
            sb = sb.Replace(',', ')', sb.ToString().LastIndexOf(','), 1);

            // Set the command text
            returnCommand.CommandText = sb.ToString();

            // Return the completed command
            return returnCommand;
        }
        private static OleDbCommand GetUpdateCommand(DataTable dataTable, OleDbConnection connection, string rowFilter)
        {
            // Create the return command
            var retCommand = connection.CreateCommand();

            // Build the command string
            var sb = new StringBuilder(string.Format("UPDATE [{0}$] SET ", dataTable.TableName));

            foreach (DataColumn col in dataTable.Columns)
            {
                // Append the command text
                sb.AppendFormat("{0} = ?, ", col.ColumnName);

                // Create the column parameter
                var par = new OleDbParameter
                {
                    ParameterName = col.ColumnName,
                    OleDbType = GetOleDbType(col.DataType),
                    Size = col.MaxLength,
                    SourceColumn = col.ColumnName,
                };

                // Add the parameter to the return command
                retCommand.Parameters.Add(par);
            }

            // Remove the last comma
            sb.Remove(sb.ToString().LastIndexOf(','), 1);

            // Add a where clause if a rowfilter was provided
            if (rowFilter != string.Empty)
                sb.AppendFormat("WHERE {0}", rowFilter);

            // Set the command text
            retCommand.CommandText = sb.ToString();

            // Return the command
            return retCommand;
        }
        private static OleDbCommand GetInsertCommand(DataTable dataTable, int rowIndex, OleDbConnection connection)
        {
            var returnCommand = connection.CreateCommand();

            // Create a new commandstring
            var sb = new StringBuilder();
            sb.AppendFormat("INSERT INTO [{0}$] ", dataTable.TableName);

            // Append the prebuilt column strings
            sb.Append(GetColumnFormat(dataTable));
            sb.Append(" ");

            // Append the values to the value formatting
            sb.AppendFormat(GetValueFormat(dataTable), dataTable.Rows[rowIndex].ItemArray);

            // Set the command text
            returnCommand.CommandText = sb.ToString();

            // Return the updated command
            return returnCommand;
        }
        private static string GetColumnFormat(DataTable dataTable)
        {
            // Create the return string
            var sb = new StringBuilder();

            // Get the columns
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                sb.AppendFormat(" [{0}],", dataTable.Columns[i].ColumnName.Replace(' ', '_'));
            }

            // Remove the last comma
            sb.Remove(sb.Length - 1, 1);

            // Return the columns
            return string.Format("({0})", sb.ToString());
        }
        private static string GetValueFormat(DataTable dataTable)
        {
            // Create a new stringbuilder
            var sb = new StringBuilder();

            // Iterate through the number of columns
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                // Append the numbered column marker
                sb.Append(" {");
                sb.Append(i);
                sb.Append("},");
            }

            // Remove the last comma
            sb.Remove(sb.Length - 1, 1);

            // Return the value format string
            return string.Format("VALUES({0})", sb.ToString());
        }
        private static OleDbType GetOleDbType(Type inputType)
        {
            switch (inputType.FullName)
            {
                // Return the appropriate type
                case "System.Boolean":
                    return OleDbType.Boolean;
                case "System.Int32":
                    return OleDbType.Integer;
                case "System.Single":
                    return OleDbType.Single;
                case "System.Double":
                    return OleDbType.Double;
                case "System.Decimal":
                    return OleDbType.Decimal;
                case "System.String":
                    return OleDbType.Char;
                case "System.Char":
                    return OleDbType.Char;
                case "System.Byte[]":
                    return OleDbType.Binary;
                default:
                    return OleDbType.Variant;
            }
        }

            #endregion

        #endregion

    }
}
