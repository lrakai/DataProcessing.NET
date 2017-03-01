using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing
{
    class CsvDataStreamWriter : IDataStreamWriter
    {
        public Stream WriteDataStream(DataTable table)
        {
            var stringBuilder = new StringBuilder();

            var columnHeadings = new List<string>();
            foreach (DataColumn dataColumn in table.Columns)
            {
                columnHeadings.Add(dataColumn.ColumnName);
            }
            stringBuilder.AppendLine(String.Join(",", columnHeadings));

            foreach (DataRow dataRow in table.Rows)
            {
                var rowValues = new List<string>();
                foreach (var item in dataRow.ItemArray)
                {
                    rowValues.Add(item.ToString());
                }
                stringBuilder.AppendLine(String.Join(",", rowValues));
            }
        
            return GenerateStreamFromString(stringBuilder.ToString());
        }

        private static MemoryStream GenerateStreamFromString(string value)
        {
            return new MemoryStream(Encoding.UTF8.GetBytes(value ?? ""));
        }
    }
}
