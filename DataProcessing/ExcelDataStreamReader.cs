using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing
{
    class ExcelDataStreamReader : IDataStreamReader
    {
        public DataTable ReadDataStream(Stream stream)
        {
            var table = new DataTable();
            var workbook = new XlsxReader.XlsxReader(stream);

            // Assume: data in first sheet
            var sheet = workbook[workbook.WorksheetNames.FirstOrDefault()];
            FillColumnValues(table, sheet);

            FillRowValues(table, sheet);

            return table;
        }
        
        private static void FillColumnValues(DataTable table, XlsxReader.XlsxReader.Sheet sheet)
        {
            foreach (var firstRow in sheet.Rows.First())
            {
                table.Columns.Add(firstRow.Key, typeof(string));
            }
        }

        private static void FillRowValues(DataTable table, XlsxReader.XlsxReader.Sheet sheet)
        {
            foreach (var row in sheet.Rows)
            {
                var rowValues = new List<string>();
                foreach (var kv in row)
                {
                    rowValues.Add(kv.Value);
                }
                table.Rows.Add(rowValues.ToArray<object>());
            }
        }
    }
}
