using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing
{
    public class ExcelCsvDataConverter : DataConverter
    {
        public ExcelCsvDataConverter()
            : base(new ExcelDataStreamReader(), new CsvDataStreamWriter())
        {
        }
    }
}
