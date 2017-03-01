using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing
{
    public abstract class DataConverter : IDataConverter
    {
        private readonly IDataStreamReader _reader;

        private readonly IDataStreamWriter _writer;

        protected DataConverter(IDataStreamReader reader, IDataStreamWriter writer)
        {
            _reader = reader;
            _writer = writer;
        }

        public Stream Convert(Stream stream)
        {
            var table = _reader.ReadDataStream(stream);
            return _writer.WriteDataStream(table);
        }
    }
    
}
