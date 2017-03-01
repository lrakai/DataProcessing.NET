using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing
{
    public interface IDataStreamReader
    {
        DataTable ReadDataStream(Stream stream);
    }
}
