using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxReader = XlsxReader.XlsxReader;

namespace DataProcessing.Test
{
    [TestClass]
    public class ExcelCsvConverterTests
    {
        private IDataConverter _dataConverter;

        [TestInitialize]
        public void Initialize()
        {
            _dataConverter = new ExcelCsvDataConverter();
        }

        [TestMethod]
        public void MultitypeConverts()
        {
            var parent = Directory.GetParent(Environment.CurrentDirectory);
            var excelPath = Path.Combine(parent.Parent.FullName, "Resources", "Multitype.xlsx");
            var csvPath = Path.Combine(parent.Parent.FullName, "Resources", "Multitype.csv");
            var csvResult = File.ReadAllText(csvPath);

            string output;
            using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            {
                using (var csvStream = _dataConverter.Convert(stream))
                {
                    using (var streamReader = new StreamReader(csvStream))
                    {
                        output = streamReader.ReadToEnd();
                    }
                }
            }

            Assert.AreEqual(csvResult, output);
        }

        [TestMethod]
        public void EmptyFileThrows()
        {
            string output;
            Exception expectedException = null;

            try
            {
                using (var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes("")))
                {
                    using (var csvStream = _dataConverter.Convert(memoryStream))
                    {
                        using (var streamReader = new StreamReader(csvStream))
                        {
                            output = streamReader.ReadToEnd();
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                expectedException = exception;
            }

            Assert.IsNotNull(expectedException);
        }
    }
}

