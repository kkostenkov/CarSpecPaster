using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CarSpecPaster
{
    public class XlsReadRequest
    {
        public string SheetName;
        public int KeysColumn;
        public int ValuesColumn;
    }


    class XlsReader
    {
        public XlsReader()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        public Dictionary<string, string> Read(string filePath, XlsReadRequest request)
        {
            var fileExists = File.Exists(filePath);
            if (!fileExists)
            {
                Console.WriteLine($"No file found at {0}", filePath);
                return null;
            }
            
            var data = new Dictionary<string, string>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    Console.WriteLine(string.Format("Read from Excel sheet {0}:", request.SheetName));
                    do
                    {
                        if (!string.Equals(request.SheetName, reader.Name))
                        {
                            continue;
                        }

                        while (reader.Read())
                        {
                            ReadRow(reader, data, request);
                        }
                    } while (reader.NextResult());
                    Console.WriteLine();
                }
            }
            return data;
        }

        private void ReadRow(IExcelDataReader reader, Dictionary<string, string> data, XlsReadRequest request)
        {
            var key = ParseCell(reader, request.KeysColumn);
            if (string.IsNullOrEmpty(key))
            {
                return;
            }
            var val = ParseCell(reader, request.ValuesColumn);
            if (!string.IsNullOrEmpty(val))
            {
                data[key] = val;
                Console.WriteLine(string.Format("{0} : {1}", key, val));
            }
        }

        private string ParseCell(IExcelDataReader reader, int columnToParse)
        {
            //GetFieldType() returns the type of a value in the current row.Always one of the types supported by Excel: 
            //double, int, bool, DateTime, TimeSpan, string, or null if there is no value.
            var fieldType = reader.GetFieldType(columnToParse);

            if (fieldType == null)
            {
                return null;
            }
            string result = null;
            if (fieldType == typeof(double))
            {
                var value = reader.GetDouble(columnToParse);
                result = value.ToString();
            }
            else if (fieldType == typeof(int))
            {
                var value = reader.GetInt32(columnToParse);
                result = value.ToString();
            }
            else if (fieldType == typeof(bool))
            {
                var value = reader.GetBoolean(columnToParse);
                result = value.ToString();
            }
            else if (fieldType == typeof(DateTime))
            {
                var value = reader.GetDateTime(columnToParse);
                result = value.ToString();
            }
            else if (fieldType == typeof(TimeSpan))
            {
                Console.WriteLine("TimeSpan parsing is not implemented");
                //var value = reader.GetTimeS(0);
                //data[key] = value;
            }
            else if (fieldType == typeof(string))
            {
                var value = reader.GetString(columnToParse);
                result = value;
            }
            else
            {
                Console.WriteLine($"unknown type from Excel field: {0}", fieldType);
            }

            return result;
        }
    }
}
