using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace ConvertToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var json = File.ReadAllText(@"D:\Goquo\HappyBooking\appsettings.json");
            json = json.Replace("export default", string.Empty);

            ConvertToExcel(json);
        }

        private static void ConvertToCsv(string json)
        {
            var myTranslation = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(json);
            var headers = new List<string> { "Key" };
            var strBuilder = new StringBuilder();

            headers.AddRange(myTranslation.SelectMany(s => s.Value.Keys).Distinct());
            strBuilder.AppendLine(string.Join(",", headers.Select(x => x)));

            foreach (var (key, value) in myTranslation)
            {
                var dataRow = new List<string> { key };
                dataRow.AddRange(value.Values);
                strBuilder.AppendLine(string.Join(",", dataRow));
            }
            File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + $@"D:\EXCEL_CSV_FILES\Resource -Translation_{DateTime.UtcNow:yyyyMMdd}.csv", strBuilder.ToString(), Encoding.UTF8);
        }

        private static void ConvertToExcel(string json)
        {
            using (var excel = new ExcelPackage())
            {
                var myTranslation = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(json);
                var ws = excel.Workbook.Worksheets.Add("Translation");
                var headers = new List<string> { "Key" };
                headers.AddRange(myTranslation.SelectMany(s => s.Value.Keys).Distinct());
                for (var i = 1; i <= headers.Count(); i++)
                {
                    var header = headers[i-1];
                    ws.Cells[1, i].Value = header;
                }
                
                var rowStart = 2;
                foreach (var (key, value) in myTranslation)
                {
                    var dataRows = new List<string> { key };
                    dataRows.AddRange(value.Values);

                    var colStart = 1;
                    foreach (var data in dataRows)
                    {
                        ws.Cells[rowStart, colStart].Value = data;
                        colStart++;
                    }
                    rowStart++;
                }
                excel.SaveAs(new FileInfo($@"D:\EXCEL_CSV_FILES\Translation_{DateTime.UtcNow:yyyyMMdd}.xlsx"));
            }
        }
    }
}
