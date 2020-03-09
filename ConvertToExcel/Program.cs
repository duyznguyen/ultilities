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
            var json = File.ReadAllText(@"C:\Users\admin\Desktop\resources.js");
            json = json.Replace("export default", string.Empty);

            ConvertToExcel(json);
            //CopyKeyValue(json);
            // ReadFileExcel();
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

                    for (var i = 1; i < headers.Count(); i++)
                    {
                        var header = headers[i];

                        if (value.ContainsKey(header))
                        {
                            dataRows.Add(value[header]);
                        }
                    }

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

        private static void CopyKeyValue(string json)
        {
            var myTranslation = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(json);
            foreach(var tl in myTranslation)
            {
                foreach(var t in tl.Value)
                {
                    if(t.Key == "en-us")
                    {
                        tl.Value.Add("en-ie", t.Value);
                        break;
                    }
                }
            }
            //string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var convertToJson = JsonConvert.SerializeObject(myTranslation);
            File.WriteAllText($@"D:\EXCEL_CSV_FILES\Resources_{DateTime.UtcNow:yyyyMMdd}.json", convertToJson, Encoding.UTF8);
        }

        private static void ReadFileExcel()
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"C:\Users\nguye\OneDrive\Desktop\DsCanBoNhanVien.xlsx")))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                var listEmployee = new List<Employee>();
                //var sb = new StringBuilder(); //this is your data
                for (int rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                {
                    var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString()).ToList();
                    if(rowNum > 1)
                    {
                        listEmployee.Add(new Employee
                        {
                            Id = int.Parse(row[0]),
                            Name = row[1],
                            Email = row[2],
                            Department = row[3],
                            Role = row[4],
                            Manager = row[5],
                            ManagerLv1 = row[6],
                        });
                    }
                    //sb.AppendLine(string.Join(",", row));
                }
                Console.WriteLine(listEmployee);
            }
        }


        public class Employee
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string Email { get; set; }
            public string Department { get; set; }
            public string Role { get; set; }
            public string Manager { get; set; }
            public string ManagerLv1 { get; set; }
        }
    }
}
