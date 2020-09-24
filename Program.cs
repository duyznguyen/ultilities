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
        
        private const string EXPORT_DEFAULT = "export default";
        private const string EN_US = "en-us";
        private const string DE_DE = "de-de";
        private const string AR_AE = "ar-ae";
        private const string ZH_CN = "zh-cn";
        private const string ZH_HK = "zh-hk";
        private const string JA_JP = "ja-jp";
        private const string KO_KR = "ko-kr";
        static void Main(string[] args)
        {
            //var jsonResources = File.ReadAllText(@"C:\Users\admin\Desktop\resources.js");
            //jsonResources = jsonResources.Replace("export default", string.Empty);
            
            // var fileExcelPath = @"C:\Users\admin\Desktop\English to Traditional Chinese_holidays v1.0.xlsx";
            
            // ConvertToExcel(jsonResources);
            //var cultureCodes = new[] {"ar-bh", "ar-om", "ar-kw", "ar-ot"};
            //CopyKeyValue(jsonResources, "ar-ae", cultureCodes);
            // ReadFileExcel();
            
            //ImportTranslation(jsonResources, fileExcelPath, ZH_HK);
            string sourceDirectory = @"D:\WorkSpaces\goquo-engine-agoda\wwwroot\multi-sites";
            string targetDirectory = @"C:\Users\admin\Desktop\multisites\";
            
            DirectoryCopy(sourceDirectory, targetDirectory);
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
                excel.SaveAs(new FileInfo($@"D:\EXCEL_CSV_FILES\Translation_Tripbundle-V4_{DateTime.UtcNow:yyyyMMdd}.xlsx"));
            }
        }

        private static void CopyKeyValue(string json, string baseCultureCode, string[] cultureCodes)
        {
            var myTranslation = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(json);
            foreach(var tl in myTranslation)
            {
                foreach(var t in tl.Value)
                {
                    if(t.Key == baseCultureCode)
                    {
                        foreach (var cultureCode in cultureCodes)
                        {
                            tl.Value.Add(cultureCode, t.Value);
                        }
                        break;
                    }
                }
            }
            //string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            ExportToJsonFile(myTranslation, "resources");
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


        private static void ImportTranslation(string jsonResources, string fileExcelPath, string expectedLanguage)
        {
            var excelToDictionary = ExcelExtentions.ReadExcel(fileExcelPath, 1, 1);
            var detectJson = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(jsonResources);

            foreach (var excelDic in excelToDictionary.Keys)
            {
                var expectedValue = excelToDictionary[excelDic];
                try
                {
                    detectJson.TryGetValue(excelDic, out var dictionary);
                    if (dictionary != null && expectedValue != null && expectedValue != "" && expectedValue != " " && expectedValue != "same as English")
                    {
                        dictionary[expectedLanguage] = expectedValue;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
            }

            ExportToJsonFile(detectJson, "new-resources");

            #region CompareLanguage
            // var incorrect = new Dictionary<string, List<string>>();
            // var correct = new Dictionary<string, List<string>>();
            //
            // var missKeyInResource = new List<string>();
            // var missKeys = new List<string>();
            // foreach (var item in detectJson.Values)
            // {
            //     try
            //     {
            //         var lang1 = item[EN_US];
            //         var lang2 = item[ZH_CN];
            //
            //         try
            //         {
            //             if (!excelToDictionary[lang1].Trim().Contains(lang2.Trim(), StringComparison.InvariantCultureIgnoreCase))
            //             {
            //                 incorrect.TryAdd(lang1, new List<string>()
            //                 {
            //                     lang2,
            //                     excelToDictionary[lang1]
            //                 });
            //             }
            //             else
            //             {
            //                 correct.TryAdd(lang1, new List<string>()
            //                 {
            //                     lang2,
            //                     excelToDictionary[lang1]
            //                 });
            //             }
            //         }
            //
            //         catch (Exception ex)
            //         {
            //             missKeyInResource.Add($"{ex.Message}");
            //         }
            //     }
            //     catch (Exception e)
            //     {
            //         missKeys.Add($"{item[EN_US]} | {e.Message}");
            //     }
            // }
            #endregion
        }

        private static void ExportToJsonFile(Dictionary<string, Dictionary<string, string>> myDic, string fileName)
        {
            var convertToJson = JsonConvert.SerializeObject(myDic);
            File.WriteAllText($@"D:\EXCEL_CSV_FILES\{fileName}_{DateTime.UtcNow:yyyyMMdd}.json", convertToJson, Encoding.UTF8);
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
        
        private static void DirectoryCopy(string sourceDirName, string destDirName)
        {
            DirectoryInfo dirSource = new DirectoryInfo(sourceDirName);

            if (!dirSource.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            const string settingsJson = "settings.json";
            const string CSS = "css";
            
            DirectoryInfo[] dirSources = dirSource.GetDirectories();
            Directory.CreateDirectory(destDirName);        
            
            foreach (DirectoryInfo dirSourcesLv1 in dirSources)
            {
                DirectoryInfo dirs1 = dirSourcesLv1.GetDirectories(CSS).First();
                FileInfo fileInfoDirSource = dirSourcesLv1.GetFiles(settingsJson).First();
                string destPath = $@"{destDirName}{dirSourcesLv1.Name}";
                
                FileInfo[] files = dirs1.GetFiles();

                foreach (var file in files)
                {
                    
                    string tempPath = Path.Combine($@"{destPath}\css", file.Name);
                    FileInfo fileInfo = new FileInfo(tempPath);
                    if (!fileInfo.Exists)
                    {
                        DirectoryInfo dirInfo = new DirectoryInfo(tempPath);
                        Directory.CreateDirectory(dirInfo.Parent.ToString());
                    }
                    file.CopyTo(tempPath, true);
                }

                fileInfoDirSource.CopyTo($@"{destPath}\{settingsJson}", true);
            }
            Console.WriteLine("Copy done!");
        }
    }
}
