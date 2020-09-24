using System;
using ClosedXML.Excel;
using System.Collections.Generic;

namespace ConvertToExcel
{
    public static class ExcelExtentions
    {
        public static Dictionary<string, string> ReadExcel(string path, int fCell, int fRow)
        {
            try
            {
                using (XLWorkbook workBook = new XLWorkbook(path))
                {
                    IXLWorksheet workSheet = workBook.Worksheet(1);
                    var range = workSheet.Range(workSheet.FirstCellUsed(), workSheet.LastCellUsed());
                    var dict = new Dictionary<string, string>();
                    var i = 0;

                    foreach (var item in range.Rows())
                    {
                        if (i >= fRow)
                            dict.TryAdd(item.Cell(fCell).Value.ToString(), item.Cell(fCell + 2).Value.ToString());
                        i++;
                    }

                    return dict;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

      
    }
}