using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartTable
{
    public static class Extensions
    {
        #region Worksheets
        public static T GetValue<T>(this ExcelWorksheet ws, string cellName)
        {
            return ws.Cells[cellName].GetValue<T>();
        }

        public static object Get(this ExcelWorksheet ws, string cellName)
        {
            return ws.Cells[cellName].Value;
        }

        public static decimal GetDecimal(this ExcelWorksheet ws, string cellName)
        {
            return (decimal)(double)ws.Cells[cellName].Value;
        }

        public static int GetInt(this ExcelWorksheet ws, string cellName)
        {
            var d = (double)ws.Cells[cellName].Value;

            return (int)d;
        }

        public static ExcelWorksheet Insert(this ExcelWorksheet ws, string cellName, object value)
        {
            ws.Cells[cellName].Value = value ?? "";
            return ws;
        }
        #endregion

        #region Workbooks 
        public static ExcelPackage Merge(this IEnumerable<ExcelPackage> packages)
        { 
            var masterPackage = new ExcelPackage(new FileInfo(@"P:\first.xlsx"));

            int i = 1;

            foreach (var package in packages)
            { 
                foreach (var sheet in package.Workbook.Worksheets)
                { 
                    var masterSheet = masterPackage.Workbook.Worksheets.SingleOrDefault(s => sheet.Name == s.Name);

                    string worksheetName = masterSheet == null ? sheet.Name : 
                                           string.Format("{0}_{1}", sheet.Name, DateTime.Now.ToString("yyyyMMddhhssmmm")+ i++);
                    
                    masterPackage.Workbook.Worksheets.Add(worksheetName, sheet);
                }
            }

            return masterPackage;
        }
        #endregion
    }
}
