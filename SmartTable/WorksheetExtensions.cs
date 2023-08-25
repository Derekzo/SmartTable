using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report.Shared.SmartWriter
{
    public static class WorksheetExtensions
    {
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
    }
}