using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace SmartTable
{
    public class EPPlusSmartTable : IDisposable
    {
        private readonly ExcelWorksheet ws;
        public ExcelWorksheet Worksheet => ws;

        private Session checkPoint;

        private Session session;
        public Session GetSession()
        {
            return session.Clone();
        }

        public EPPlusSmartTable CheckPoint()
        {
            checkPoint = GetSession();
            return this;
        }

        public EPPlusSmartTable RollBack()
        {
            session = checkPoint;
            return this;
        }

        public EPPlusSmartTable(ExcelWorksheet ws, IEnumerable<string> header, int initialRow = 0)
        {
            this.ws = ws;
            session = new Session(header.ToArray(), initialRow);
        }

        public EPPlusSmartTable NextRow(int i = 1)
        {
            session.NextRow(i);
            return this;
        }

        private ExcelRange Cell(string columnName)
        {
            return ws.Cells[session.CurrentRow, session.IndexOf(columnName)];
        }

        public EPPlusSmartTable Insert(string columnName, object value)
        {
            if (session.Contains(columnName))
                Cell(columnName).Value = value ?? "";

            return this;
        }

        public T GetValue<T>(string columnName)
        {
            return Cell(columnName).GetValue<T>();
        }

        public string GetString(string columnName)
        {
            return Cell(columnName).GetValue<string>();
        }

        public decimal GetDecimal(string columnName)
        {
            return Cell(columnName).GetValue<decimal>();
        }

        public int GetInt(string columnName)
        {
            return Cell(columnName).GetValue<int>();
        }

        public void Dispose() { ws.Dispose(); }
    }
}
