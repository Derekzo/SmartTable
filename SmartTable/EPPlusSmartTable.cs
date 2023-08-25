using OfficeOpenXml;
using OfficeOpenXml.Style; 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartWriter
{
    public class EPPlusSmartWriter : IDisposable
    {
        private readonly ExcelWorksheet ws;
        public ExcelWorksheet Worksheet => ws;

        private Session checkPoint;

        private Session session;
        public Session GetSession()
        {
            return session.Clone();
        }

        public EPPlusSmartWriter CheckPoint()
        {
            checkPoint = GetSession();
            return this;
        }

        public EPPlusSmartWriter GoBackToCheckPoint()
        {
            session = checkPoint;
            return this;
        }

        public EPPlusSmartWriter(ExcelWorksheet ws, IEnumerable<string> header, int initialRow = 0)
        {
            this.ws = ws;
            session = new Session(header.ToArray(), initialRow);
        }

        public EPPlusSmartWriter NextRow(int i = 1)
        {
            session.NextRow(i);
            return this;
        }

        private ExcelRange Cell(string columnName)
        {
            return ws.Cells[session.CurrentRow, session.IndexOf(columnName)];
        }

        public EPPlusSmartWriter Insert(string columnName, object value)
        {
            if (session.Contains(columnName))
            {
                Cell(columnName).Value = value ?? "";
            }

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

        public void Dispose()
        {
            ws.Dispose();
        }

        private EPPlusSmartWriter SquaredCell(string column)
        {
            var border = Worksheet.Cells[
                GetSession().CurrentRow, 
                GetSession().IndexOf(column)].Style.Border;

            border.Top.Style = 
                border.Bottom.Style = 
                border.Left.Style = 
                border.Right.Style = ExcelBorderStyle.Thin;

            return this;
        }

        public EPPlusSmartWriter SquaredMargins(int colonne)
        {
            for (int i = 0; i < colonne; i++)
            {
                SquaredCell(GetSession().header[i]);
            }

            return this;
        }

        public EPPlusSmartWriter SquaredMargins(int start, int count)
        {
            for (int i = start; i < start + count; i++)
            {
                SquaredCell(GetSession().header[i]);
            }

            return this;
        }

        public EPPlusSmartWriter SquaredMargins(IEnumerable<string> columns)
        {
            foreach (var col in columns)
            {
                SquaredCell(col);
            }

            return this;
        }

        public EPPlusSmartWriter SquaredMargins()
        {
            foreach (var column in GetSession().header)
            {
                SquaredCell(column);
            }

            return this;
        }
    }
}