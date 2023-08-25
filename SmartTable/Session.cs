using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartWriter
{
    public class Session
    {
        public readonly string[] header;
        public readonly int INITIAL_ROW;

        private int currentRow = 0;
        public int CurrentRow => currentRow;

        public Session(string[] header) : this(header, 0) { }
        public Session(string[] header, int initialRow) : this(header, initialRow, initialRow) { }

        private Session(string[] header, int initialRow, int currentRow)
        {
            this.header = header;
            INITIAL_ROW = initialRow;
            this.currentRow = currentRow;
        }

        public Session Clone()
        {
            return new Session(header, INITIAL_ROW, currentRow);
        }

        public Session NextRow(int i = 1)
        {
            if (i < 1)
            {
                throw new ArgumentException("Cannot go back in worksheet by this method.");
            }

            currentRow += i;
            return this;
        }

        public int IndexOf(string columnName)
        {
            return Array.IndexOf(header, columnName) + 1; //EPPlus è 1-indexed
        }

        public bool Contains(string columnName)
        {
            return header.Contains(columnName);
        }

    }
}