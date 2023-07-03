// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

using System.Collections;

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class Worksheet : IWorksheet
    {
        internal XlSpreadsheet.Worksheet worksheet;
        private XlSpreadsheet.SheetData sheetData;
        private RowList rows;

        public Worksheet()
        {
            this.worksheet = new XlSpreadsheet.Worksheet();
            this.sheetData = new XlSpreadsheet.SheetData();
            this.worksheet.AppendChild<XlSpreadsheet.SheetData>(this.sheetData);
            rows = new RowList(sheetData);
            LeftHeader = "";
            RightHeader = "";
        }
        internal Worksheet(XlSpreadsheet.Worksheet worksheet)
        {
            this.worksheet = worksheet;
            var headerFooters = worksheet.Descendants<XlSpreadsheet.HeaderFooter>();
            LeftHeader = "";
            RightHeader = "";
            if (headerFooters.Any())
            {
                var headerFooter = headerFooters.First();
                var oddHeader = headerFooter.OddHeader;
                if (oddHeader != null)
                {
                    var headerText = oddHeader.Text;
                    var headerPartSeparators = new string[] { "&L", "&C", "&R" };
                    var headerParts = headerText.Split(headerPartSeparators, StringSplitOptions.TrimEntries);
                    LeftHeader = headerParts[0];
                    RightHeader = headerParts[1];
                }
            }
            this.sheetData = this.worksheet.Elements<XlSpreadsheet.SheetData>().First();
            rows = new RowList(sheetData);
        }

        public IList<IRow> Rows
        {
            get => rows;
            set
            {
                sheetData.RemoveAllChildren<XlSpreadsheet.Row>();
                foreach (var row in value)
                {
                    var excelRow = new Row();
                    excelRow.Cells = row.Cells;
                    sheetData.AddChild(excelRow.row);
                }
                rows = new RowList(sheetData);
            }
        }

        public string LeftHeader { get; set; }
        public string RightHeader { get; set; }

        public IList<IColumn> Columns => throw new NotImplementedException();

        public IWorksheet DeepCopy()
        {
            var worksheetCopy = (XlSpreadsheet.Worksheet)worksheet.Clone();
            return new Worksheet(worksheetCopy);
        }

        private class RowList : IList<IRow>
        {
            private XlSpreadsheet.SheetData sheetData;

            public RowList(XlSpreadsheet.SheetData sheetData)
            {
                this.sheetData = sheetData;
            }

            public IRow this[int index]
            {
                get => new Row(sheetData.Elements<XlSpreadsheet.Row>().ElementAt(index));
                set
                {
                    var excelRow = new Row();
                    excelRow.Cells = value.Cells;
                    var oldChild = sheetData.Elements<XlSpreadsheet.Row>().ElementAt(index);
                    sheetData.ReplaceChild(excelRow.row, oldChild);
                }
            }

            public int Count => sheetData.Elements<XlSpreadsheet.Row>().Count();

            public bool IsReadOnly => false;

            public void Add(IRow item)
            {
                var excelRow = new Row();
                excelRow.Cells = item.Cells;
                sheetData.AddChild(excelRow.row);
            }

            public void Clear()
            {
                sheetData.RemoveAllChildren<XlSpreadsheet.Row>();
            }

            public bool Contains(IRow item)
            {
                var result = from row in sheetData.Elements<XlSpreadsheet.Row>()
                             where row == item
                             select row;
                return result.Any();
            }

            public void CopyTo(IRow[] array, int arrayIndex)
            {
                for (int i = 0; i < Count; i++)
                {
                    var row = sheetData.Elements<XlSpreadsheet.Row>().ElementAt(i);
                    array.SetValue(new Row(row), arrayIndex++);
                }
            }

            public IEnumerator<IRow> GetEnumerator()
            {
                return (IEnumerator<IRow>)(from row in sheetData.Elements<XlSpreadsheet.Row>()
                                           select new Row(row));
            }

            public int IndexOf(IRow item)
            {
                var rows = sheetData.Elements<XlSpreadsheet.Row>();
                for (int i = 0; i < rows.Count(); i++)
                {
                    var row = rows.ElementAt(i);
                    if (row == item)
                    {
                        return i;
                    }
                }
                return -1;
            }

            public void Insert(int index, IRow item)
            {
                var excelRow = new Row();
                excelRow.Cells = item.Cells;
                sheetData.InsertAt(excelRow.row, index);
            }

            public bool Remove(IRow item)
            {
                var excelRow = new Row();
                excelRow.Cells = item.Cells;
                if (sheetData.RemoveChild(excelRow.row) != null)
                {
                    return true;
                }
                return false;
            }

            public void RemoveAt(int index)
            {
                var row = sheetData.Elements<XlSpreadsheet.Row>().ElementAt(index);
                sheetData.RemoveChild(row);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
    }
}
