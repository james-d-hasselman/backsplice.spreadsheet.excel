// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

using System.Collections;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class ExcelWorksheet : IWorksheet
    {
        internal Worksheet worksheet;
        private SheetData sheetData;
        private RowList rows;

        public ExcelWorksheet()
        {
            this.worksheet = new Worksheet();
            this.sheetData = new SheetData();
            this.worksheet.AppendChild<SheetData>(this.sheetData);
            rows = new RowList(sheetData);
            LeftHeader = "";
            RightHeader = "";
        }
        internal ExcelWorksheet(Worksheet worksheet)
        {
            this.worksheet = worksheet;
            var headerFooters = worksheet.Descendants<HeaderFooter>();
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
            this.sheetData = this.worksheet.Elements<SheetData>().First();
            rows = new RowList(sheetData);
        }

        public IList<IRow> Rows
        {
            get => rows;
            set
            {
                sheetData.RemoveAllChildren<Row>();
                foreach (var row in value)
                {
                    var excelRow = new ExcelRow();
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
            var worksheetCopy = (Worksheet)worksheet.Clone();
            return new ExcelWorksheet(worksheetCopy);
        }

        private class RowList : IList<IRow>
        {
            private SheetData sheetData;

            public RowList(SheetData sheetData)
            {
                this.sheetData = sheetData;
            }

            public IRow this[int index]
            {
                get => new ExcelRow(sheetData.Elements<Row>().ElementAt(index));
                set
                {
                    var excelRow = new ExcelRow();
                    excelRow.Cells = value.Cells;
                    var oldChild = sheetData.Elements<Row>().ElementAt(index);
                    sheetData.ReplaceChild(excelRow.row, oldChild);
                }
            }

            public int Count => sheetData.Elements<Row>().Count();

            public bool IsReadOnly => false;

            public void Add(IRow item)
            {
                var excelRow = new ExcelRow();
                excelRow.Cells = item.Cells;
                sheetData.AddChild(excelRow.row);
            }

            public void Clear()
            {
                sheetData.RemoveAllChildren<Row>();
            }

            public bool Contains(IRow item)
            {
                var result = from row in sheetData.Elements<Row>()
                             where row == item
                             select row;
                return result.Any();
            }

            public void CopyTo(IRow[] array, int arrayIndex)
            {
                for (int i = 0; i < Count; i++)
                {
                    var row = sheetData.Elements<Row>().ElementAt(i);
                    array.SetValue(new ExcelRow(row), arrayIndex++);
                }
            }

            public IEnumerator<IRow> GetEnumerator()
            {
                return (IEnumerator<IRow>)(from row in sheetData.Elements<Row>()
                                            select new ExcelRow(row));
            }

            public int IndexOf(IRow item)
            {
                var rows = sheetData.Elements<Row>();
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
                var excelRow = new ExcelRow();
                excelRow.Cells = item.Cells;
                sheetData.InsertAt(excelRow.row, index);
            }

            public bool Remove(IRow item)
            {
                var excelRow = new ExcelRow();
                excelRow.Cells = item.Cells;
                if (sheetData.RemoveChild(excelRow.row) != null)
                {
                    return true;
                }
                return false;
            }

            public void RemoveAt(int index)
            {
                var row = sheetData.Elements<Row>().ElementAt(index);
                sheetData.RemoveChild(row);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
    }
}
