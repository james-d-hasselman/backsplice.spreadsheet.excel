// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class Row : IRow
    {
        internal XlSpreadsheet.Row row;
        private CellsList cells;

        public Row()
        {
            row = new XlSpreadsheet.Row();
            cells = new CellsList(row);
        }

        internal Row(XlSpreadsheet.Row row)
        {
            this.row = row;
            cells = new CellsList(row);
        }

        public double? Height
        {
            get
            {
                var doubleValue = row.Height;
                if (doubleValue is null)
                {
                    return null;
                }

                return row.Height!.Value;
            }
            set
            {
                row.Height = value;
            }
        }

        public IList<ICell> Cells
        {
            get => cells;
            set
            {
                row.RemoveAllChildren<XlSpreadsheet.Cell>();
                foreach (var cell in value)
                {
                    var excelCell = new Cell();
                    excelCell.Value = cell.Value;
                    row.AddChild(excelCell.cell);
                }
                cells = new CellsList(row);
            }
        }

        private class CellsList : IList<ICell>
        {
            private XlSpreadsheet.Row row;

            public CellsList(XlSpreadsheet.Row row)
            {
                this.row = row;
            }

            public ICell this[int index]
            {
                get => new Cell(row.Elements<XlSpreadsheet.Cell>().ElementAt(index));
                set
                {
                    var excelCell = new Cell();
                    excelCell.Value = value.Value;
                    var oldChild = row.Elements<XlSpreadsheet.Cell>().ElementAt(index);
                    row.ReplaceChild(excelCell.cell, oldChild);
                }
            }

            public int Count => row.Elements<XlSpreadsheet.Cell>().Count();

            public bool IsReadOnly => false;

            public void Add(ICell item)
            {
                var excelCell = new Cell();
                excelCell.Value = item.Value;
                row.AddChild(excelCell.cell);
            }

            public void Clear()
            {
                row.RemoveAllChildren<XlSpreadsheet.Cell>();
            }

            public bool Contains(ICell item)
            {
                var result = from cell in row.Elements<XlSpreadsheet.Cell>()
                             where cell.CellValue != null && cell.CellValue.ToString() == item.Value
                             select cell;
                return result.Any();
            }

            public void CopyTo(ICell[] array, int arrayIndex)
            {
                for (int i = 0; i < Count; i++)
                {
                    var cell = row.Elements<XlSpreadsheet.Cell>().ElementAt(i);
                    array.SetValue(new Cell(cell), arrayIndex++);
                }
            }

            public IEnumerator<ICell> GetEnumerator()
            {
                return (IEnumerator<ICell>)(from cell in row.Elements<XlSpreadsheet.Cell>()
                                            select new Cell(cell));
            }

            public int IndexOf(ICell item)
            {
                var cells = row.Elements<XlSpreadsheet.Cell>();
                for (int i = 0; i < cells.Count(); i++)
                {
                    var cell = cells.ElementAt(i);
                    if (cell.CellValue != null && cell.CellValue.ToString() == item.Value)
                    {
                        return i;
                    }
                }
                return -1;
            }

            public void Insert(int index, ICell item)
            {
                var excelCell = new Cell();
                excelCell.Value = item.Value;
                row.InsertAt(excelCell.cell, index);
            }

            public bool Remove(ICell item)
            {
                var excelCell = new Cell();
                excelCell.Value = item.Value;
                if (row.RemoveChild(excelCell.cell) != null)
                {
                    return true;
                }
                return false;
            }

            public void RemoveAt(int index)
            {
                var cell = row.Elements<XlSpreadsheet.Cell>().ElementAt(index);
                row.RemoveChild(cell);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
    }
}
