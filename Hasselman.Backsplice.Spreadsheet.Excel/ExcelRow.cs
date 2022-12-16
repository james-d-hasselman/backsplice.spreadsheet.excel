// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;
using System.Collections.Specialized;

#if DEBUG
using System.Runtime.CompilerServices;
[assembly: InternalsVisibleToAttribute("UnitTests")]
#endif

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class ExcelRow : IRow
    {
        internal Row row;
        private CellsList cells;

        public ExcelRow()
        {
            row = new Row();
            cells = new CellsList(row);
        }

        internal ExcelRow(Row row)
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
                row.RemoveAllChildren<Cell>();
                foreach(var cell in value)
                {
                    var excelCell = new ExcelCell();
                    excelCell.Value = cell.Value;
                    row.AddChild(excelCell.cell);
                }
                cells = new CellsList(row);
            }
        }

        private class CellsList : IList<ICell>
        {
            private Row row;

            public CellsList(Row row)
            {
                this.row = row;
            }

            public ICell this[int index]
            {
                get => new ExcelCell(row.Elements<Cell>().ElementAt(index));
                set
                {
                    var excelCell = new ExcelCell();
                    excelCell.Value = value.Value;
                    var oldChild = row.Elements<Cell>().ElementAt(index);
                    row.ReplaceChild(excelCell.cell, oldChild);
                }
            }

            public int Count => row.Elements<Cell>().Count();

            public bool IsReadOnly => false;

            public void Add(ICell item)
            {
                var excelCell = new ExcelCell();
                excelCell.Value = item.Value;
                row.AddChild(excelCell.cell);
            }

            public void Clear()
            {
                row.RemoveAllChildren<Cell>();
            }

            public bool Contains(ICell item)
            {
                var result = from cell in row.Elements<Cell>()
                             where cell.CellValue != null && cell.CellValue.ToString() == item.Value
                             select cell;
                return result.Any();
            }

            public void CopyTo(ICell[] array, int arrayIndex)
            {
                for (int i = 0; i < Count; i++)
                {
                    var cell = row.Elements<Cell>().ElementAt(i);
                    array.SetValue(new ExcelCell(cell), arrayIndex++);
                }
            }

            public IEnumerator<ICell> GetEnumerator()
            {
                return (IEnumerator<ICell>)(from cell in row.Elements<Cell>()
                                            select new ExcelCell(cell));
            }

            public int IndexOf(ICell item)
            {
                var cells = row.Elements<Cell>();
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
                var excelCell = new ExcelCell();
                excelCell.Value = item.Value;
                row.InsertAt(excelCell.cell, index);
            }

            public bool Remove(ICell item)
            {
                var excelCell = new ExcelCell();
                excelCell.Value = item.Value;
                if (row.RemoveChild(excelCell.cell) != null)
                {
                    return true;
                }
                return false;
            }

            public void RemoveAt(int index)
            {
                var cell = row.Elements<Cell>().ElementAt(index);
                row.RemoveChild(cell);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
    }
}
