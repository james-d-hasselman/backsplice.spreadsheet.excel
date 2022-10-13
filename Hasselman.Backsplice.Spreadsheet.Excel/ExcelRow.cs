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
        private ObservableList<ExcelCell> cells;

        public ExcelRow()
        {
            row = new Row();
            cells = new ObservableList<ExcelCell>();
            cells.ItemAdded += Cells_ItemAdded;
            cells.ItemInserted += Cells_ItemInserted;
            cells.ItemRemoved += Cells_ItemRemoved;
            cells.ItemRemovedAt += Cells_ItemRemovedAt;
            cells.ItemUpdated += Cells_ItemUpdated;
        }

        internal ExcelRow(Row row)
        {
            this.row = row;
            if (row.ChildElements.Count > 0)
            {
                var excelCells = from cell in row.ChildElements as IEnumerable<Cell>
                                 select new ExcelCell(cell);
                cells = new ObservableList<ExcelCell>(excelCells);
            }
            else
            {
                cells = new ObservableList<ExcelCell>();
            }
            cells.ItemAdded += Cells_ItemAdded;
            cells.ItemInserted += Cells_ItemInserted;
            cells.ItemRemoved += Cells_ItemRemoved;
            cells.ItemRemovedAt += Cells_ItemRemovedAt;
            cells.ItemUpdated += Cells_ItemUpdated;
        }

        private void Cells_ItemRemovedAt(object? sender, int index)
        {
            var item = row.Elements<Cell>().ElementAt(index);
            row.RemoveChild(item);
        }

        private void Cells_ItemRemoved(object? sender, ExcelCell item)
        {
            row.RemoveChild(item.cell);
        }

        private void Cells_ItemInserted(object? sender, int index, ExcelCell item)
        {
            row.InsertAt(item.cell, index);
        }

        private void Cells_ItemAdded(object? sender, ExcelCell item)
        {
            row.AddChild(item.cell);
        }

        private void Cells_ItemUpdated(object? sender, int index, ExcelCell item)
        {
            var oldItem = row.ElementAt(index) as Cell;
            row.ReplaceChild<Cell>(item.cell, oldItem);
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
            get
            {
                var cells = new List<ICell>(this.cells);
                return cells;
            }
            set
            {
                cells.Clear();
                // TODO fix
                foreach (var cell in value)
                {
                    cells.Add(new ExcelCell());
                }
            }
        }
    }
}
