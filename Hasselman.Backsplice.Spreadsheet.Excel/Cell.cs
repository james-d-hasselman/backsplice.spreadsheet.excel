// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class Cell : ICell
    {
        internal XlSpreadsheet.Cell cell;

        public Cell()
        {
            cell = new XlSpreadsheet.Cell();
        }

        internal Cell(XlSpreadsheet.Cell cell)
        {
            this.cell = cell;
        }

        public string Value
        {
            get
            {
                if (cell.CellValue == null)
                {
                    return "";
                }

                return cell.CellValue.Text;
            }
            set
            {
                cell.CellValue = new XlSpreadsheet.CellValue(value);
            }
        }
    }
}
