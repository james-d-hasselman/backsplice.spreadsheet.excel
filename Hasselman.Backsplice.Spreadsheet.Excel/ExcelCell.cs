// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

using DocumentFormat.OpenXml.Spreadsheet;

#if DEBUG
using System.Runtime.CompilerServices;
[assembly: InternalsVisibleToAttribute("UnitTests")]
#endif

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class ExcelCell : ICell
    {
        internal Cell cell;

        public ExcelCell() {
            cell = new Cell();
        }

        internal ExcelCell(Cell cell)
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
                cell.CellValue = new CellValue(value);
            }
        }
    }
}
