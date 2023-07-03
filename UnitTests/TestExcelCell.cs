// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

namespace Hasselman.Backsplice.Spreadsheet.Excel.UnitTests
{
    [TestClass]
    public class TestExcelCell
    {
        [TestMethod]
        public void EmptyCell()
        {
            var referenceCell = new XlSpreadsheet.Cell();
            var excelCell = new Cell();
            var testCell = excelCell.cell;
            Assert.AreEqual(testCell.OuterXml, referenceCell.OuterXml);
        }

        [TestMethod]
        public void NonEmptyCell()
        {
            var referenceCell = new XlSpreadsheet.Cell(new XlSpreadsheet.CellValue("TEST"));
            var excelCell = new Cell();
            excelCell.Value = "TEST";
            var testCell = excelCell.cell;
            Assert.AreEqual(testCell.OuterXml, referenceCell.OuterXml);
        }

        [TestMethod]
        public void UpdateCellValue()
        {
            var excelCell = new Cell();
            excelCell.Value = "HELLO";
            var referenceCell = excelCell.cell;
            excelCell.Value = "WORLD";
            var testCell = excelCell.cell;
            Assert.AreEqual(testCell.OuterXml, referenceCell.OuterXml);
        }
    }
}
