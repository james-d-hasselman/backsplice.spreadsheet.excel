// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

namespace Hasselman.Backsplice.Spreadsheet.Excel.UnitTests
{
    [TestClass]
    public class TestExcelRow
    {
        [TestMethod]
        public void SetRowHeight()
        {
            var referenceRow = new XlSpreadsheet.Row();
            referenceRow.Height = 30.5;
            var excelRow = new Row();
            excelRow.Height = 30.5;
            var row = excelRow.row;
            Assert.AreEqual(referenceRow.Height, row.Height);
            Assert.AreEqual(referenceRow.OuterXml, row.OuterXml);
        }

        [TestMethod]
        public void CreateRowWithCell()
        {
            var referenceCell = new DocumentFormat.OpenXml.Spreadsheet.Cell(new XlSpreadsheet.CellValue("TEST"));
            var referenceRow = new DocumentFormat.OpenXml.Spreadsheet.Row(referenceCell);
            var excelCell = new Cell();
            excelCell.Value = "TEST";
            var excelRow = new Hasselman.Backsplice.Spreadsheet.Excel.Row();
            excelRow.Cells.Add(excelCell);
            var row = excelRow.row;
            Assert.AreEqual(referenceRow.OuterXml, row.OuterXml);
        }

        [TestMethod]
        public void ModifyExistingCell()
        {
            var referenceCell = new DocumentFormat.OpenXml.Spreadsheet.Cell(new XlSpreadsheet.CellValue("World"));
            var referenceRow = new DocumentFormat.OpenXml.Spreadsheet.Row(referenceCell);
            var excelCell = new Hasselman.Backsplice.Spreadsheet.Excel.Cell();
            excelCell.Value = "Hello";
            var excelRow = new Hasselman.Backsplice.Spreadsheet.Excel.Row();
            excelRow.Cells.Add(excelCell);
            excelRow.Cells[0].Value = "World";
            Assert.AreEqual(referenceRow.OuterXml, excelRow.row.OuterXml);
        }
    }
}
