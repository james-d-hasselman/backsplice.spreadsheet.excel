using DocumentFormat.OpenXml.Spreadsheet;
using Hasselman.Backsplice.Spreadsheet.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTests
{
    [TestClass]
    public class TestExcelRow
    {
        [TestMethod]
        public void SetRowHeight()
        {
            var referenceRow = new Row();
            referenceRow.Height = 30.5;
            var excelRow = new ExcelRow();
            excelRow.Height = 30.5;
            var row = excelRow.row;
            Assert.AreEqual(referenceRow.Height, row.Height);
            Assert.AreEqual(referenceRow.OuterXml, row.OuterXml);
        }

        [TestMethod]
        public void CreateRowWithCell()
        {
            var referenceCell = new Cell(new CellValue("TEST"));
            var referenceRow = new Row(referenceCell);
            var excelCell = new ExcelCell();
            excelCell.Value = "TEST";
            var excelRow = new ExcelRow();
            excelRow.Cells.Add(excelCell);
            var row = excelRow.row;
            Assert.AreEqual(referenceRow.OuterXml, row.OuterXml);
        }
    }
}
