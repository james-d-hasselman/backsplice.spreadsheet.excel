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
    public class TestExcelCell
    {
        [TestMethod]
        public void EmptyCell()
        {
            var referenceCell = new Cell();
            var excelCell = new ExcelCell();
            var testCell = excelCell.cell;
            Assert.AreEqual(testCell.OuterXml, referenceCell.OuterXml);
        }

        [TestMethod]
        public void NonEmptyCell()
        {
            var referenceCell = new Cell(new CellValue("TEST"));
            var excelCell = new ExcelCell();
            excelCell.Value = "TEST";
            var testCell = excelCell.cell;
            Assert.AreEqual(testCell.OuterXml, referenceCell.OuterXml);
        }

        [TestMethod]
        public void UpdateCellValue()
        {
            var excelCell = new ExcelCell();
            excelCell.Value = "HELLO";
            var referenceCell = excelCell.cell;
            excelCell.Value = "WORLD";
            var testCell = excelCell.cell;
            Assert.AreEqual(testCell.OuterXml, referenceCell.OuterXml);
        }
    }
}
