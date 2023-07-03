// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class SpreadsheetDocument : ISpreadsheetDocument
    {
        private XlPackage.SpreadsheetDocument spreadsheetDocument;
        private string filePath;
        private XlPackage.WorkbookPart workbookPart;
        private XlPackage.WorksheetPart worksheetPart;
        private XlSpreadsheet.Sheets sheets;
        private bool isEditable;

        public SpreadsheetDocument()
        {
            this.filePath = "";
            this.spreadsheetDocument = null!;
            this.workbookPart = null!;
            this.worksheetPart = null!;
            sheets = null!;
        }

        public void Open(string filePath, SpreadsheetDocumentMode mode)
        {
            if (mode == SpreadsheetDocumentMode.ReadOnly)
            {
                isEditable = false;
            }
            else
            {
                isEditable = true;
            }

            if (File.Exists(filePath))
            {
                spreadsheetDocument = XlPackage.SpreadsheetDocument.Open(filePath, isEditable);
            }
            else
            {
                spreadsheetDocument = XlPackage.SpreadsheetDocument.Create(filePath, Xl.SpreadsheetDocumentType.Workbook);
                workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new XlSpreadsheet.Workbook();
                worksheetPart = workbookPart.AddNewPart<XlPackage.WorksheetPart>();
                worksheetPart.Worksheet = new XlSpreadsheet.Worksheet(new XlSpreadsheet.SheetData());
                sheets = workbookPart.Workbook.AppendChild(new XlSpreadsheet.Sheets());
                var sheet = new XlSpreadsheet.Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);
            }
        }

        public IReadOnlyList<IWorksheet> Sheets
        {
            get
            {
                var excelWorkseheets = from worksheetPart in workbookPart.WorksheetParts
                                       select new Worksheet(worksheetPart.Worksheet);
                return excelWorkseheets.ToList().AsReadOnly();
            }
        }

        public void AddSheet(IWorksheet worksheet)
        {
            worksheetPart = workbookPart.AddNewPart<XlPackage.WorksheetPart>();
            worksheetPart.Worksheet = new XlSpreadsheet.Worksheet(new XlSpreadsheet.SheetData());
            sheets = workbookPart.Workbook.AppendChild(new XlSpreadsheet.Sheets());
            uint sheetId = (uint)workbookPart.Workbook.Sheets.Count();
            var sheet = new XlSpreadsheet.Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = Xl.UInt32Value.FromUInt32(sheetId), Name = $"Sheet{sheetId}" };
            sheets.Append(sheet);
        }

        public void Close()
        {
            spreadsheetDocument.Close();
        }

        public void RemoveSheet(IWorksheet worksheet)
        {
            var excelWorksheet = worksheet as Worksheet;
            var sheetId = workbookPart.GetIdOfPart(worksheetPart);
            var sheets = from s in this.sheets as IEnumerable<XlSpreadsheet.Sheet>
                         where s.Id == sheetId
                         select s;
            if (sheets.Any())
            {
                sheets.First().Remove();
            }

            excelWorksheet.worksheet.Remove();
        }

        public void Save()
        {
            spreadsheetDocument.Save();
        }

        public void SaveAs(string path)
        {
            spreadsheetDocument.SaveAs(path);
        }
    }
}
