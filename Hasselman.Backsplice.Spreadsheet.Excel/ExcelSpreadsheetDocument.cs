// SPDX-FileCopyrightText: 2022 James D. Hasselman <james.d.hasselman@gmail.com>
// SPDX-License-Identifier: GPL-3.0-or-later

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class ExcelSpreadsheetDocument : ISpreadsheetDocument
    {
        private SpreadsheetDocument spreadsheetDocument;
        private string filePath;
        private WorkbookPart workbookPart;
        private WorksheetPart worksheetPart;
        private Sheets sheets;
        private bool isEditable;

        public ExcelSpreadsheetDocument()
        {
            this.filePath = "";
            this.spreadsheetDocument = null!;
            this.workbookPart = null!;
            this.worksheetPart = null!;
            sheets = null!;
        }

        public void Open(string filePath, SpreadsheetDocumentMode mode)
        {
            if(mode == SpreadsheetDocumentMode.ReadOnly)
            {
                isEditable = false;
            } else
            {
                isEditable = true;
            }

            if (File.Exists(filePath))
            {
                spreadsheetDocument = SpreadsheetDocument.Open(filePath, isEditable);
            } else
            {
                spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
                workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                sheets = workbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);
            }
        }

        public IReadOnlyList<IWorksheet> Sheets
        {
            get
            {
                var excelWorkseheets = from worksheetPart in workbookPart.WorksheetParts
                                       select new ExcelWorksheet(worksheetPart.Worksheet);
                return excelWorkseheets.ToList().AsReadOnly();
            }
        }

        public void AddSheet(IWorksheet worksheet)
        {
            worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            sheets = workbookPart.Workbook.AppendChild(new Sheets());
            uint sheetId = (uint)workbookPart.Workbook.Sheets.Count();
            var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = UInt32Value.FromUInt32(sheetId), Name = $"Sheet{sheetId}" };
            sheets.Append(sheet);
        }

        public void Close()
        {
            spreadsheetDocument.Close();
        }

        public void RemoveSheet(IWorksheet worksheet)
        {
            var excelWorksheet = worksheet as ExcelWorksheet;
            var sheetId = workbookPart.GetIdOfPart(worksheetPart);
            var sheets = from s in this.sheets as IEnumerable<Sheet>
                        where s.Id == sheetId
                        select s;
            if(sheets.Any())
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
