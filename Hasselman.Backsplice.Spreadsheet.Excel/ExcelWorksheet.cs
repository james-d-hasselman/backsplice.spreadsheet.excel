using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Hasselman.Backsplice.Spreadsheet.Excel
{
    public class ExcelWorksheet : IWorksheet
    {
        internal Worksheet worksheet;
        private SheetData sheetData;

        public ExcelWorksheet()
        {
            this.worksheet = new Worksheet();
            this.sheetData = new SheetData();
            this.worksheet.AppendChild<SheetData>(this.sheetData);
            LeftHeader = "";
            RightHeader = "";
        }
        internal ExcelWorksheet(Worksheet worksheet)
        {
            this.worksheet = worksheet;
            var headerFooters = worksheet.Descendants<HeaderFooter>();
            LeftHeader = "";
            RightHeader = "";
            if (headerFooters.Any())
            {
                var headerFooter = headerFooters.First();
                var oddHeader = headerFooter.OddHeader;
                if (oddHeader != null)
                {
                    var headerText = oddHeader.Text;
                    var headerPartSeparators = new string[] { "&L", "&C", "&R" };
                    var headerParts = headerText.Split(headerPartSeparators, StringSplitOptions.TrimEntries);
                    LeftHeader = headerParts[0];
                    RightHeader = headerParts[1];
                }
            }
            this.sheetData = this.worksheet.Elements<SheetData>().First();
        }

        public IList<IRow> Rows {
            get {
                var rowEnumerable = from row in sheetData.ChildElements as IEnumerable<Row>
                                    select new ExcelRow(row);
                return (IList<IRow>)rowEnumerable.ToList();
            }
        }

        public string LeftHeader { get; set; }
        public string RightHeader { get; set; }

        public IList<IColumn> Columns => throw new NotImplementedException();

        public IWorksheet DeepCopy()
        {
            var worksheetCopy = (Worksheet)worksheet.Clone();
            return new ExcelWorksheet(worksheetCopy);
        }
    }
}
