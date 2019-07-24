/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using DocumentFormat.OpenXml;
using System.Windows.Forms;
using System.Text;

namespace G1ANT.Addon.Xlsx.Api
{
    public class ColumnNamesGenerator
    {
        private readonly char[] letters;
        private readonly string[] columnNames;

        public ColumnNamesGenerator()
        {
            letters = Enumerable.Range(0, 26).Select(x => (char)(x + 64)).ToArray();
            columnNames = Enumerable
                .Range(0, ushort.MaxValue)
                .Select(x => CalcString(x))
                .ToArray();
        }
           
        public string[] GetColumnsBetweenInclusive(string column1, string column2)
        {
            var start = column1;
            var end = column2;

            if (start.CompareTo(end) > 0)
            {
                var b = end;
                end = start;
                start = b;
            }

            var result = columnNames
                .SkipWhile(x => x != start)
                .TakeWhile(x => x != end)
                .ToList();

            result.Add(end);

            return result.ToArray();
        }

        private string CalcString(int index)
        {
            StringBuilder sb = new StringBuilder();

            while (index > 0)
            {
                sb.Append(letters[index % letters.Length]);
                index /= letters.Length;
            }

            return new string(sb.ToString().Reverse().ToArray());
        }
    }
    
    public class XlsxWrapper
    {
        /// <summary>
        /// It's responsible for caching string values in workbook
        /// </summary>
        /// <remarks>All members of this class are sensitive to the context of a sheet in XlsxWrapper</remarks>
        private class DataCache
        {
            private struct CellRef
            {
                public string sheetID;
                public string adress;

                public override int GetHashCode()
                {
                    return sheetID.GetHashCode() ^ adress.GetHashCode();
                }
            }

            private readonly XlsxWrapper owner;

            private readonly Dictionary<CellRef, string> adress2value = new Dictionary<CellRef, string>();
            private readonly Dictionary<string, IList<CellRef>> value2adress = new Dictionary<string, IList<CellRef>>();

            public string GetValue(string adress)
            {
                return adress2value[new CellRef() { sheetID = owner.sheet.Id, adress = adress }];
            }

            public IEnumerable<string> GetAdresses(string value)
            {
                return value2adress[value].Where(r => r.sheetID == owner.sheet.Id).Select(r => r.adress);
            }

            public bool CotainsAdress(string adress)
            {
                return adress2value.ContainsKey(new CellRef() { adress = adress, sheetID = owner.sheet.Id });
            }

            public bool ContainsValue(string value)
            {
                return value2adress.ContainsKey(value) && value2adress[value].Where(r => r.sheetID == owner.sheet.Id).Count() > 0;
            }

            public DataCache(XlsxWrapper xlsxWrapper)
            {
                owner = xlsxWrapper;
                WorkbookPart wbPart = xlsxWrapper.spreadsheetDocument.WorkbookPart;

                List<string> sharedStringCache = new List<string>();
                if (wbPart.SharedStringTablePart != null)
                {
                    using (OpenXmlReader shareStringReader = OpenXmlReader.Create(wbPart.SharedStringTablePart))
                    {
                        while (shareStringReader.Read())
                        {
                            if (shareStringReader.ElementType == typeof(SharedStringItem))
                            {
                                SharedStringItem stringItem = (SharedStringItem)shareStringReader.LoadCurrentElement();
                                sharedStringCache.Add(stringItem.Text?.Text ?? string.Empty);
                            }
                        }
                    }

                }

                foreach (WorksheetPart sheetPart in wbPart.WorksheetParts)
                {
                    void AddEntry(CellRef reference, string value)
                    {
                        if (value2adress.ContainsKey(value))
                            value2adress[value].Add(reference);
                        else
                            value2adress[value] = new List<CellRef>() { reference };
                        adress2value[reference] = value;
                    }

                    string sheetID = wbPart.GetIdOfPart(sheetPart);

                    using (OpenXmlReader sheetReader = OpenXmlReader.Create(sheetPart))
                    {
                        while (sheetReader.Read())
                        {
                            if (sheetReader.ElementType == typeof(Cell))
                            {
                                Cell cell = (Cell)sheetReader.LoadCurrentElement();
                                CellRef cellAdress = new CellRef()
                                {
                                    sheetID = sheetID,
                                    adress = cell.CellReference
                                };
                                if ((cell.DataType?.Value ?? CellValues.Error) == CellValues.SharedString)
                                    AddEntry(
                                        cellAdress,
                                        sharedStringCache[Int32.Parse(cell.CellValue.InnerText)]);
                                else
                                    AddEntry(cellAdress, owner.GetStringValue(cell));
                            }
                        }
                    }
                }
            }
        }

        private SpreadsheetDocument spreadsheetDocument = null;
        private Sheet sheet;
        private WorkbookPart wbPart;
        private DataCache dataCache;

        private XlsxWrapper()
        {
        }

        public XlsxWrapper(int id)
        {
            this.Id = id;
        }

        public int Id { get; set; }
        public CellR[] SelectedCells { get; set; }

        public Sheet GetSheetByName(string name)
        {
            var sheets = wbPart.Workbook.Sheets.Cast<Sheet>().ToList();
            return sheets.Find(x => x.Name == name);
        }

        public List<String> GetSheetsNames()
        {
            List<string> names = new List<string>();
            var sheets = wbPart.Workbook.Sheets.Cast<Sheet>().ToList();
            foreach (Sheet sh in sheets)
            {
                names.Add(sh.Name);
            }
            return names;
        }

        public int CountRows()
        {
            //WorksheetPart worksheetPart = wbPart.WorksheetParts.First();
            //SheetData sheetData = sheet.Descendants<SheetData>().First();
            int a = 0;// sheetData.Elements<Row>().Count();

            //IEnumerable<Row> row = sheetData.Elements<Row>();
            //a = row.Count();
            IEnumerable<WorksheetPart> worksheetPart = wbPart.WorksheetParts;
            WorksheetPart wsPart =
             (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            Worksheet worksheet = wsPart.Worksheet;
            //find sheet data
            SheetData sheetData = worksheet.Elements<SheetData>().First();
            // Iterate through every sheet inside Excel sheet

            IEnumerable<Row> row = sheetData.Elements<Row>(); // Get the row IEnumerator
            a = row.Count(); // Will give you the count of rows

            return a;
        }

        // TODO: Implemnt using cached values
        public List<object> GetColumn(string rowSpan, string column)
        {
            string[] startEndtemp = rowSpan.Split(':');
            if (startEndtemp.Length > 2)
                throw new ArgumentException($"Range has to have format of 'start:end'", nameof(rowSpan));

            string startRow = startEndtemp[0].Trim();
            string endRow = startEndtemp[1].Trim();

            int start = string.IsNullOrWhiteSpace(startRow) ? 1 : int.Parse(startRow);
            int end = string.IsNullOrWhiteSpace(endRow) ? CountRows() : int.Parse(endRow);

            IEnumerable<WorksheetPart> worksheetPart = wbPart.WorksheetParts;
            WorksheetPart wsPart =
             (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            Worksheet worksheet = wsPart.Worksheet;
            SheetData data = worksheet.GetFirstChild<SheetData>();

            object[] rows = new object[end - start + 1];

#if DEBUG
            int i = 0;
#endif

            foreach (OpenXmlElement element in data.ChildElements)
            {
                if (element is Row row)
                {
                    if (start <= row.RowIndex && row.RowIndex <= end)
                    {
                        Cell cell = (Cell)row.FirstOrDefault(e =>
                        {
                            if (e is Cell c)
                            {
                                string columnName = new string(c.CellReference.Value.TakeWhile(symbol => char.IsLetter(symbol)).ToArray());

                                if (column == columnName)
                                    return true;
                            }

                            return false;
                        });

                        if (cell != null)
                        {
                            rows[row.RowIndex - start] = GetStringValue(cell);
                        }
                    }
                }
#if DEBUG
                i++;
#endif

            }

            return rows.ToList();
        }

        public void SetValue(int row, string column, string value)
        {
            var Position = FormatInput(column, row);

            WorksheetPart wsPart =
             (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
         Where(c => c.CellReference == Position.ToUpper()).FirstOrDefault();

            if (theCell != null)
            {
                setCellValue(value, theCell);
                spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            }
            else
            {
                Worksheet worksheet = wsPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                Row newrow;

                newrow = CheckForRow(Convert.ToUInt32(row), wsPart);
                theCell = CheckForCell(ColumnNumberToLetter(column), newrow);
                setCellValue(value, theCell);
                spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            }
        }

        public string GetValue(int row, string column)
        {
            string position = FormatInput(column, row);
            if (dataCache.CotainsAdress(position))
            {
                return dataCache.GetValue(position);
            }
            else
            {

                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == position.ToUpper()).FirstOrDefault();

                return GetStringValue(theCell);
            }
        }

        private string GetStringValue(Cell theCell)
        {
            if (theCell != null)
            {
                if (theCell.DataType != null && theCell.DataType.Value == CellValues.SharedString)
                {
                    return wbPart.SharedStringTablePart.SharedStringTable.ElementAt(Int32.Parse(theCell.CellValue.InnerText)).InnerText.ToString();
                }
                else if (theCell.StyleIndex == "3")
                {
                    double oaDateAsDouble;
                    if (double.TryParse(theCell.InnerText.ToString(), out oaDateAsDouble))
                        return DateTime.FromOADate(oaDateAsDouble).ToString();

                }
                else if (theCell.StyleIndex == "6")
                {

                    double intres = 0;
                    if (double.TryParse(theCell.InnerText.ToString(), out intres))
                        return ((intres * 100 + "%").ToString());

                }
                else
                {
                    return theCell?.CellValue?.InnerText?.ToString() ?? string.Empty;
                }
                return string.Empty;
            }
            else
            {
                return string.Empty;
            }
        }

        public string FormatInput(string column, int row)
        {
            var position = string.Empty;
            position += ColumnNumberToLetter(column);
            position += row.ToString();
            return position;
        }
        public int[] FormatInput(string position)
        {
            int[] result = new int[2];
            var lettersOnly = position.TakeWhile(x => !Char.IsDigit(x)).ToArray();
            result[0] = FormatLetterToNumber(lettersOnly);
            var lol = position.SkipWhile(x => !Char.IsDigit(x)).ToArray();
            result[1] = Int32.Parse(new string(lol));
            return result;
        }
        private string ColumnNumberToLetter(string column)
        {
            var position = string.Empty;
            int columnToConvert = 0;
            var newBase = 26;
            if (Int32.TryParse(column, out columnToConvert))
            {
                var baseRange = Enumerable.Range('A', newBase).ToArray();
                do
                {
                    columnToConvert--;
                    position = (char)baseRange[columnToConvert % newBase] + position;
                    columnToConvert = columnToConvert / newBase;
                } while (columnToConvert > 0);
            }
            else
            {
                position += column.ToUpper();
            }
            return position;
        }
        private int FormatLetterToNumber(char[] position)
        {
            var oldBase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int column = 0; //because we have 1-based starting index
            var reversed = position.Reverse().ToArray();
            for (int i = reversed.Length - 1; i >= 0; i--)
            {
                column += (oldBase.IndexOf(reversed[i]) + 1) * (int)Math.Pow(26, i);
            }
            return column;
        }
        private Row CheckForRow(uint index, WorksheetPart wsPart)
        {

            if (wsPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == index).Count() != 0)
            {
                return wsPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == index).First();
            }
            else
            {
                Row row = new Row() { RowIndex = index };
                wsPart.Worksheet.GetFirstChild<SheetData>().Append(row);
                return row;
            }
        }
        private void setCellValue(string val, Cell cell)
        {
            int v = 0;
            decimal d = 0.0000M;
            const int maxCellLength = 32000;
            if (Int32.TryParse(val, out v))
            {
                cell.CellValue = new CellValue(val);
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
            }
            else if (decimal.TryParse(val, out d))
            {
                cell.CellValue = new CellValue(val);
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
            }
            //else if (DateTime.TryParse(val, out DateTime date))
            //{
            //    cell.CellValue = new CellValue(date.ToOADate().ToString());
            //    cell.StyleIndex = 1;
            //}

            else
            {

                if (val.Length >= maxCellLength)
                {
                    cell.CellValue = new CellValue(val.Substring(0, maxCellLength));
                    cell.AddAnnotation("this text has been truncated to 32000 characters due to excel's cell limit");
                }
                else
                {
                    cell.CellValue = new CellValue(val);
                }
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
            }
        }

        public void SelectRange(CellR startCellReference, CellR endCellReference)
        {
            SelectedCells = startCellReference.BuildMatrix(endCellReference);
        }

        public void CopySelectionToClipboard()
        {
            string textValue = "";
            string oldColumn = SelectedCells.FirstOrDefault()?.Column;
            int oldRow = SelectedCells.FirstOrDefault()?.Row ?? 0;

            foreach (var cell in SelectedCells)
            {
                if (oldColumn != cell.Column)
                {
                    textValue += "\t";
                }
                if (oldRow != cell.Row)
                {
                    textValue += "\r\n";
                }

                string position = cell.Address;

                if (dataCache.CotainsAdress(position))
                {
                    textValue += dataCache.GetValue(position);
                }
                else
                {
                    WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
                    Cell theCell = wsPart.Worksheet.Descendants<Cell>()
                        .Where(c => c.CellReference == position.ToUpper())
                        .FirstOrDefault();

                    textValue += GetStringValue(theCell);
                }

                oldColumn = cell.Column;
                oldRow = cell.Row;
            }

            Clipboard.SetText(textValue);
        }

        private Cell CheckForCell(string column, Row row)
        {
            string position = column + row.RowIndex;
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == position).Count() > 0)
            {
                Cell cell = row.Elements<Cell>().Where(c => c.CellReference.Value == position).First();
                return cell;
            }
            else
            {
                Cell refCell = null;
                foreach (Cell cella in row.Elements<Cell>())
                {
                    if (cella.CellReference.Value.Length == position.Length)
                    {
                        if (string.Compare(cella.CellReference.Value, position, true) > 0)
                        {
                            refCell = cella;
                            break;
                        }
                    }
                }
                Cell newCell = new Cell() { CellReference = position };
                row.InsertBefore(newCell, refCell);
                return newCell;
            }
        }

        public void ActivateSheet(string name)
        {
            Sheet foundSheet = GetSheetByName(name);
            sheet = foundSheet ?? throw new InvalidOperationException("Attempt to set null as active sheet");
        }

        public void Create(string filePath)
        {
            using (var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                doc.AddWorkbookPart().AddNewPart<WorksheetPart>().Worksheet = new Worksheet(new SheetData());

                doc.WorkbookPart.Workbook =
                  new Workbook(
                    new Sheets(
                      new Sheet
                      {
                          Id = doc.WorkbookPart.GetIdOfPart(doc.WorkbookPart.WorksheetParts.First()),
                          SheetId = 1,
                          Name = "Sheet 1"
                      }));
                doc.WorkbookPart.Workbook.CalculationProperties = new CalculationProperties();
                doc.Close();
            }
        }

        public bool Open(string filePath, string accessMode = "ReadWrite")
        {
            if (string.IsNullOrEmpty(accessMode))
            {
                accessMode = "ReadWrite";
            }

            FileAccess access;
            if (Enum.TryParse(accessMode, true, out access) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(accessMode), accessMode, "Accessmode specified an invalid value");
            }

            Package spreadsheetPackage = Package.Open(filePath, FileMode.Open, access);
            try
            {
                spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage);
                wbPart = spreadsheetDocument.WorkbookPart;
                ActivateSheet(GetSheetsNames()[0]);
            }
            catch
            {
                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, access))
                {
                    UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                }
                spreadsheetPackage = Package.Open(filePath, FileMode.Open, access);
                spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetPackage);
                wbPart = spreadsheetDocument.WorkbookPart;
                ActivateSheet(GetSheetsNames()[0]);
                remhyp();
            }

            dataCache = new DataCache(this);

            if (spreadsheetDocument != null) return true;
            else return false;
        }

        /// <summary>
        /// Close underlying file and save changes if it was opened with write access.
        /// </summary>
        public void Close()
        {
            try
            {
                spreadsheetDocument.Close();
            }
            catch { }
        }

        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }

        public string Find(string value)
        {
            if (dataCache.ContainsValue(value))
            {
                return dataCache.GetAdresses(value).First();
            }
            return null;
        }

        public void remhyp()
        {
            Uri z = new Uri("http://broken-link/");
            WorksheetPart wsPart =
            (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            var hyperLinks = wsPart.Worksheet.Descendants<Hyperlinks>().First();
            var hyperRel = wsPart.HyperlinkRelationships.Where(c => c.Uri == z).FirstOrDefault();
            foreach (Hyperlink item in hyperLinks)
            {
                if (hyperRel.Id == item.Id)
                {
                    wsPart.DeleteReferenceRelationship(item.Id.ToString());

                    item.Remove();
                }
                if (hyperLinks.Count() == 0)
                {
                    hyperLinks.Remove();
                }
            }
            wsPart.Worksheet.Save();
        }
    }
}

