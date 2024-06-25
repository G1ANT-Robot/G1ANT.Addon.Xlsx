/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.Xlsx
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Windows.Forms;

namespace G1ANT.Addon.Xlsx.Api
{
    public class XlsxWrapper
    {
        private XLWorkbook workbook;
        private IXLWorksheet sheet;

        public XlsxWrapper(int id)
        {
            this.Id = id;
            workbook = CreateWorkbook();
            ActivateSheet(GetSheetsNames()[0]);
        }

        public int Id { get; set; }
        public IXLRange SelectedCells { get; private set; }
        public IXLWorksheet ActiveSheet
        {
            get
            {
                if (sheet == null)
                    throw new ApplicationException("Sheet has not been selected.");
                return sheet;
            }
        }


        public IXLWorksheet GetSheetByName(string name)
        {
            return workbook?.Worksheets.Worksheet(name);
        }

        public List<string> GetSheetsNames()
        {
            return workbook?.Worksheets.Select(x => x.Name).ToList();
        }

        public int CountRows()
        {
            var lastRow = ActiveSheet.LastRowUsed();
            return lastRow != null ? lastRow.RowNumber() : 0;
        }

        public void SetValue(int row, string column, string value)
        {
            ActiveSheet.Cell(row, column).Value = value;
        }

        public object GetValue(int row, string column)
        {
            return ActiveSheet.Cell(row, column).Value;
        }

        public void SelectRange(int startRow, string startColumn, int endRow, string endColumn)
        {
            var startCell = ActiveSheet.Cell(startRow, startColumn);
            var endCell = ActiveSheet.Cell(endRow, endColumn);
            SelectedCells = ActiveSheet.Range(startCell, endCell);
        }

        public void CopySelectionToClipboard()
        {
            string textValue = "";
            if (SelectedCells == null)
                return;

            foreach (var row in SelectedCells.Rows())
            {
                if (!string.IsNullOrEmpty(textValue))
                    textValue += "\r\n";
                textValue += string.Join("\t", row.Cells().Select(x => x.Value));
            }
            Clipboard.SetText(textValue);
        }

        public void PasteFromClipboard()
        {
            if (SelectedCells == null || !SelectedCells.Cells().Any())
            {
                throw new ArgumentException("Attempt to paste text into null selection");
            }

            SelectedCells.Cells().First().Value = Clipboard.GetText();
        }

        public Tuple<System.Drawing.Color?, System.Drawing.Color?> GetCellColor(int row, string column)
        {
            try
            {
                System.Drawing.Color? backgroundColor = null;
                System.Drawing.Color? fontColor = null;

                var cell = ActiveSheet.Cell(row, column);
                backgroundColor = cell.Style.Fill.BackgroundColor.Color;
                fontColor = cell.Style.Font.FontColor.Color;

                return new Tuple<System.Drawing.Color?, System.Drawing.Color?>(backgroundColor, fontColor);
            }
            catch
            {
                throw new ArgumentException("Could not read color from given cell.");
            }
        }

        public void SetCellBackgroundColor(int row, string column, System.Drawing.Color color)
        {
            try
            {
                var cell = ActiveSheet.Cell(row, column);
                cell.Style.Fill.BackgroundColor = XLColor.FromColor(color);
            }
            catch
            {
                throw new ArgumentException("Could not set color of given cell.");
            }
        }

        public void SetCellFontColor(int row, string column, System.Drawing.Color color)
        {
            try
            {
                var cell = ActiveSheet.Cell(row, column);
                cell.Style.Font.FontColor = XLColor.FromColor(color);
            }
            catch
            {
                throw new ArgumentException("Could not set color of given cell.");
            }
        }

        public void ActivateSheet(string name)
        {
            var foundSheet = GetSheetByName(name);
            sheet = foundSheet ?? throw new InvalidOperationException("Attempt to set null as active sheet");
        }

        private XLWorkbook CreateWorkbook()
        {
            var doc = new XLWorkbook();
            doc.AddWorksheet("Sheet 1");
            return doc;
        }

        public void Create(string filePath)
        {
            using (var doc = CreateWorkbook())
            {
                doc.SaveAs(filePath);
            }
        }

        public bool Open(string filePath, string accessMode = "ReadWrite")
        {
            workbook = new XLWorkbook(filePath);
            ActivateSheet(GetSheetsNames()[0]);
            return ActiveSheet!=null;
        }

        /// <summary>
        /// Close underlying file and save changes if it was opened with write access.
        /// </summary>
        public void Close()
        {
            try
            {
                workbook?.Save();
            }
            catch { }
            finally
            {
                sheet = null;
                workbook = null;
            }
        }

        public void Save()
        {
            workbook.Save();
        }

        public IEnumerable<IXLAddress> Find(string value, bool inSelection, bool ignoreCase = false)
        {
            var compareOptions = CompareOptions.None;
            if (ignoreCase)
                compareOptions |= CompareOptions.IgnoreCase;
            var result = ActiveSheet.Search(value, compareOptions).AsEnumerable();
            if (inSelection)
                result = result.Where(x => SelectedCells.Contains(x));
            return result.Select(x => x.Address);
        }
    }
}

