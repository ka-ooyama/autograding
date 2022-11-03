using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Bibliography;

//https://qiita.com/gushwell/items/5c689e86014105313017

namespace Autograding
{
    public record CellInfo
    {
        public int Row { get; set; }

        public string Col { get; set; }

        public string Value { get; set; }
    }

    sealed class MyExcelBook : IDisposable
    {
        private MyExcelBook() { }

        public void Dispose() => _document?.Dispose();

        private SpreadsheetDocument _document;
        private WorkbookPart _workbookpart;
        private WorksheetPart _worksheetPart;
        private Sheets _sheets;

        public static MyExcelBook CreateBook(string filepath)
        {
            var book = new MyExcelBook();
            // Excelのドキュメントを新規に作成する
            book._document = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // ドキュメントにワークブックを追加する
            book._workbookpart = book._document.AddWorkbookPart();
            book._workbookpart.Workbook = new Workbook();

            // ワークブックにシートの入れ物を追加する
            book._sheets = book._workbookpart.Workbook.AppendChild<Sheets>(new Sheets());
            return book;
        }

        public void Save()
        {
            _workbookpart.Workbook.Save();
            _document.Close();
        }

        public void CreateSheet(string sheetname)
        {
            // シート作成のための準備
            _worksheetPart = _workbookpart.AddNewPart<WorksheetPart>();
            _worksheetPart.Worksheet = new Worksheet(new SheetData());

            // シートを作成し追加する
            var max = (uint)(_sheets.Count() + 1);
            var sheet = new Sheet
            {
                Id = _workbookpart.GetIdOfPart(_worksheetPart),
                SheetId = max,
                Name = sheetname
            };
            _sheets.Append(sheet);
        }

        // シートを選択する
        public void SelectSheet(string sheetName)
        {
            // workbookから指定したシートを見つける
            var theSheet = _workbookpart.Workbook.Descendants<Sheet>().
                Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
                throw new ArgumentException("sheetName");
            //  WorkSheetPartを取り出す。これが以降の操作で必要
            _worksheetPart = (WorksheetPart)(_workbookpart.GetPartById(theSheet.Id));
        }

        // セルを作成。存在していたらそのセルを返す
        public Cell CreateCell(string columnName, uint rowIndex)
        {
            // SheetDataを見つける
            var worksheet = _worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            // 行を見つける
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                //  行がなければ追加する
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            // その行の指定したカラムのセルを取り出す
            var cellReference = columnName + rowIndex;
            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
            if (cell != null)
                return cell;
            // セルがなければセルを追加する 
            var refCell = row.Elements<Cell>()
                .FirstOrDefault(c => string.Compare(c.CellReference.Value, cellReference, true) > 0);
            var newCell = new Cell()
            {
                CellReference = cellReference
            };
            row.InsertBefore(newCell, refCell);
            return newCell;
        }

        // 指定したセルに値を設定する
        public void SetValueAtCell(Cell cell, string text)
        {
            cell.CellValue = new CellValue(text);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }

        // column, rowで指定したセルに値を設定する。
        public void SetValue(string text, string columnName, uint rowIndex)
        {
            Cell cell = CreateCell(columnName, rowIndex);
            SetValueAtCell(cell, text);
        }

        // 既存のエクセルファイルをオープンする
        public static MyExcelBook Open(string filePath)
        {
            var obj = new MyExcelBook();
            obj._document = SpreadsheetDocument.Open(filePath, false);
            obj._workbookpart = obj._document.WorkbookPart;
            return obj;
        }

        // 指定したセルオブジェクトを取得する
        public Cell GetCell(string columnName, uint rowIndex)
        {
            string addressName = columnName + rowIndex;
            return _worksheetPart.Worksheet.Descendants<Cell>()
                            .Where(c => c.CellReference == addressName)
                            .FirstOrDefault();
        }

        // セルの値を取得する　(column, row指定)
        public string GetCellValue(string columnName, uint rowIndex)
        {
            return GetCellValue(GetCell(columnName, rowIndex));
        }

        // セルの値を取得する　(Cell指定)
        public string GetCellValue(Cell cell)
        {
            if (cell == null)
                return null;
            // 式かどうか判定
            if (cell.CellFormula != null)
                // 式ならTextプロパティを返す。
                return cell.CellValue.Text;
            var value = cell.InnerText;
            if (cell.DataType == null)
                return value;
            // セルのデータタイプによって処理を振り分ける
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:  // 文字列
                                               // 単一の SharedStringTablePart への参照を取得する必要がある
                    var stringTable = _workbookpart
                        .GetPartsOfType<SharedStringTablePart>()
                        .FirstOrDefault();
                    if (stringTable != null)
                    {
                        // valueはindexを表している。SharedStringTableから要素を取得
                        var element = stringTable.SharedStringTable
                            .ElementAt(int.Parse(value));
                        // そのInnterTextがセルの値
                        value = element.InnerText;
                    }
                    break;
                case CellValues.Boolean:    // ブール値
                    value = (value != "0").ToString().ToLower();
                    break;
            }
            return value;
        }

        public IEnumerable<CellInfo[]> GetAllCellValues()
        {
            if (_worksheetPart == null)
            {
                Console.WriteLine("WorksheetPart Not Found !!");
                yield break;
            }
            var ws = _worksheetPart.Worksheet;
            foreach (var row in ws.Descendants<Row>())
            {
                var list = new List<CellInfo>();
                // 注意:値が空のセルがあると列挙から除外される。
                // これに対応するには、cell.CellReferenceの値を見て
                // セルの位置を確認する必要がある。
                foreach (Cell cell in row)
                {
#if true
                    string str = Regex.Replace(cell.CellReference.Value, @"[^a-zA-Z]", "");
                    int columnNumber = StringExtensions.ToColumnNumber(str);
                    while (columnNumber > list.Count + 1)
                    {
                        list.Add(new CellInfo());
                    }
#endif
                    list.Add(ToCellInfo(cell));
                }
                yield return list.ToArray();
            }
        }

        private CellInfo ToCellInfo(Cell cell)
        {
            return new CellInfo
            {
                Col = Regex.Match(cell.CellReference.Value, @"^[A-Z]+").Value,
                Row = int.Parse(Regex.Match(cell.CellReference.Value, @"[0-9]+$").Value),
                Value = GetCellValue(cell)
            };
        }
    }
}
