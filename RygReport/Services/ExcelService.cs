using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace RygReport.Services;

public class ExcelService
{
    public const int NotFound = -1;

    private IWorkbook _workbook;
    private IFormulaEvaluator? FormulaEvaluator { get; set; }

    public void Read(string fullPath)
    {
        try
        {
            this.FormulaEvaluator = new XSSFFormulaEvaluator(this._workbook);
            this._workbook = WorkbookFactory.Create(new FileStream(fullPath, FileMode.Open, FileAccess.Read));
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            this._workbook = new XSSFWorkbook();
        }
    }

    public ISheet? GetSheet(string name)
    {
        return this._workbook.GetSheet(name);
    }

    public string GetStringCellValue(string sheetName, int row, int col)
    {
        var cell = this._workbook.GetSheet(sheetName).GetRow(row).GetCell(col);

        if (cell == null)
        {
            return "";
        }

        var type = cell?.CellType;
        if (cell?.CellType == CellType.Formula)
        {
            type = this.FormulaEvaluator?.EvaluateFormulaCell(cell);
        }

        switch (type)
        {
            case CellType.Numeric:
                return cell.NumericCellValue.ToString().Trim();
            case CellType.String:
                return cell.StringCellValue.ToString().Trim();
            case CellType.Boolean:
                return cell.BooleanCellValue.ToString().Trim();
            case CellType.Error:
                return cell.ErrorCellValue.ToString().Trim();
            default:
                return "";
        }
    }

    public double GetNumericCellValue(string sheetName, int row, int col)
    {
        return this._workbook.GetSheet(sheetName).GetRow(row).GetCell(col).NumericCellValue;
    }

    public List<int> FindConcurrentValueRows(string sheetName, string searchValue, CellRangeAddress rangeAddress)
    {
        var sheet = this.GetSheet(sheetName);
        var firstRow = rangeAddress.FirstRow;
        var lastRow = rangeAddress.LastRow;
        var firstCol = rangeAddress.FirstColumn;
        var found = false;

        var rows = new List<int>();

        for (var rowIdx = firstRow; rowIdx <= lastRow; rowIdx++)
        {
            var row = sheet.GetRow(rowIdx);
            var cell = row?.GetCell(firstCol);
            if (cell is not { CellType: CellType.String }) continue;
            if (cell.StringCellValue.Equals(searchValue, StringComparison.OrdinalIgnoreCase))
            {
                rows.Add(cell.RowIndex);
                found = true;
            }
            else if (found)
            {
                break;
            }
        }

        return rows;
    }

    public int FindValueInRange(string sheetName, string searchValue, CellRangeAddress rangeAddress)
    {
        var sheet = this.GetSheet(sheetName);
        int firstRow = rangeAddress.FirstRow;
        int lastRow = rangeAddress.LastRow;
        int firstCol = rangeAddress.FirstColumn;
        int lastCol = rangeAddress.LastColumn;

        for (int rowIdx = firstRow; rowIdx <= lastRow; rowIdx++)
        {
            IRow row = sheet.GetRow(rowIdx);
            if (row == null) continue;

            for (int colIdx = firstCol; colIdx <= lastCol; colIdx++)
            {
                ICell cell = row.GetCell(colIdx);
                if (cell != null && cell.CellType == CellType.String)
                {
                    if (cell.StringCellValue.Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                    {
                        return cell.RowIndex;
                    }
                }
            }
        }

        return NotFound;
    }

    public List<int> FindNotEmptyColumns(string sheetName, CellRangeAddress rangeAddress)
    {
        return FindNotEmptyCells(sheetName, rangeAddress).Select(x => x.ColumnIndex).ToList();
    }

    public List<int> FindNotEmptyRows(string sheetName, CellRangeAddress rangeAddress)
    {
        return FindNotEmptyCells(sheetName, rangeAddress).Select(x => x.RowIndex).ToList();
    }

    private List<ICell> FindNotEmptyCells(string sheetName, CellRangeAddress rangeAddress)
    {
        var columns = new List<ICell>();
        var sheet = this.GetSheet(sheetName);

        var firstRow = rangeAddress.FirstRow;
        var lastRow = rangeAddress.LastRow;
        var firstCol = rangeAddress.FirstColumn;
        var lastCol = rangeAddress.LastColumn;

        for (var colIdx = firstCol; colIdx <= lastCol; colIdx++)
        {
            for (var rowIdx = firstRow; rowIdx <= lastRow; rowIdx++)
            {
                var row = sheet.GetRow(rowIdx);
                if (row == null) continue;

                var cell = row.GetCell(colIdx);
                if (!IsEmpty(cell))
                {
                    columns.Add(cell);
                }
            }
        }

        return columns;
    }

    private static bool IsEmpty(ICell? cell)
    {
        if (cell == null)
        {
            return true;
        }

        switch (cell.CellType)
        {
            case CellType.Blank:
                return true;
            case CellType.String:
                return string.IsNullOrEmpty(cell.StringCellValue);
            case CellType.Numeric:
                return cell.NumericCellValue == 0.0;
            case CellType.Boolean:
                return !cell.BooleanCellValue;
            default:
                return false;
        }
    }

    public IEnumerable<string> GetUniqueStringValues(string sheetName, CellRangeAddress rangeAddress)
    {
        var sheet = this.GetSheet(sheetName);
        var uniqueValues = new HashSet<string>();

        var firstRow = rangeAddress.FirstRow;
        var lastRow = rangeAddress.LastRow;
        var firstCol = rangeAddress.FirstColumn;
        var lastCol = rangeAddress.LastColumn;

        for (var rowIdx = firstRow; rowIdx <= lastRow; rowIdx++)
        {
            var row = sheet.GetRow(rowIdx);
            for (var colIdx = firstCol; colIdx <= lastCol; colIdx++)
            {
                var cell = row?.GetCell(colIdx);
                if (cell != null && cell.CellType == CellType.String)
                {
                    uniqueValues.Add(cell.StringCellValue);
                }
            }
        }

        return uniqueValues;
    }

    private static IRow GetOrCreateRow(ISheet sheet, int rowNumber)
    {
        return sheet.GetRow(rowNumber) ?? sheet.CreateRow(rowNumber);
    }

    private Dictionary<string, ISheet> Sheets { get; } = new();

    public ISheet GetOrCreateSheet(string sheetName)
    {
        if (Sheets.TryGetValue(sheetName, out var value))
        {
            return value;
        }

        var sheet = _workbook.GetSheet(sheetName) ?? this._workbook.CreateSheet(sheetName);
        Sheets.Add(sheetName, sheet);
        return sheet;
    }

    public void Save(string fullPath)
    {
        using var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write);
        this._workbook.Write(fs);
    }

    public void WriteCellString(string sheetName, int rowNumber, int colNumber, string value)
    {
        var cell = this.GetOrCreateCell(sheetName, rowNumber, colNumber);
        cell.SetCellValue(value);
    }

    public void WriteCellFormula(string sheetName, int rowNumber, int colNumber, string value)
    {
        try
        {
            var cell = this.GetOrCreateCell(sheetName, rowNumber, colNumber);
            cell.SetCellFormula(value);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
        }
    }

    public void WriteCellNumber(string sheetName, int rowNumber, int colNumber, double value)
    {
        var cell = this.GetOrCreateCell(sheetName, rowNumber, colNumber);
        cell.SetCellValue(value);
        cell.SetCellType(CellType.Numeric);
    }

    private ICell GetOrCreateCell(string sheetName, int rowNumber, int colNumber)
    {
        var sheet = this.GetOrCreateSheet(sheetName);
        var row = GetOrCreateRow(sheet, rowNumber);
        return row.CreateCell(colNumber);
    }

    public static string GetColumnLetter(int col)
    {
        return CellReference.ConvertNumToColString(col);
    }
}