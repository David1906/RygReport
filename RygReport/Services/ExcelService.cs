using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace RygReport.Services;

public class ExcelService
{
    public const int NotFound = -1;

    private IWorkbook _workbook;

    public void Read(string fullPath)
    {
        this._workbook = WorkbookFactory.Create(new FileStream(fullPath, FileMode.Open, FileAccess.Read));
    }

    public ISheet? GetSheet(string name)
    {
        return this._workbook.GetSheet(name);
    }

    public string GetStringCellValue(string sheetName, int row, int i)
    {
        return this._workbook.GetSheet(sheetName).GetRow(row).GetCell(i).StringCellValue;
    }

    public int FindValueInRange(string sheetName, string searchValue, CellRangeAddress rangeAddress)
    {
        var sheet = this.GetSheet(sheetName);
        int firstRow = rangeAddress.FirstRow;
        int lastRow = rangeAddress.LastRow;
        int firstCol = rangeAddress.FirstColumn + 1;
        int lastCol = rangeAddress.LastColumn + 1;

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
                        return cell.RowIndex + 1;
                    }
                }
            }
        }

        return NotFound;
    }

    public List<int> FindNotEmptyColumns(string sheetName, CellRangeAddress rangeAddress)
    {
        var columnsWithValue1 = new List<int>();
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
                    columnsWithValue1.Add(colIdx);
                }
            }
        }

        return columnsWithValue1;
    }

    public static bool IsEmpty(ICell? cell)
    {
        if (cell == null)
        {
            return true; // Cell is null, considered empty
        }

        switch (cell.CellType)
        {
            case CellType.Blank:
                return true; // Cell is explicitly blank
            case CellType.String:
                return string.IsNullOrEmpty(cell.StringCellValue); // Check for empty string
            case CellType.Numeric:
                return cell.NumericCellValue == 0.0; // Check for zero value
            case CellType.Boolean:
                return !cell.BooleanCellValue; // Check for false value
            default:
                return false; // Other cell types might not be considered empty
        }
    }
}