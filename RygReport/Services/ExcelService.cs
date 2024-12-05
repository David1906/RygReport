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

    public string GetStringCellValue(string sheetName, int row, int col)
    {
        return this._workbook.GetSheet(sheetName).GetRow(row).GetCell(col).StringCellValue;
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
                    columnsWithValue1.Add(cell.ColumnIndex);
                }
            }
        }

        return columnsWithValue1;
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
}