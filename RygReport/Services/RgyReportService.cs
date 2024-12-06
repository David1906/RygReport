using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using NPOI.SS.Util;
using RygReport.Models;

namespace RygReport.Services;

public partial class RgyReportService : ObservableObject
{
    const int HorizonWeeks = 26;
    const string WorkbookInputFullPath = @"C:\Users\david\Downloads\Ryg\RYG_PPS3_base.xlsx";
    const string WorkbookOutFullPath = @"C:\Users\david\Downloads\Ryg\RYG_OUT.xlsx";
    const string InfoRefSheet = "Info Referencia";
    const string MmAmlSheet = "MM y AML";
    const string DemandSheet = "Demanda";
    const string DataSheet = "Data";
    const string RygSheet = "Ryg";
    const string RygCleanRange = "A2:AZ10000";
    const string DemandGroupRange = "A3:A10000";

    [ObservableProperty] private string _status = "IDLE";

    private readonly ExcelService _excelService;
    private readonly ExcelService _excelOutService;

    public RgyReportService()
    {
        this._excelService = new ExcelService();

        this._excelOutService = new ExcelService();
    }

    public void Generate()
    {
        this._excelService.Read(WorkbookInputFullPath);
        this._excelOutService.Read(WorkbookOutFullPath);

        // TODO : _excelService.Clean(RygSheet, CellRangeAddress.ValueOf(DemandGroupRange));
        this.Status = "Getting groups...";
        var groups = _excelService
            .GetUniqueStringValues(DemandSheet, CellRangeAddress.ValueOf(DemandGroupRange))
            //.Take(2)
            .ToList();

        var nextRow = 1;
        var item = 1;
        foreach (var group in groups)
        {
            this.Status = $"Processing group [{group}] [{item}/{groups.Count}]...";
            nextRow = this.ProcessMaterialGroup(group, nextRow, item);
            item++;
        }

        this._excelOutService.GetOrCreateSheet(InfoRefSheet);
        this._excelOutService.GetOrCreateSheet(MmAmlSheet);
        this._excelOutService.GetOrCreateSheet(DemandSheet);
        this._excelOutService.GetOrCreateSheet(DataSheet);
        this._excelOutService.GetOrCreateSheet(RygSheet);
        this.Status = "Saving workbook...";
        this._excelOutService.Save(WorkbookOutFullPath);
        this.Status = "IDLE";
        _lastDemandRow = 0;
    }

    private int ProcessMaterialGroup(string materialGroup, int startingRow, int item)
    {
        var materials = this.GetMaterials(materialGroup);
        var startDataRow = startingRow + 1;
        var materialsEndRow = WriteMaterials(startDataRow, item, materials);
        var modelsEndRow = WriteModels(startDataRow, materials);
        var endRow = Math.Max(materialsEndRow, modelsEndRow);
        WriteDemandAnalysis(startDataRow, endRow);
        return endRow + 1;
    }

    private void WriteDemandAnalysis(int startingDataRow, int endDataRow)
    {
        var typeDateColIdx = 16;
        var demandRow = startingDataRow - 1;
        var balanceRow = endDataRow;

        for (var currentRow = startingDataRow; currentRow < endDataRow; currentRow++)
        {
            // Write headers
            var supplyNo = currentRow - startingDataRow + 1;
            _excelOutService.WriteCellString(RygSheet, currentRow, typeDateColIdx,
                $"Supply {supplyNo}{(supplyNo == 1 ? " (Main)" : "")}");

            // Demand Header
            _excelOutService.WriteCellString(RygSheet, demandRow, typeDateColIdx, "Demand");


            if (string.IsNullOrEmpty(_excelOutService.GetStringCellValue(RygSheet, currentRow, 9)))
            {
                continue;
            }

            // Status formula
            _excelOutService.WriteCellFormula(RygSheet, currentRow, 6,
                $"IF(MIN({ExcelService.GetColumnLetter(typeDateColIdx)}{balanceRow + 1}:{ExcelService.GetColumnLetter(typeDateColIdx + HorizonWeeks)}{balanceRow + 1})<0,\"R\",\"G\")");

            // Write Formulas
            for (var colDelta = 1; colDelta < HorizonWeeks; colDelta++)
            {
                var currentCol = typeDateColIdx + colDelta;
                var columnLetter = ExcelService.GetColumnLetter(currentCol);
                var columnLetterBefore = ExcelService.GetColumnLetter(currentCol - 1);

                // Po formula
                _excelOutService.WriteCellFormula(RygSheet, currentRow, currentCol,
                    $"SUMIFS(Data!R:R,Data!Q:Q,$J${currentRow + 1},Data!S:S,\">=\"&{columnLetter}1,Data!S:S,\"<\"&{columnLetter}1+7)");

                // Balance formula
                var balanceFormula =
                    $"{columnLetterBefore}{balanceRow + 1}+SUM({columnLetter}{startingDataRow + 1}:{columnLetter}{endDataRow})-{columnLetter}{demandRow + 1}";

                if (colDelta == 1)
                {
                    balanceFormula =
                        $"SUM(O{startingDataRow + 1}:O{endDataRow})+SUM(R{startingDataRow + 1}:R{endDataRow})-R{demandRow + 1}";
                }

                _excelOutService.WriteCellFormula(RygSheet, balanceRow, typeDateColIdx + colDelta, balanceFormula);


                // Demand formula
                if (currentRow == startingDataRow)
                {
                    _excelOutService.WriteCellFormula(RygSheet, demandRow, typeDateColIdx + colDelta,
                        $"IFERROR(VLOOKUP($J${startingDataRow + 1},{DemandSheet}!$B:$CX, MATCH({columnLetter}1,{DemandSheet}!1:1,0)-1,FALSE),0)");
                }
            }
        }

        _excelOutService.WriteCellString(RygSheet, balanceRow, typeDateColIdx, "Balance");
    }

    private int WriteMaterials(int startingRow, int item, List<Material> materials)
    {
        var currentRow = startingRow;
        foreach (var material in materials)
        {
            var partNumberRange = $"J{currentRow + 1}";
            var model = material.Models.FirstOrDefault(ProductModel.Null);
            _excelOutService.WriteCellString(RygSheet, currentRow, 0, model.Risk);
            _excelOutService.WriteCellFormula(RygSheet, currentRow, 1,
                $"VLOOKUP({partNumberRange},'{InfoRefSheet}'!H:I,2,FALSE)");
            _excelOutService.WriteCellString(RygSheet, currentRow, 2, material.MaterialType);
            _excelOutService.WriteCellNumber(RygSheet, currentRow, 3, item);
            _excelOutService.WriteCellString(RygSheet, currentRow, 4, "Fox-GDL");
            _excelOutService.WriteCellString(RygSheet, currentRow, 5, model.Program);
            _excelOutService.WriteCellFormula(RygSheet, currentRow, 7,
                $"VLOOKUP({partNumberRange},'{MmAmlSheet}'!B:F,5,FALSE)");
            _excelOutService.WriteCellFormula(RygSheet, currentRow, 8,
                $"VLOOKUP({partNumberRange},'{MmAmlSheet}'!I:K,3,FALSE)");
            _excelOutService.WriteCellString(RygSheet, currentRow, 9, material.PartNumber);
            _excelOutService.WriteCellFormula(RygSheet, currentRow, 10,
                $"VLOOKUP({partNumberRange},'{MmAmlSheet}'!B:F,5,FALSE)");
            _excelOutService.WriteCellFormula(RygSheet, currentRow, 14,
                $"SUMIF('{DataSheet}'!G:M,{partNumberRange},'{DataSheet}'!I:I)");
            _excelOutService.WriteCellFormula(RygSheet, currentRow, 15,
                $"SUMIF('{DataSheet}'!Q:W,{partNumberRange},'{DataSheet}'!R:R)");
            currentRow++;
        }

        return currentRow;
    }

    private int WriteModels(int startingRow, List<Material> materials)
    {
        var currentRow = startingRow;
        var hashSet = new HashSet<string>();

        foreach (var material in materials)
        {
            foreach (var model in material.Models)
            {
                if (hashSet.Contains(model.Name)) continue;

                _excelOutService.WriteCellString(RygSheet, currentRow, 11, model.ApnPcba);
                _excelOutService.WriteCellString(RygSheet, currentRow, 12, model.Name);
                _excelOutService.WriteCellString(RygSheet, currentRow, 13,
                    model.Qty.ToString(CultureInfo.InvariantCulture));
                hashSet.Add(model.Name);
                currentRow++;
            }
        }

        return currentRow;
    }

    private static string GetDemandModelsRange(int row) => $"E{row + 1}:Z{row + 1}";

    private int _lastDemandRow;

    private List<Material> GetMaterials(string materialGroup)
    {
        var materialRows =
            _excelService.FindConcurrentValueRows(DemandSheet, materialGroup,
                CellRangeAddress.ValueOf($"A{_lastDemandRow + 1}:A10000"));

        var materials = new List<Material>();
        foreach (var materialRow in materialRows)
        {
            materials.Add(new Material()
            {
                Group = materialGroup,
                PartNumber = this._excelService.GetStringCellValue(DemandSheet, materialRow, 1),
                Models = this.GetModels(materialRow)
            });
        }

        _lastDemandRow = materialRows.Last();

        return materials;
    }

    private List<ProductModel> GetModels(int materialRow)
    {
        var columns =
            this._excelService.FindNotEmptyColumns(DemandSheet,
                CellRangeAddress.ValueOf(GetDemandModelsRange(materialRow)));

        var models = new List<ProductModel>();
        foreach (var column in columns)
        {
            var modelName = this._excelService.GetStringCellValue(DemandSheet, 1, column);
            var model = this.GetSingleModel(modelName);

            if (model == ProductModel.Null) continue;

            model.Qty = this._excelService.GetNumericCellValue(DemandSheet, materialRow, column);
            models.Add(model);
        }

        return models;
    }

    private ProductModel GetSingleModel(string modelName)
    {
        var row = this._excelService.FindValueInRange(InfoRefSheet, modelName, CellRangeAddress.ValueOf("A1:A100"));
        if (row == ExcelService.NotFound)
        {
            return ProductModel.Null;
        }

        return new ProductModel()
        {
            Name = this.GetInfoRefStringCellValue(row, 0),
            Risk = this.GetInfoRefStringCellValue(row, 1),
            Program = this.GetInfoRefStringCellValue(row, 2),
            ApnPcba = this.GetInfoRefStringCellValue(row, 3),
            ApnDescription = this.GetInfoRefStringCellValue(row, 4)
        };
    }

    private string GetInfoRefStringCellValue(int row, int i)
    {
        try
        {
            return this._excelService.GetStringCellValue(InfoRefSheet, row, i);
        }
        catch (Exception)
        {
            return "";
        }
    }
}