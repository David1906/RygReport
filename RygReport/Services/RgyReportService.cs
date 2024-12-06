using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Util;
using RygReport.Models;

namespace RygReport.Services;

public class RgyReportService
{
    const string InfoRefSheet = "Info Referencia";
    const string DemandSheet = "Demanda";
    private readonly ExcelService _excelService;

    public RgyReportService()
    {
        this._excelService = new ExcelService();
        this._excelService.Read(@"C:\Users\david_ascencio\Downloads\RYG\RYG_PPS3_Week49.xlsm");
    }

    public void Generate()
    {
        this.ProcessMaterialGroup(new Material()
        {
            Group = "TPM-001",
            PartNumber = "1A624J500-600-G"
        });
    }

    private void ProcessMaterialGroup(Material material)
    {
        var materials = this.GetMaterials(material.Group);
        var models = this.GetModels(materials.First());
    }

    private List<Material> GetMaterials(string materialGroup)
    {
        var materials = _excelService.FindValueRows(DemandSheet, materialGroup, CellRangeAddress.ValueOf("A1:A10000"));
        return materials.Select(x => new Material()
            {
                Group = materialGroup,
                PartNumber = this._excelService.GetStringCellValue(DemandSheet, x + 1, 1)
            }
        ).ToList();
    }

    private List<ProductModel> GetModels(Material material)
    {
        var models = new List<ProductModel>();
        foreach (var modelName in this.GetModelNames(material))
        {
            models.Add(this.GetSingleModel(modelName));
        }

        return models;
    }

    private List<string> GetModelNames(Material material)
    {
        var materialRow =
            this._excelService.FindValueInRange(DemandSheet, material.Group, CellRangeAddress.ValueOf("A1:A10000"));
        var columns =
            this._excelService.FindNotEmptyColumns(DemandSheet,
                CellRangeAddress.ValueOf($"E{materialRow + 1}:Z{materialRow + 1}"));

        var modelNames = new List<string>();
        foreach (var column in columns)
        {
            modelNames.Add(this._excelService.GetStringCellValue(DemandSheet, 1, column));
        }

        return modelNames;
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