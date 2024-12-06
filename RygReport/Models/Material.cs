using System.Collections.Generic;

namespace RygReport.Models;

public class Material
{
    public string PartNumber { get; set; } = "";
    public string Group { get; set; } = "";
    public List<ProductModel> Models { get; set; } = [];
    public string MaterialType => Models.Count <= 1 ? "Single" : "Multiple";
}