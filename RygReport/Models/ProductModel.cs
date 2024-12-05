namespace RygReport.Models;

public class ProductModel
{
    public static ProductModel Null = new NullProductModel();

    public string Name { get; set; } = "";
    public string Risk { get; set; } = "";
    public string Program { get; set; } = "";
    public string ApnPcba { get; set; } = "";
    public string ApnDescription { get; set; } = "";
}

class NullProductModel : ProductModel
{
    public NullProductModel()
    {
        this.Name = "NULL";
        this.Risk = "NULL";
        this.Program = "NULL";
        this.ApnPcba = "NULL";
        this.ApnDescription = "NULL";
    }
}