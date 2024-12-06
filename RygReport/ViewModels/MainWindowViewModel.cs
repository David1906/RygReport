using System.IO;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using RygReport.Services;

namespace RygReport.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    [ObservableProperty] private string _status = "";
    public string Greeting { get; } = "RYG Report Generator!";

    public RgyReportService RgyReportService { get; } = new();

    [RelayCommand]
    private async Task Generate()
    {
        await Task.Run(() => RgyReportService.Generate());
    }
}