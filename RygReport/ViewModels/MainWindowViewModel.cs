using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using RygReport.Services;

namespace RygReport.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    [ObservableProperty] private string _status = "";
    public string Greeting { get; } = "RYG Report Generator!";

    public MainWindowViewModel()
    {
        new RgyReportService().Generate();
    }
}