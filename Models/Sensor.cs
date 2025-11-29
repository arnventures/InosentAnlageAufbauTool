using CommunityToolkit.Mvvm.ComponentModel;

namespace InosentAnlageAufbauTool.Models
{
    public partial class Sensor : ObservableObject
    {
        [ObservableProperty] private bool isSelected = true;
        [ObservableProperty] private int index;
        [ObservableProperty] private int excelRow;
        [ObservableProperty] private string model = string.Empty;
        [ObservableProperty] private string location = string.Empty;
        [ObservableProperty] private int address;
        [ObservableProperty] private string buzzer = "Enable";
        [ObservableProperty] private int? serial;
        [ObservableProperty] private string status = "Pending";
    }
}
