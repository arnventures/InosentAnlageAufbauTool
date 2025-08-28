using CommunityToolkit.Mvvm.ComponentModel;

namespace InosentAnlageAufbauTool.Models
{
    public partial class Led : ObservableObject
    {
        [ObservableProperty] private bool isSelected = true;
        [ObservableProperty] private int index;
        [ObservableProperty] private string model = string.Empty;     // optional display only
        [ObservableProperty] private string location = string.Empty;  // optional display only
        [ObservableProperty] private int address;
        [ObservableProperty] private int timeOut;                     // 0 or 180
        [ObservableProperty] private string status = "Pending";
    }
}
