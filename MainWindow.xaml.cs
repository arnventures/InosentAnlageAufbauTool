using System.Windows;
using InosentAnlageAufbauTool.ViewModels;

namespace InosentAnlageAufbauTool
{
    public partial class MainWindow : Window
    {
        public MainWindow(MainViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
        }
    }
}
