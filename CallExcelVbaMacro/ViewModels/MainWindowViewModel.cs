using CallExcelVbaMacro.Services;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;

namespace CallExcelVbaMacro.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        private string _argument1;

        private string _argument2;
        public ExcelService _excelService;

        private string _macroName;

        private string _selectedPath;

        public MainWindowViewModel()
        {
            _excelService = new ExcelService();
            SearchFileCommand = new RelayCommand<string>(SearchFileDialog, s => true);
            Argument1SearchFileCommand = new RelayCommand<string>(Argument1SearchFileDialog, s => true);
            Argument2SearchFileCommand = new RelayCommand<string>(Argument2SearchFileDialog, s => true);
            ExecuteCommand = new RelayCommand(ExecuteSearch, () => true);
        }

        public RelayCommand<string> SearchFileCommand { get; }
        public RelayCommand<string> Argument1SearchFileCommand { get; }
        public RelayCommand<string> Argument2SearchFileCommand { get; }
        public RelayCommand ExecuteCommand { get; }

        public string SelectedPath
        {
            get => _selectedPath;
            set
            {
                _selectedPath = value;
                RaisePropertyChanged(nameof(SelectedPath));
            }
        }

        public string MacroName
        {
            get => _macroName;
            set
            {
                _macroName = value;
                RaisePropertyChanged(nameof(MacroName));
            }
        }

        public string Argument1
        {
            get => _argument1;
            set
            {
                _argument1 = value;
                RaisePropertyChanged(nameof(Argument1));
            }
        }

        public string Argument2
        {
            get => _argument2;
            set
            {
                _argument2 = value;
                RaisePropertyChanged(nameof(Argument2));
            }
        }

        private void ExecuteSearch()
        {
            _excelService.RunMacro(SelectedPath, MacroName, Argument1, Argument2);
        }

        private void SearchFileDialog(string fileName)
        {
            SelectedPath = fileName;
        }

        private void Argument1SearchFileDialog(string argument)
        {
            Argument1 = argument;
        }

        private void Argument2SearchFileDialog(string argument)
        {
            Argument2 = argument;
        }
    }
}