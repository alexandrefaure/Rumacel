using System.Windows;
using CallExcelVbaMacro.ViewModels;
using GalaSoft.MvvmLight.Command;
using Microsoft.Win32;

namespace CallExcelVbaMacro
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainWindowViewModel viewModel;
        public MainWindow()
        {
            InitializeComponent();
            viewModel = DataContext as MainWindowViewModel;
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            var viewModelSearchFileCommand = viewModel?.SearchFileCommand;
            OpenFileDialogAndExecuteCommand(ref viewModelSearchFileCommand);
        }

        private void Argument1SearchButton_Click(object sender, RoutedEventArgs e)
        {
            var viewModelSearchFileCommand = viewModel?.Argument1SearchFileCommand;
            OpenFileDialogAndExecuteCommand(ref viewModelSearchFileCommand);
        }

        private void Argument2SearchButton_Click(object sender, RoutedEventArgs e)
        {
            var viewModelSearchFileCommand = viewModel?.Argument2SearchFileCommand;
            OpenFileDialogAndExecuteCommand(ref viewModelSearchFileCommand);
        }

        private void Argument3SearchButton_Click(object sender, RoutedEventArgs e)
        {
            var viewModelSearchFileCommand = viewModel?.Argument3SearchFileCommand;
            OpenFileDialogAndExecuteCommand(ref viewModelSearchFileCommand);
        }

        private void OpenFileDialogAndExecuteCommand<T>(ref T command) where T:RelayCommand<string>
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true && command.CanExecute(openFileDialog.FileName))
            {
                command.Execute(openFileDialog.FileName);
            }
        }
    }
}
