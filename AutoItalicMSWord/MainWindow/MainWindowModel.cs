using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using AutoItalicMSWord.App;
using AutoItalicMSWord.Extensions;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using Application = System.Windows.Application;

namespace AutoItalicMSWord.MainWindow
{
    public class MainWindowModel
    {
        public MainWindowModel(MainWindowViewModel mainWindowViewModel)
        {
            _mainWindowViewModel = mainWindowViewModel;
        }

        public void Load()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Word documents (*.docx)|*.docx|All files (*.*)|*.*",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            DisableButtons();

            new Thread(() => LoadFileAndProcess(dialog.FileName)).Start();
        }

        public void Save()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Word documents (*.docx)|*.docx",
                AddExtension = true
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            DisableButtons();

            new Thread(() => SaveFile(dialog.FileName)).Start();
        }

        private void LoadFileAndProcess(string filePath)
        {
            var filename = Path.GetFileName(filePath);
            
            _mainWindowViewModel.StatusText = $"Loading file {filename}";

            _app.WordApplication.CloseAllDocuments();

            try
            {
                _app.WordApplication.Documents.Open(filePath);
            }
            catch (COMException exception)
            {
                SendError(exception.Message, "Open file error", false);

                return;
            }
            
            _mainWindowViewModel.StatusText = $"Processing file {filename}";

            var count = _app.WordApplication.ActiveDocument.Words
                .Cast<Range>()
                .Where(word => _englishWordsRegex.IsMatch(word.Text))
                .Select(range => range.Italic = 1)
                .Count();
            
            SendSuccess($"File {filename} loaded and {count} words processed");
        }

        private void SaveFile(string filePath)
        {
            var filename = Path.GetFileName(filePath);
            
            _mainWindowViewModel.StatusText = $"Saving file {filename}";
            
            try
            {
                _app.WordApplication.ActiveDocument.SaveAs(filePath);
            }
            catch (COMException exception)
            {
                SendError(exception.Message, "Save file error", true);

                return;
            }

            SendSuccess($"File saved as {filename}");
        }

        private void SendError(string text, string title, bool isSaveEnabled)
        {
            MessageBox.Show(text, title, MessageBoxButton.OK, MessageBoxImage.Error);
            
            _mainWindowViewModel.IsLoadButtonEnabled = true;
            _mainWindowViewModel.IsSaveButtonEnabled = isSaveEnabled;

            _mainWindowViewModel.StatusText = string.Empty;
        }

        private void SendSuccess(string statusText)
        {
            _mainWindowViewModel.IsLoadButtonEnabled = true;
            _mainWindowViewModel.IsSaveButtonEnabled = true;

            _mainWindowViewModel.StatusText = statusText;
        }

        private void DisableButtons()
        {
            _mainWindowViewModel.IsLoadButtonEnabled = false;
            _mainWindowViewModel.IsSaveButtonEnabled = false;
        }

        private readonly MainWindowViewModel _mainWindowViewModel;
        
        private readonly AutoItalicApplication _app = (Application.Current as AutoItalicApplication)!;
        
        private readonly Regex _englishWordsRegex = new(@"\w*[a-zA-Z]\w*");
    }
}