using System.ComponentModel;
using System.Windows.Input;

namespace AutoItalicMSWord.MainWindow
{
    public sealed class MainWindowViewModel : INotifyPropertyChanged
    {
        public MainWindowViewModel()
        {
            var mainWindowModel = new MainWindowModel(this);
            
            Load = new Command(() => { mainWindowModel.Load(); });
            Save = new Command(() => { mainWindowModel.Save(); });
        }
        
        public bool IsLoadButtonEnabled
        {
            get => _isLoadButtonEnabled;
            set
            {
                _isLoadButtonEnabled = value;
                
                OnPropertyChanged(nameof(IsLoadButtonEnabled));
            }
        }

        public bool IsSaveButtonEnabled
        {
            get => _isSaveButtonEnabled;
            set
            {
                _isSaveButtonEnabled = value;
                
                OnPropertyChanged(nameof(IsSaveButtonEnabled));
            }
        }

        public string StatusText
        {
            get => _statusText;
            set
            {
                _statusText = value;
                
                OnPropertyChanged(nameof(StatusText));
            }
        }

        public ICommand Load { get; }
        
        public ICommand Save { get; }

        public event PropertyChangedEventHandler? PropertyChanged;
        
        private bool _isLoadButtonEnabled = true;

        private bool _isSaveButtonEnabled;

        private string _statusText = string.Empty;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}