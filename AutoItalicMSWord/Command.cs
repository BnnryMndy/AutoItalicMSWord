using System;
using System.Windows.Input;

namespace AutoItalicMSWord
{
    public class Command : ICommand
    {
        public Command(Action command)
        {
            _command = command;
        }
        
        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter) => _command();

        public event EventHandler CanExecuteChanged  
        {  
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        } 

        private readonly Action _command;
    }
}