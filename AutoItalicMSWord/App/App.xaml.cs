using Microsoft.Office.Interop.Word;

namespace AutoItalicMSWord.App
{
    public partial class AutoItalicApplication
    {
        public readonly Application WordApplication = new();

        ~AutoItalicApplication()
        {
            foreach (Document document in WordApplication.Documents)
            {
                document.Close(WdSaveOptions.wdDoNotSaveChanges);
            }
            
            WordApplication.Quit();
        }
    }
}