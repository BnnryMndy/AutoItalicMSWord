using Microsoft.Office.Interop.Word;

namespace AutoItalicMSWord.Extensions
{
    public static class WordExtensions
    {
        public static void CloseAllDocuments(this Application wordApplication)
        {
            foreach (Document document in wordApplication.Documents)
            {
                document.Close(WdSaveOptions.wdDoNotSaveChanges);
            }
        }
    }
}