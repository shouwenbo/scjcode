using Common;

namespace JJBTranslateApp
{
    public class MessageBoxDialogService : IDialogService
    {
        public void Alert(string message)
        {
            MessageBox.Show(message);
        }
    }
}