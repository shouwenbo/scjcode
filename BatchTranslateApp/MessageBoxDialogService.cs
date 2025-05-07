using Common;

namespace BatchTranslateApp
{
    public class MessageBoxDialogService : IDialogService
    {
        public void Alert(string message)
        {
            MessageBox.Show(message);
        }
    }
}