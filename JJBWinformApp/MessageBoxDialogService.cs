using Common;

namespace JJBWinformApp
{
    public class MessageBoxDialogService : IDialogService
    {
        public void Alert(string message)
        {
            MessageBox.Show(message);
        }
    }
}