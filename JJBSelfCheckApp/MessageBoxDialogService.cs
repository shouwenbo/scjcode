using Common;

namespace JJBSelfCheckApp
{
    public class MessageBoxDialogService : IDialogService
    {
        public void Alert(string message)
        {
            MessageBox.Show(message);
        }
    }
}