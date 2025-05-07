using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Common
{
    /// <summary>
    /// 字符串扩展
    /// </summary>
    public static class CommentExtension
    {
        public static bool CDTypeValid(this IComment comment, out string msg)
        {
            var type = comment.ClientAnchor.GetType();
            object instance = Activator.CreateInstance(type);
            // {Name = "XSSFAnchor" FullName = "NPOI.XSSF.UserModel.XSSFAnchor"}
            NPOI.XSSF.UserModel.XSSFAnchor x = (NPOI.XSSF.UserModel.XSSFAnchor)instance;
            Console.WriteLine();

            var xssf = comment as XSSFComment;

            

            Console.WriteLine();
            msg = "";
            
            //Regex regex = new Regex(pattern);
            //Match match = regex.Match(input);
            //return match.Success;
            return false;
        }

        public static string ToLogString(this Exception ex)
        {
            return $"{ex.Message} \r\n {ex.InnerException} \r\n {ex.Source} \r\n {ex.StackTrace}";
        }
    }
}