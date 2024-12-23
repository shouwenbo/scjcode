using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Common
{
    /// <summary>
    /// 字符串扩展
    /// </summary>
    public static class StringExtension
    {
        public static bool IsMatchRegex(this string input, string pattern)
        {
            Regex regex = new Regex(pattern);
            Match match = regex.Match(input);
            return match.Success;
        }

        public static bool IsPositiveInteger(this string input) => IsMatchRegex(input, @"^\d+$");

        public static bool IsAllKorean(this string input) => IsMatchRegex(input, @"^[\p{IsHangulSyllables}]+$");
        
        public static bool IsKoreanOrSpace(this string input) => IsMatchRegex(input, @"^[\p{IsHangulSyllables}\s]+$");

        [Obsolete]
        public static bool IsSixDateFormat(this string input)
        {
            string pattern = @"^\d{6}$";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(input);

            if (match.Success)
            {
                int year = int.Parse(input.Substring(0, 2));
                int month = int.Parse(input.Substring(2, 2));
                int day = int.Parse(input.Substring(4, 2));

                if (year >= 0 && year <= 99 && month >= 1 && month <= 12 && day >= 1 && day <= 31)
                {
                    try
                    {
                        DateTime date = new DateTime(year + 1900, month, day);
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                }
            }

            return false;
        }

        public static bool IsEightDateFormat(this string input)
        {
            string pattern = @"^\d{8}$";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(input);

            if (match.Success)
            {
                int year = int.Parse(input.Substring(0, 4));
                int month = int.Parse(input.Substring(4, 2));
                int day = int.Parse(input.Substring(6, 2));

                if (year >= 1900 && month >= 1 && month <= 12 && day >= 1 && day <= 31)
                {
                    try
                    {
                        DateTime date = new DateTime(year, month, day);
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                }
            }

            return false;
        }

        public static bool IsActualDateFormat(this string input)
        {
            // return DateTime.TryParseExact(input, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out _);
            string pattern = @"^\d{4}-\d{2}-\d{2}\((\+|-)\)$";
            Regex regex = new Regex(pattern);

            Match match = regex.Match(input);

            if (match.Success)
            {
                string[] parts = input.Split(new[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 2)
                {
                    string datePart = parts[0];
                    string signPart = parts[1];

                    if ((signPart == "+") || (signPart == "-"))
                    {
                        if (DateTime.TryParseExact(datePart, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime date))
                        {
                            return true;
                        }
                    }
                }

                return false;
            }

            return false;
        }
        
        public static bool IsKoreanAddress(this string input) => IsMatchRegex(input, @"^[가-힣A-Za-z0-9\s()（）-]+$");

        public static bool IsChineseAddress(this string input) => IsMatchRegex(input, @"^[\u4e00-\u9fa5A-Za-z0-9\s()（）-]+$");

        public static int CountSpaces(this string input)
        {
            MatchCollection matches = Regex.Matches(input.Trim(), @"\s+");
            return matches.Count;
        }

        // 410607更新 电话格式更新位 +86 000-0000-0000
        // public static bool IsPhoneNumberAndFormatValid(this string input) => IsMatchRegex(input, @"^86\)1[3-9]\d{1}-\d{4}-\d{4}$") || IsMatchRegex(input, @"^31\)6-\d{4}-\d{4}$");
        public static bool IsPhoneNumberAndFormatValid(this string input) => IsMatchRegex(input, @"^\+86 \d{3}-\d{4}-\d{4}$") || IsMatchRegex(input, @"^\+31 6-\d{4}-\d{4}$");


        public static bool IsNoneOrValidEmail(this string input)
        {
            if (string.Equals(input, "없음", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string emailPattern = @"^\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b";
            Regex regex = new Regex(emailPattern);

            return regex.IsMatch(input);
        }

        public static bool IsReasonableHeight(this string input, out string msg)
        {
            string pattern = @"^\d{3}$";
            Regex regex = new Regex(pattern);

            if (regex.IsMatch(input))
            {
                int height = int.Parse(input);

                // 假设华人18岁以上的普遍身高范围在 150 到 190 之间
                if (height >= 140 && height <= 200)
                {
                    msg = "";
                    return true;
                }
                else
                {
                    msg = "不符合常理, 请检查";
                    return false;
                }
            }
            else
            {
                msg = "格式不正确";
                return false;
            }
        }

        public static bool IsBloodTypeValid(this string input) => IsMatchRegex(input, @"^(A|B|O|AB)?$");

        public static bool IsMaritalStatusValid(this string input) => IsMatchRegex(input, @"^(기혼|미혼)$");

        public static bool IsValidPointDate(this string input, out DateTime date)
        {
            // 使用正则表达式验证是否是日期
            string pattern = @"^(\d{4})\.(\d{1,2})\.(\d{1,2})$";
            Regex regex = new Regex(pattern);

            Match match = regex.Match(input);

            if (match.Success)
            {
                // 提取匹配的年、月、日
                int year = int.Parse(match.Groups[1].Value);
                int month = int.Parse(match.Groups[2].Value);
                int day = int.Parse(match.Groups[3].Value);
                
                if (match.Groups[2].Value.StartsWith("0") && match.Groups[2].Value.Length > 1 ||  match.Groups[3].Value.StartsWith("0") && match.Groups[3].Value.Length > 1)
                {
                    date = new DateTime(year, month, day);
                    return false;
                }

                // 尝试构建DateTime对象
                try
                {
                    date = new DateTime(year, month, day);
                    return true;
                }
                catch (Exception)
                {
                    // 构建DateTime对象失败
                    date = DateTime.MinValue;
                    return false;
                }
            }

            // 不符合日期格式
            date = DateTime.MinValue;
            return false;
        }

        public static bool IsValidyyyyMMddDate(this string input, out DateTime date)
        {
            // 使用正则表达式验证是否是日期
            string pattern = @"^(\d{4})-(\d{2})-(\d{2})$";
            Regex regex = new Regex(pattern);

            Match match = regex.Match(input);

            if (match.Success)
            {
                // 提取匹配的年、月、日
                int year = int.Parse(match.Groups[1].Value);
                int month = int.Parse(match.Groups[2].Value);
                int day = int.Parse(match.Groups[3].Value);

                // 尝试构建DateTime对象
                try
                {
                    date = new DateTime(year, month, day);
                    return true;
                }
                catch (Exception)
                {
                    // 构建DateTime对象失败
                    date = DateTime.MinValue;
                    return false;
                }
            }

            // 不符合日期格式
            date = DateTime.MinValue;
            return false;
        }

        public static string ReplaceFirst(this string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
    }
}