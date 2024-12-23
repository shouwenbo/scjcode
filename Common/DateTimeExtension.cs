namespace Common
{
    public static class DateTimeExtension
    {
        /// <summary>
        /// 获取本周日
        /// </summary>
        public static DateTime GetSundayOfCurrentWeek(this DateTime date)
        {
            int daysUntilSunday = ((int)DayOfWeek.Sunday - (int)date.DayOfWeek + 7) % 7;
            return date.AddDays(daysUntilSunday);
        }

        /// <summary>
        /// 获取上周日
        /// </summary>
        public static DateTime GetSundayOfPreviousWeek(this DateTime date)
        {
            DateTime currentSunday = date.GetSundayOfCurrentWeek();
            return currentSunday.AddDays(-7);
        }
    }
}
