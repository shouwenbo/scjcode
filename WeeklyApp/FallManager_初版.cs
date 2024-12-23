using Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace WeeklyApp
{
    public class FallManager_初版
    {
        public void Output(List<string[]> list)
        {
            string inputFilePath = @"template\韩文掉落模板.docx";
            string outputDir = $"output/生成时间 {DateTime.Now:yyyy年MM月dd日HH时mm分ss秒}/";

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var sunday = DateTime.Now.GetSundayOfCurrentWeek();
            var year = sunday.Year;
            var month = sunday.Month;
            var day = sunday.Day;
            var scjyear = year - 1983;

            foreach (var item in list)
            {
                var 姓名 = item[1].Trim();
                var 韩文姓名 = item[2].Trim();
                var 期数 = item[5].Trim();
                var 探访者 = item[16].Trim();
                var 探访次数 = item[17].Trim();
                var 讲师意见 = item[18].Trim();
                var 传道师意见 = item[19].Trim();
                var 掉落课数 = item[23].Trim();
                var 期数全称 = item[20].Trim();
                var 掉落日期 = item[21].Trim();
                var 电话 = item[22].Trim();
                var 中级讲师 = item[24].Trim();
                var 高级讲师 = item[25].Trim();
                var 传道师 = item[26].Trim();
                var 掉落事由 = item[27].Trim();


                string outputFileName = $@"{outputDir}부산야고보-탈락신청서중국무한교육관비){期数}({韩文姓名})-신{scjyear}({year}).{month}.{day}..docx";

                using (DocX document = DocX.Load(inputFilePath))
                {
                    document.ReplaceText("{{期数}}", 期数全称);
                    document.ReplaceText("{{姓名}}", 姓名);
                    document.ReplaceText("{{电话}}", 电话);
                    document.ReplaceText("{{掉落课数}}", 掉落课数);
                    document.ReplaceText("{{掉落日期}}", 掉落日期);
                    document.ReplaceText("{{商谈者}}", 探访者);
                    document.ReplaceText("{{3次}}", 探访次数);
                    document.ReplaceText("{{掉落事由}}", 掉落事由);
                    document.ReplaceText("{{传道师}}", 传道师);
                    document.ReplaceText("{{传道师事由}}", 传道师意见);
                    document.ReplaceText("{{讲师}}", ((!string.IsNullOrEmpty(高级讲师) && !高级讲师.Contains("#") && !高级讲师.Contains("0")) ? 高级讲师 : 中级讲师));
                    document.ReplaceText("{{讲师事由}}", 讲师意见);
                    document.ReplaceText("{{部长}}", "费增明/159-7200-5442");
                    document.ReplaceText("{{新天纪年}}", scjyear.ToString());
                    document.ReplaceText("{{新天纪月}}", month.ToString());
                    document.ReplaceText("{{新天纪日}}", day.ToString());

                    // Save the modified document as a new file
                    document.SaveAs(outputFileName);
                }
            }

            MessageBox.Show("导出完成");
        }

        public void ParseData(string text)
        {
            string[] lines = text.Split('\n');

            var list = new List<string[]>();

            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] cells = line.Split('\t');

                if (cells.Length != 28)
                {
                    MessageBox.Show("格式有误，请重新复制！");
                }

                list.Add(cells);
            }

            Output(list);
        }
    }
}
