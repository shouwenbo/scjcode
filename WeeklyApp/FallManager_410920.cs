using Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace WeeklyApp
{
    public class FallManager_410920
    {
        public void Output(List<string[]> list)
        {
            string inputFilePath = @"template\韩文掉落模板.docx";
            string outputDir = $"output/生成时间 {DateTime.Now:yyyyMMdd_HHmmss}/";

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
                var 掉落阶段 = item[9].Trim().Replace("중등", "中级").Replace("고등", "高级").Replace("새신자", "新家族");
                var 掉落课数 = item[10].Trim().Replace("과", "课").Replace("회", "回");

                string outputFileName = $@"{outputDir}부산야고보-탈락신청서중국무한교육관비){期数}({韩文姓名})-신{scjyear}({year}).{month}.{day}..docx";

                using (DocX document = DocX.Load(inputFilePath))
                {
                    document.ReplaceText("{{期数}}", 期数 + "期");
                    document.ReplaceText("{{姓名}}", 姓名);
                    document.ReplaceText("{{掉落课数}}", 掉落阶段 + 掉落课数);
                    document.ReplaceText("{{新天纪年}}", scjyear.ToString());
                    document.ReplaceText("{{新天纪月}}", month.ToString().PadLeft(2, '0'));
                    document.ReplaceText("{{新天纪日}}", day.ToString().PadLeft(2, '0'));

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

                if (cells.Length != 15)
                {
                    MessageBox.Show("格式有误，请重新复制！");
                }

                list.Add(cells);
            }

            Output(list);
        }
    }
}
