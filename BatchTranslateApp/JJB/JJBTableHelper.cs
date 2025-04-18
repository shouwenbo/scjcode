using Common;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace BatchTranslateApp
{
    public class JJBTableHelper
    {
        private readonly IDialogService dialogService;
        private RowIndexOptions rowIndexOptions;
        public JJBTableHelper(IDialogService dialogService)
        {
            this.dialogService = dialogService;
            this.rowIndexOptions = new RowIndexOptions()
            {
                List = new List<RowIndexItem>()
                {
                    new RowIndexItem(){ Index = 1, Name = "排序" },
                    new RowIndexItem(){ Index = 2, Name = "韩语姓名" },
                    new RowIndexItem(){ Index = 3, Name = "工号" },
                    new RowIndexItem(){ Index = 4, Name = "汉语姓名" },
                    new RowIndexItem(){ Index = 5, Name = "国籍" },
                    new RowIndexItem(){ Index = 6, Name = "法定出生生日" },
                    new RowIndexItem(){ Index = 7, Name = "性别" },
                    new RowIndexItem(){ Index = 8, Name = "现住址" },
                    new RowIndexItem(){ Index = 9, Name = "身份证地址" },
                    new RowIndexItem(){ Index = 10, Name = "电话" },
                    new RowIndexItem(){ Index = 11, Name = "身高" },
                    new RowIndexItem(){ Index = 12, Name = "血型" },
                    new RowIndexItem(){ Index = 13, Name = "婚否" },
                    new RowIndexItem(){ Index = 14, Name = "CD类型" },
                    new RowIndexItem(){ Index = 15, Name = "出身教团" },
                    new RowIndexItem(){ Index = 16, Name = "其他宗教" },
                    new RowIndexItem(){ Index = 17, Name = "教会名" },
                    new RowIndexItem(){ Index = 18, Name = "职分" },
                    new RowIndexItem(){ Index = 19, Name = "信仰时间" },
                    new RowIndexItem(){ Index = 20, Name = "母胎信仰" },
                    new RowIndexItem(){ Index = 21, Name = "邮箱" },
                    new RowIndexItem(){ Index = 23, Name = "最终学历" },

                    // new RowIndexItem(){ Index = 24, Name = "学历信息1分类" },
                    // new RowIndexItem(){ Index = 25, Name = "学历信息1学校名" },
                    // new RowIndexItem(){ Index = 26, Name = "学历信息1专业" },
                    // new RowIndexItem(){ Index = 28, Name = "学历信息1状态" },
                    // new RowIndexItem(){ Index = 29, Name = "学历信息1期间" },
                    
                    new RowIndexItem(){ Index = 29, Name = "职业分类" },
                    new RowIndexItem(){ Index = 30, Name = "其他职业" },
                    new RowIndexItem(){ Index = 31, Name = "工作单位名称" },
                    new RowIndexItem(){ Index = 32, Name = "职位" },
                    new RowIndexItem(){ Index = 38, Name = "兄弟姐妹关系" },
                    new RowIndexItem(){ Index = 40, Name = "家族事项1姓名" },
                    new RowIndexItem(){ Index = 41, Name = "家族事项1性别" },
                    new RowIndexItem(){ Index = 42, Name = "家族事项1生日" },
                    new RowIndexItem(){ Index = 43, Name = "家族事项1关系" },
                    new RowIndexItem(){ Index = 44, Name = "家族事项1出席教会" },
                    new RowIndexItem(){ Index = 45, Name = "家族事项1SCJ工号" },
                    new RowIndexItem(){ Index = 46, Name = "家族事项2姓名" },
                    new RowIndexItem(){ Index = 47, Name = "家族事项2性别" },
                    new RowIndexItem(){ Index = 48, Name = "家族事项2生日" },
                    new RowIndexItem(){ Index = 49, Name = "家族事项2关系" },
                    new RowIndexItem(){ Index = 50, Name = "家族事项2出席教会" },
                    new RowIndexItem(){ Index = 51, Name = "家族事项2SCJ工号" },
                },
                StartColumn = 4
            };
        }

        private int GetPropertyRowIndex(string propertyName)
        {
            var item = rowIndexOptions.List.FirstOrDefault(p => p.Name == propertyName);
            if (item == null)
            {
                var msg = $"配置文件中没有‘{propertyName}’的配置";
                throw new Exception(msg);
            }
            return item.Index;
        }

        private string ExcelColumnIndexToName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = "";

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private IColor? GetCellBackgroundColor(ISheet sheet, int column, string propertyName)
        {
            var row = sheet.GetRow(GetPropertyRowIndex(propertyName) - 1);
            if (row == null)
            {
                var msg = $"第{(column + 1)}列找不到‘{propertyName}’这一行";
                throw new Exception(msg);
            }
            var cell = row.GetCell(column);
            if (cell == null)
            {
                var msg = $"第{(column + 1)}列找不到‘{propertyName}’这一格";
                throw new Exception(msg);
            }
            return cell.CellStyle.FillBackgroundColorColor;
        }

        private string GetCellString(ICell cell)
        {
            string cellValue = "";

            if (cell.CellType == CellType.Numeric)
            {
                // 如果是数字单元格，可以将其转换为文本
                cellValue = cell.NumericCellValue.ToString();
            }
            else if (cell.CellType == CellType.String)
            {
                // 如果是文本单元格，直接获取文本值
                cellValue = cell.StringCellValue;
            }

            return cellValue;
        }

        private string GetColumnString(ISheet sheet, int column, string propertyName)
        {
            // 由于学历会在表格中间插入多行，所以学历以下的行需要加上dynamicRowCount
            var row = sheet.GetRow(GetPropertyRowIndex(propertyName) - 1);
            if (row == null)
            {
                var msg = $"第{(column + 1)}列找不到‘{propertyName}’这一行";
                throw new Exception(msg);
            }
            var cell = row.GetCell(column);
            if (cell == null)
            {
                return ""; // TODO 问题还需要继续排查
                //var msg = $"第{(column + 1)}列找不到‘{propertyName}’这一格";
                //throw new Exception(msg);
            }
            return GetCellString(cell);
        }

        private IComment GetColumnComment(ISheet sheet, int column, string propertyName)
        {
            var row = sheet.GetRow(GetPropertyRowIndex(propertyName) - 1);
            if (row == null)
            {
                var msg = $"第{(column + 1)}列找不到‘{propertyName}’这一行";
                throw new Exception(msg);
            }
            var cell = row.GetCell(column);
            if (cell == null)
            {
                var msg = $"第{(column + 1)}列找不到‘{propertyName}’这一格";
                throw new Exception(msg);
            }
            return cell.CellComment;
        }

        private string? GetColumnCommentString(ISheet sheet, int column, string propertyName)
        {
            return GetColumnComment(sheet, column, propertyName)?.String?.String;
        }

        /// <summary>
        /// 读取表格
        /// </summary>
        /// <param name="path">表格地址</param>
        /// <returns>学员列表</returns>
        public List<JJBMember>? Load(string path)
        {
            #region 尝试打开文件
            IWorkbook workbook = path.GetNpoiWorkbook(out string openErr, "12000", "144000", "wh12000", "0314");
            if (workbook == null)
            {
                dialogService.Alert(openErr);
                return null;
            }
            #endregion

            try
            {
                #region 读取工作簿、首行、总列
                ISheet sheet = workbook.GetSheetAt(0);
                var firstRowNum = sheet.FirstRowNum;
                if (firstRowNum < 0)
                {
                    dialogService.Alert("找不到第一行, 信息读取失败");
                    return null;
                }
                var firstRow = sheet.GetRow(firstRowNum); // 获取首行
                var fourRow = sheet.GetRow(3); // 获取第四行
                var lastColumnNum = firstRow.LastCellNum; // 获取列的总数按
                #endregion

                var list = new List<JJBMember>(); // 定义学生列表

                #region 计算开始列 startColumn

                var startColumn = 0;
                var findFirstColumn = false;
                for (int column = 0; column < lastColumnNum; column++)
                {
                    var firstRowString = GetCellString(firstRow.GetCell(column));
                    var fourRowString = GetCellString(fourRow.GetCell(column));
                    if (new Regex(@"^(?!基本信息$)[\u4e00-\u9fff]{2,4}$").IsMatch(fourRowString) && firstRowString != "예시" && firstRowString != "示例")
                    {
                        findFirstColumn = true;
                        startColumn = column + 1;
                        break;
                    }

                    // if (int.TryParse(GetCellString(fourRow.GetCell(column)), out int firstColumn))
                    // {
                    //     findFirstColumn = true;
                    //     startColumn = column + 1;
                    //     break;
                    // }
                }

                if (!findFirstColumn)
                {
                    dialogService.Alert("无法找到首列，请检查是否设置了序号");
                    return null;
                }

                #endregion

                #region 验证行，获取新增动态行数量

                var lastScjNumberRowIndex = 0;
                var dynamicRowCount = 0;
                var educationNumber = 1; // 几个学历（动态、至少一个）

                for (int i =  0; i <= sheet.LastRowNum; i++)
                {
                    var cellStr = GetCellString(sheet.GetRow(i).GetCell(1));
                    if (cellStr == "工号" ||  cellStr == "신천지 고유번호")
                    {
                        lastScjNumberRowIndex = i;
                    }
                }

                if (lastScjNumberRowIndex == 50) // 1个学历
                {
                    dynamicRowCount = 0; // 动态增加0行
                    educationNumber = 1;
                }
                else if (lastScjNumberRowIndex == 55) // 2个学历
                {
                    dynamicRowCount = 5; // 动态增加5行
                    educationNumber = 2;
                }
                else if (lastScjNumberRowIndex == 60) // 3个学历
                {
                    dynamicRowCount = 10; // 动态增加10行
                    educationNumber = 3;
                }
                else if (lastScjNumberRowIndex == 65) // 4个学历
                {
                    dynamicRowCount = 15; // 动态增加15行
                    educationNumber = 4;
                }
                else
                {
                    dialogService.Alert("表格的行数不正确，请检查学历部分是否按要求新增了");
                    return null;
                }

                for (int i = 0; i < educationNumber; i++)
                {
                    this.rowIndexOptions.List.Add(new RowIndexItem() { Index = 24 + (i * 5), Name = $"学历信息{i + 1}分类" });
                    this.rowIndexOptions.List.Add(new RowIndexItem() { Index = 25 + (i * 5), Name = $"学历信息{i + 1}学校名" });
                    this.rowIndexOptions.List.Add(new RowIndexItem() { Index = 26 + (i * 5), Name = $"学历信息{i + 1}专业" });
                    this.rowIndexOptions.List.Add(new RowIndexItem() { Index = 27 + (i * 5), Name = $"学历信息{i + 1}状态" });
                    this.rowIndexOptions.List.Add(new RowIndexItem() { Index = 28 + (i * 5), Name = $"学历信息{i + 1}期间" });
                }
                    
                this.rowIndexOptions.List.First(p => p.Name == "职业分类").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "其他职业").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "工作单位名称").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "职位").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "兄弟姐妹关系").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项1姓名").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项1性别").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项1生日").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项1关系").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项1出席教会").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项1SCJ工号").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项2姓名").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项2性别").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项2生日").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项2关系").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项2出席教会").Index += dynamicRowCount;
                this.rowIndexOptions.List.First(p => p.Name == "家族事项2SCJ工号").Index += dynamicRowCount;

                #endregion


                for (int column = startColumn - 1; column < lastColumnNum; column++)
                {
                    var member = new JJBMember()
                    {
                        ColumnIndex = column + 1,
                        ColumnName = ExcelColumnIndexToName(column),
                        Index = GetColumnString(sheet, column, "排序"),
                        KoreanName = GetColumnString(sheet, column, "韩语姓名"),
                        ScjNumber = GetColumnString(sheet, column, "工号"),
                        ChineseName = GetColumnString(sheet, column, "汉语姓名"),
                        Country = GetColumnString(sheet, column, "国籍"),
                        IDCardBirthday = GetColumnString(sheet, column, "法定出生生日"),
                        Gender = GetColumnString(sheet, column, "性别"),
                        // ActualBirthday = GetColumnString(sheet, column, "实际生日"),
                        // IDCardAddressChinese = GetColumnString(sheet, column, "身份证地址中文"),
                        // IDCardAddressKorean = GetColumnString(sheet, column, "身份证地址韩文"),
                        CurrentAddressChinese = GetColumnString(sheet, column, "现住址"),
                        IDCardAddressChinese = GetColumnString(sheet, column, "身份证地址"),
                        // CurrentAddressKorean = GetColumnString(sheet, column, "现住址翻译"),
                        Phone = GetColumnString(sheet, column, "电话"),
                        Height = GetColumnString(sheet, column, "身高"),
                        BloodType = GetColumnString(sheet, column, "血型"),
                        IsMarried = GetColumnString(sheet, column, "婚否"),
                        CDType = GetColumnString(sheet, column, "CD类型"),
                        CDTypeComment = GetColumnComment(sheet, column, "CD类型"),
                        CDTypeCommentString = GetColumnCommentString(sheet, column, "CD类型"),
                        BirthReligion = GetColumnString(sheet, column, "出身教团"),
                        OtherReligion = GetColumnString(sheet, column, "其他宗教"),
                        ChurchName = GetColumnString(sheet, column, "教会名"),
                        Position = GetColumnString(sheet, column, "职分"),
                        YearsOfFaith = GetColumnString(sheet, column, "信仰时间"),
                        MaternalFaith = GetColumnString(sheet, column, "母胎信仰"),
                        Email = GetColumnString(sheet, column, "邮箱"),
                        FinalEducation = GetColumnString(sheet, column, "最终学历"),
                        JobType = GetColumnString(sheet, column, "职业分类"),
                        JobOther = GetColumnString(sheet, column, "其他职业"),
                        JobCompany = GetColumnString(sheet, column, "工作单位名称"),
                        JobName = GetColumnString(sheet, column, "职位"),
                        SiblingRelationship = GetColumnString(sheet, column, "兄弟姐妹关系"),
                        EducationList = new List<EducationInfo>(),
                        FamilyMemberList = new List<FamilyMemberInfo>()
                    };

                    #region 不需要那么严格

                    member.CDType = member.CDType.TrimEnd();
                    member.CDTypeCommentString = (member.CDTypeCommentString ?? "").TrimEnd();
                    member.FinalEducation = member.FinalEducation.TrimEnd();
                    // member.SiblingRelationship = member.SiblingRelationship.Trim();

                    #endregion

                    for (int i = 1; i <= educationNumber; i++)
                    {
                        var educationInfo = new EducationInfo()
                        {
                            Index = i,
                            Category = GetColumnString(sheet, column, $"学历信息{i}分类"),
                            SchoolName = GetColumnString(sheet, column, $"学历信息{i}学校名"),
                            Major = GetColumnString(sheet, column, $"学历信息{i}专业"),
                            Status = GetColumnString(sheet, column, $"学历信息{i}状态"),
                            Period = GetColumnString(sheet, column, $"学历信息{i}期间")
                        };
                        if (!string.IsNullOrEmpty(educationInfo.Category) || !string.IsNullOrEmpty(educationInfo.SchoolName) || !string.IsNullOrEmpty(educationInfo.Major) || !string.IsNullOrEmpty(educationInfo.Status) || !string.IsNullOrEmpty(educationInfo.Period))
                        {
                            member.EducationList.Add(educationInfo);
                        }
                    }
                    for (int i = 1; i <= 2; i++)
                    {
                        var familyInfo = new FamilyMemberInfo()
                        {
                            Index = i,
                            Name = GetColumnString(sheet, column, $"家族事项{i}姓名"),
                            Gender = GetColumnString(sheet, column, $"家族事项{i}性别"),
                            Birthday = GetColumnString(sheet, column, $"家族事项{i}生日"),
                            Relationship = GetColumnString(sheet, column, $"家族事项{i}关系"),
                            // Religion = GetColumnString(sheet, column, $"家族事项{i}教团"),
                            // OtherReligion = GetColumnString(sheet, column, $"家族事项{i}教团-其他"),
                            // BirthChurch = GetColumnString(sheet, column, $"家族事项{i}出身教会"),
                            BirthChurch = GetColumnString(sheet, column, $"家族事项{i}出席教会"),
                            SCJNumber = GetColumnString(sheet, column, $"家族事项{i}SCJ工号")
                        };

                        #region 不需要那么严格

                        familyInfo.Name = familyInfo.Name.Trim();
                        familyInfo.Relationship = familyInfo.Relationship.TrimEnd();

                        #endregion

                        if (!string.IsNullOrEmpty(familyInfo.Name) || !string.IsNullOrEmpty(familyInfo.Gender) || !string.IsNullOrEmpty(familyInfo.Birthday) || !string.IsNullOrEmpty(familyInfo.Relationship) || !string.IsNullOrEmpty(familyInfo.Religion) || !string.IsNullOrEmpty(familyInfo.OtherReligion) || !string.IsNullOrEmpty(familyInfo.BirthChurch) || !string.IsNullOrEmpty(familyInfo.SCJNumber))
                        {
                            member.FamilyMemberList.Add(familyInfo);
                        }
                    }

                    // 序号 韩文名 中文名
                    if (!string.IsNullOrEmpty(member.ChineseName))
                    {
                        list.Add(member);
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                dialogService.Alert(ex.Message);
                return null;
            }
            finally
            {
                workbook.Close();
            }
        }

        public List<JJBMember>? LoadGatherList(string path, int startColumn = 4)
        {
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                try
                {
                    ISheet sheet = workbook.GetSheetAt(0);
                    var firstRowNum = sheet.FirstRowNum;
                    if (firstRowNum < 0)
                    {
                        dialogService.Alert("找不到第一行, 信息读取失败");
                        return null;
                    }
                    var lastColumnNum = sheet.GetRow(firstRowNum).LastCellNum;
                    var list = new List<JJBMember>();
                    if (startColumn == 0)
                    {
                        startColumn = rowIndexOptions.StartColumn;
                    }
                    for (int column = startColumn - 1; column < lastColumnNum; column++)
                    {
                        var member = new JJBMember()
                        {
                            ColumnIndex = column + 1,
                            ColumnName = ExcelColumnIndexToName(column),
                            Index = GetColumnString(sheet, column, "排序"),
                            KoreanName = GetColumnString(sheet, column, "韩语姓名"),
                            ChineseName = GetColumnString(sheet, column, "汉语姓名"),
                            Country = GetColumnString(sheet, column, "国籍"),
                            IDCardBirthday = GetColumnString(sheet, column, "法定出生生日"),
                            Gender = GetColumnString(sheet, column, "性别"),
                            ActualBirthday = GetColumnString(sheet, column, "实际生日"),
                        };
                        if (!string.IsNullOrEmpty(member.Index) && !string.IsNullOrEmpty(member.KoreanName) && !string.IsNullOrEmpty(member.ChineseName))
                        {
                            list.Add(member);
                        }
                    }

                    return list;
                }
                catch (Exception ex)
                {
                    dialogService.Alert(ex.Message);
                    return null;
                }
                finally
                {
                    workbook.Close();
                    fileStream.Close();
                }
            }
        }
    }
}
