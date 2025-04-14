using Microsoft.Extensions.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace JJBWinformApp
{
    public class TableService
    {
        private readonly IDialogService dialogService;
        private readonly RowIndexOptions rowIndexOptions;
        public TableService(IDialogService dialogService, IOptions<RowIndexOptions> rowIndexOptions)
        {
            this.dialogService = dialogService;
            this.rowIndexOptions = rowIndexOptions.Value;
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
        /// <param name="startColumn">开始列</param>
        /// <param name="relax">是否宽松处理 默认为true 若为false将读取完整的表格内容</param>
        /// <param name="backColor">是否必须包含背景色</param>
        /// <param name="requirePhone">是否必须包含手机号</param>
        /// <returns>学员列表</returns>
        public List<Member>? Load(string path, int startColumn, bool relax = true, bool backColor = false, bool requirePhone = false)
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
                    var list = new List<Member>();
                    if (startColumn == 0)
                    {
                        startColumn = rowIndexOptions.StartColumn;
                    }
                    for (int column = startColumn - 1; column < lastColumnNum; column++)
                    {
                        var indexBackColor = GetCellBackgroundColor(sheet, column, "排序");
                        if (backColor && indexBackColor == null)
                        {
                            continue;
                        }
                        if (requirePhone && string.IsNullOrEmpty(GetColumnString(sheet, column, "电话")))
                        {
                            continue;
                        }

                        var member = new Member()
                        {
                            ColumnIndex = column + 1,
                            ColumnName = ExcelColumnIndexToName(column),
                            Index = GetColumnString(sheet, column, "排序"),
                            KoreanName = GetColumnString(sheet, column, "韩语姓名"),
                            ChineseName = GetColumnString(sheet, column, "汉语姓名"),
                            Country = GetColumnString(sheet, column, "国籍"),
                            IDCardBirthday = GetColumnString(sheet, column, "身份证生日"),
                            Gender = GetColumnString(sheet, column, "性别"),
                            ActualBirthday = GetColumnString(sheet, column, "实际生日"),
                            IDCardAddressChinese = GetColumnString(sheet, column, "身份证地址中文"),
                            IDCardAddressKorean = GetColumnString(sheet, column, "身份证地址韩文"),
                            CurrentAddressChinese = GetColumnString(sheet, column, "现住址"),
                            CurrentAddressKorean = GetColumnString(sheet, column, "现住址翻译"),
                            Phone = GetColumnString(sheet, column, "电话"),
                            Email = GetColumnString(sheet, column, "邮箱"),
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
                        member.SiblingRelationship = member.SiblingRelationship.Trim();

                        #endregion

                        for (int i = 1; i <= 3; i++)
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
                        for (int i = 1; i <= 6; i++)
                        {
                            var familyInfo = new FamilyMemberInfo()
                            {
                                Index = i,
                                Name = GetColumnString(sheet, column, $"家族事项{i}姓名"),
                                Gender = GetColumnString(sheet, column, $"家族事项{i}性别"),
                                Birthday = GetColumnString(sheet, column, $"家族事项{i}生日"),
                                Relationship = GetColumnString(sheet, column, $"家族事项{i}关系"),
                                Religion = GetColumnString(sheet, column, $"家族事项{i}教团"),
                                OtherReligion = GetColumnString(sheet, column, $"家族事项{i}教团-其他"),
                                BirthChurch = GetColumnString(sheet, column, $"家族事项{i}出身教会"),
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

        public List<Member>? LoadGatherList(string path, int startColumn = 4)
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
                    var list = new List<Member>();
                    if (startColumn == 0)
                    {
                        startColumn = rowIndexOptions.StartColumn;
                    }
                    for (int column = startColumn - 1; column < lastColumnNum; column++)
                    {
                        var member = new Member()
                        {
                            ColumnIndex = column + 1,
                            ColumnName = ExcelColumnIndexToName(column),
                            Index = GetColumnString(sheet, column, "排序"),
                            KoreanName = GetColumnString(sheet, column, "韩语姓名"),
                            ChineseName = GetColumnString(sheet, column, "汉语姓名"),
                            Country = GetColumnString(sheet, column, "国籍"),
                            IDCardBirthday = GetColumnString(sheet, column, "身份证生日"),
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
