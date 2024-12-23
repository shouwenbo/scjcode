using Common;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace JJBSelfCheckApp
{
    public class TableHelper
    {
        private readonly IDialogService dialogService;
        private RowIndexOptions rowIndexOptions;
        public TableHelper(IDialogService dialogService)
        {
            this.dialogService = dialogService;
            this.rowIndexOptions = new RowIndexOptions()
            {
                List = new List<RowIndexItem>()
                {
                    new RowIndexItem(){ Index = 1, Name = "排序" },
                    new RowIndexItem(){ Index = 2, Name = "韩语姓名" },
                    new RowIndexItem(){ Index = 4, Name = "汉语姓名" },
                    new RowIndexItem(){ Index = 5, Name = "国籍" },
                    new RowIndexItem(){ Index = 6, Name = "身份证生日" },
                    new RowIndexItem(){ Index = 7, Name = "性别" },
                    new RowIndexItem(){ Index = 8, Name = "实际生日" },
                    new RowIndexItem(){ Index = 9, Name = "身份证地址中文" },
                    new RowIndexItem(){ Index = 10, Name = "身份证地址韩文" },
                    new RowIndexItem(){ Index = 11, Name = "现住址" },
                    new RowIndexItem(){ Index = 12, Name = "现住址翻译" },
                    new RowIndexItem(){ Index = 13, Name = "电话" },
                    new RowIndexItem(){ Index = 14, Name = "邮箱" },
                    new RowIndexItem(){ Index = 15, Name = "身高" },
                    new RowIndexItem(){ Index = 16, Name = "血型" },
                    new RowIndexItem(){ Index = 17, Name = "婚否" },
                    new RowIndexItem(){ Index = 18, Name = "CD类型" },
                    new RowIndexItem(){ Index = 19, Name = "出身教团" },
                    new RowIndexItem(){ Index = 20, Name = "其他宗教" },
                    new RowIndexItem(){ Index = 21, Name = "教会名" },
                    new RowIndexItem(){ Index = 22, Name = "职分" },
                    new RowIndexItem(){ Index = 23, Name = "信仰时间" },
                    new RowIndexItem(){ Index = 24, Name = "母胎信仰" },
                    new RowIndexItem(){ Index = 30, Name = "最终学历" },
                    new RowIndexItem(){ Index = 31, Name = "学历信息1分类" },
                    new RowIndexItem(){ Index = 32, Name = "学历信息1学校名" },
                    new RowIndexItem(){ Index = 33, Name = "学历信息1专业" },
                    new RowIndexItem(){ Index = 34, Name = "学历信息1状态" },
                    new RowIndexItem(){ Index = 35, Name = "学历信息1期间" },
                    new RowIndexItem(){ Index = 36, Name = "学历信息2分类" },
                    new RowIndexItem(){ Index = 37, Name = "学历信息2学校名" },
                    new RowIndexItem(){ Index = 38, Name = "学历信息2专业" },
                    new RowIndexItem(){ Index = 39, Name = "学历信息2状态" },
                    new RowIndexItem(){ Index = 40, Name = "学历信息2期间" },
                    new RowIndexItem(){ Index = 41, Name = "学历信息3分类" },
                    new RowIndexItem(){ Index = 42, Name = "学历信息3学校名" },
                    new RowIndexItem(){ Index = 43, Name = "学历信息3专业" },
                    new RowIndexItem(){ Index = 44, Name = "学历信息3状态" },
                    new RowIndexItem(){ Index = 45, Name = "学历信息3期间" },
                    new RowIndexItem(){ Index = 86, Name = "兄弟姐妹关系" },
                    new RowIndexItem(){ Index = 87, Name = "家族事项1姓名" },
                    new RowIndexItem(){ Index = 88, Name = "家族事项1性别" },
                    new RowIndexItem(){ Index = 89, Name = "家族事项1生日" },
                    new RowIndexItem(){ Index = 90, Name = "家族事项1关系" },
                    new RowIndexItem(){ Index = 91, Name = "家族事项1教团" },
                    new RowIndexItem(){ Index = 92, Name = "家族事项1教团-其他" },
                    new RowIndexItem(){ Index = 93, Name = "家族事项1出身教会" },
                    new RowIndexItem(){ Index = 94, Name = "家族事项1SCJ工号" },
                    new RowIndexItem(){ Index = 95, Name = "家族事项2姓名" },
                    new RowIndexItem(){ Index = 96, Name = "家族事项2性别" },
                    new RowIndexItem(){ Index = 97, Name = "家族事项2生日" },
                    new RowIndexItem(){ Index = 98, Name = "家族事项2关系" },
                    new RowIndexItem(){ Index = 99, Name = "家族事项2教团" },
                    new RowIndexItem(){ Index = 100, Name = "家族事项2教团-其他" },
                    new RowIndexItem(){ Index = 101, Name = "家族事项2出身教会" },
                    new RowIndexItem(){ Index = 102, Name = "家族事项2SCJ工号" },
                    new RowIndexItem(){ Index = 103, Name = "家族事项3姓名" },
                    new RowIndexItem(){ Index = 104, Name = "家族事项3性别" },
                    new RowIndexItem(){ Index = 105, Name = "家族事项3生日" },
                    new RowIndexItem(){ Index = 106, Name = "家族事项3关系" },
                    new RowIndexItem(){ Index = 107, Name = "家族事项3教团" },
                    new RowIndexItem(){ Index = 108, Name = "家族事项3教团-其他" },
                    new RowIndexItem(){ Index = 109, Name = "家族事项3出身教会" },
                    new RowIndexItem(){ Index = 110, Name = "家族事项3SCJ工号" },
                    new RowIndexItem(){ Index = 111, Name = "家族事项4姓名" },
                    new RowIndexItem(){ Index = 112, Name = "家族事项4性别" },
                    new RowIndexItem(){ Index = 113, Name = "家族事项4生日" },
                    new RowIndexItem(){ Index = 114, Name = "家族事项4关系" },
                    new RowIndexItem(){ Index = 115, Name = "家族事项4教团" },
                    new RowIndexItem(){ Index = 116, Name = "家族事项4教团-其他" },
                    new RowIndexItem(){ Index = 117, Name = "家族事项4出身教会" },
                    new RowIndexItem(){ Index = 118, Name = "家族事项4SCJ工号" },
                    new RowIndexItem(){ Index = 119, Name = "家族事项5姓名" },
                    new RowIndexItem(){ Index = 120, Name = "家族事项5性别" },
                    new RowIndexItem(){ Index = 121, Name = "家族事项5生日" },
                    new RowIndexItem(){ Index = 122, Name = "家族事项5关系" },
                    new RowIndexItem(){ Index = 123, Name = "家族事项5教团" },
                    new RowIndexItem(){ Index = 124, Name = "家族事项5教团-其他" },
                    new RowIndexItem(){ Index = 125, Name = "家族事项5出身教会" },
                    new RowIndexItem(){ Index = 126, Name = "家族事项5SCJ工号" },
                    new RowIndexItem(){ Index = 127, Name = "家族事项6姓名" },
                    new RowIndexItem(){ Index = 128, Name = "家族事项6性别" },
                    new RowIndexItem(){ Index = 129, Name = "家族事项6生日" },
                    new RowIndexItem(){ Index = 130, Name = "家族事项6关系" },
                    new RowIndexItem(){ Index = 131, Name = "家族事项6教团" },
                    new RowIndexItem(){ Index = 132, Name = "家族事项6教团-其他" },
                    new RowIndexItem(){ Index = 133, Name = "家族事项6出身教会" },
                    new RowIndexItem(){ Index = 134, Name = "家族事项6SCJ工号" }
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
                IWorkbook workbook = null;
                try
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                catch (ZipException)
                {
                    try
                    {
                        fileStream.Position = 0;
                        workbook = WorkbookFactory.Create(fileStream, "12000", true);
                    }
                    catch (Exception)
                    {
                        try
                        {
                            fileStream.Position = 0;
                            workbook = WorkbookFactory.Create(fileStream, "144000", true);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                fileStream.Position = 0;
                                workbook = WorkbookFactory.Create(fileStream, "wh12000", true);
                            }
                            catch (Exception)
                            {
                                fileStream.Position = 0;
                                workbook = WorkbookFactory.Create(fileStream, "0314", true);
                            }
                        }
                    }
                }
                 
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

                        // 以序号 韩文名 中文名 为依据
                        if (!string.IsNullOrEmpty(member.Index) && !string.IsNullOrEmpty(member.KoreanName) && !string.IsNullOrEmpty(member.ChineseName))
                        {
                            if (!string.IsNullOrEmpty(member.Phone) || !string.IsNullOrEmpty(member.CurrentAddressChinese) || !string.IsNullOrEmpty(member.IDCardAddressChinese) || !string.IsNullOrEmpty(member.BirthReligion))
                            {
                                list.Add(member);
                            }
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
