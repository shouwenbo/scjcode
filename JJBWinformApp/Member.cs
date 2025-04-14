using Common;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace JJBWinformApp
{
    public class Member
    {
        #region
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 列名称
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 排序
        /// </summary>
        [PropertyName("排序")]
        public string Index { get; set; }

        /// <summary>
        /// 韩语姓名
        /// </summary>
        [PropertyName("韩语姓名")]
        public string KoreanName { get; set; }

        /// <summary>
        /// 汉语姓名
        /// </summary>
        [PropertyName("汉语姓名")]
        public string ChineseName { get; set; }

        /// <summary>
        /// 汉语姓氏
        /// </summary>
        public string ChineseFirstName
        {
            get
            {
                if (ChineseName.Length == 4)
                {
                    return ChineseName[..2];
                }
                if (ChineseName.Length == 2 || ChineseName.Length == 3)
                {
                    return ChineseName[..1];
                }
                return "";
            }
        }

        /// <summary>
        /// 国籍
        /// </summary>
        [PropertyName("国籍")]
        public string Country { get; set; }

        /// <summary>
        /// 身份证生日
        /// </summary>
        [PropertyName("身份证生日")]
        public string IDCardBirthday { get; set; }

        /// <summary>
        /// 性别
        /// </summary>
        [PropertyName("性别")]
        public string Gender { get; set; }

        /// <summary>
        /// 实际生日
        /// </summary>
        [PropertyName("实际生日")]
        public string ActualBirthday { get; set; }

        /// <summary>
        /// 身份证地址中文
        /// </summary>
        [PropertyName("身份证地址中文")]
        public string IDCardAddressChinese { get; set; }

        /// <summary>
        /// 身份证地址翻译
        /// </summary>
        [PropertyName("身份证地址翻译")]
        public string IDCardAddressKorean { get; set; }

        /// <summary>
        /// 现住址
        /// </summary>
        [PropertyName("现住址中文")]
        public string CurrentAddressChinese { get; set; }

        /// <summary>
        /// 现住址翻译
        /// </summary>
        [PropertyName("现住址翻译")]
        public string CurrentAddressKorean { get; set; }

        /// <summary>
        /// 电话
        /// </summary>
        [PropertyName("电话")]
        public string Phone { get; set; }

        /// <summary>
        /// 邮箱
        /// </summary>
        [PropertyName("邮箱")]
        public string Email { get; set; }

        /// <summary>
        /// 身高
        /// </summary>
        [PropertyName("身高")]
        public string Height { get; set; }

        /// <summary>
        /// 血型
        /// </summary>
        [PropertyName("血型")]
        public string BloodType { get; set; }

        /// <summary>
        /// 婚否
        /// </summary>
        [PropertyName("婚否")]
        public string IsMarried { get; set; }

        /// <summary>
        /// 是否结婚了
        /// </summary>
        public bool IsMarry
        {
            get
            {
                return IsMarried == "기혼";
            }
        }

        /// <summary>
        /// CD类型
        /// </summary>
        [PropertyName("CD类型")]
        public string CDType { get; set; }

        /// <summary>
        /// CD类型批注
        /// </summary>
        public IComment CDTypeComment { get; set; }

        /// <summary>
        /// CD类型批注
        /// </summary>
        [PropertyName("CD类型批注")]
        public string? CDTypeCommentString { get; set; }

        /// <summary>
        /// 出身教团
        /// </summary>
        [PropertyName("出身教团")]
        public string BirthReligion { get; set; }

        /// <summary>
        /// 其他宗教
        /// </summary>
        [PropertyName("其他宗教")]
        public string OtherReligion { get; set; }

        /// <summary>
        /// 教会名
        /// </summary>
        [PropertyName("教会名")]
        public string ChurchName { get; set; }

        /// <summary>
        /// 职分
        /// </summary>
        [PropertyName("职分")]
        public string Position { get; set; }

        /// <summary>
        /// 信仰时间
        /// </summary>
        [PropertyName("信仰时间")]
        public string YearsOfFaith { get; set; }

        /// <summary>
        /// 母胎信仰
        /// </summary>
        [PropertyName("母胎信仰")]
        public string MaternalFaith { get; set; }

        /// <summary>
        /// 最终学历
        /// </summary>
        [PropertyName("最终学历")]
        public string FinalEducation { get; set; }

        public int FinalEducationLevel
        {
            get
            {
                return FinalEducation switch
                {
                    "无学" => 1,
                    "小学在读" => 5,
                    "小学辍学" => 10,
                    "小学毕业" => 15,
                    "初中在读" => 20,
                    "初中辍学" => 25,
                    "初中毕业" => 30,
                    "高中在读" => 35,
                    "高中辍学" => 40,
                    "高中毕业" => 45,
                    "专科在读" => 50,
                    "专科辍学" => 55,
                    "专科毕业" => 60,
                    "大学在读" => 65,
                    "大学辍学" => 70,
                    "大学毕业" => 75,
                    "研究生以上" => 80,
                    "神学院在读" => 85,
                    "神学院辍学" => 90,
                    "神学院毕业" => 95,
                    "神学研究生在读" => 100,
                    "神学研究生辍学" => 105,
                    "神学研究生毕业" => 110,
                    _ => 0,
                };
            }
        }

        /// <summary>
        /// 学历信息列表
        /// </summary>
        public List<EducationInfo> EducationList { get; set; }

        [PropertyName("职业分类")]
        public string JobType { get; set; }

        [PropertyName("其他职业")]
        public string JobOther { get; set; }

        [PropertyName("工作单位名称")]
        public string JobCompany { get; set; }

        [PropertyName("职位")]
        public string JobName { get; set; }

        /// <summary>
        /// 兄弟姐妹关系
        /// </summary>
        [PropertyName("兄弟姐妹关系")]
        public string SiblingRelationship { get; set; }

        /// <summary>
        /// 家族事项列表
        /// </summary>
        public List<FamilyMemberInfo> FamilyMemberList { get; set; }

        /// <summary>
        /// 列表名称
        /// </summary>
        public string ListName
        {
            get
            {
                return $"{ColumnName,3}: {Index} {ChineseName}({KoreanName})";
            }
        }

        /// <summary>
        /// 验证错误
        /// </summary>
        public bool ValidateError { get; set; } = false;

        #endregion

        /// <summary>
        /// 检查
        /// </summary>
        public List<string> Check()
        {
            var result = new List<string>();
            if (!Index.IsPositiveInteger())
            {
                result.Add($"排序:{Index} 必须为正整数");
            }
            if (!KoreanName.IsAllKorean())
            {
                result.Add($"韩语名称:{KoreanName} 包含不为韩文的字符");
            }
            if (Country != "중화인민공화국")
            {
                result.Add($"国籍信息有误");
            }
            if (!IDCardBirthday.IsSixDateFormat())
            {
                result.Add($"身份证生日:{IDCardBirthday} 格式有误");
            }
            if (Gender!= "남" && Gender!= "여")
            {
                result.Add($"性别:{Gender} 有误");
            }
            #region 实际生日
            var actualAge = 0;
            if (!ActualBirthday.IsActualDateFormat())
            {
                result.Add($"实际生日:{ActualBirthday} 格式有误");
                
            }
            try
            {
                actualAge = DateTime.Now.Year - int.Parse(ActualBirthday[..4]);
            }
            catch
            {
                result.Add($"实际生日:{ActualBirthday} 格式异常");
            }
            #endregion
            #region 地址检查

            if (!IDCardAddressChinese.IsChineseAddress())
            {
                result.Add($"身份证地址中文:{IDCardAddressChinese} 包含异常字符");
            }

            // 155/155-1/155-2 开始不再校验地址翻译，因为我会统一进行翻译
            // if (!IDCardAddressKorean.IsKoreanAddress())
            // {
            //     result.Add($"身份证地址翻译:{IDCardAddressKorean} 包含异常字符");
            // }

            //var cardAddrCnSpaceCount = IDCardAddressChinese.CountSpaces();
            //var cardAddrKoSpaceCount = IDCardAddressKorean.CountSpaces();
            //if (cardAddrCnSpaceCount != cardAddrKoSpaceCount)
            //{
            //    result.Add($"身份证地址:{IDCardAddressChinese}({IDCardAddressKorean}) 空格数不一致({cardAddrCnSpaceCount}<>{cardAddrKoSpaceCount}), 请进一步确认");
            //}

            if (!CurrentAddressChinese.IsChineseAddress())
            {
                result.Add($"现住址中文:{CurrentAddressChinese} 包含异常字符");
            }

            // 155/155-1/155-2 开始不再校验地址翻译，因为我会统一进行翻译
            // if (!CurrentAddressKorean.IsKoreanAddress())
            // {
            //     result.Add($"现住址翻译:{CurrentAddressKorean} 包含异常字符");
            // }

            //var currAddrCnSpaceCount = CurrentAddressChinese.CountSpaces();
            //var currAddrKoSpaceCount = CurrentAddressKorean.CountSpaces();
            //if (currAddrCnSpaceCount != currAddrKoSpaceCount)
            //{
            //    result.Add($"现住址:{CurrentAddressChinese}({CurrentAddressKorean}) 空格数不一致({currAddrCnSpaceCount}<>{currAddrKoSpaceCount}), 请进一步确认");
            //}

            #endregion
            if (!Phone.IsPhoneNumberAndFormatValid())
            {
                result.Add($"电话:{Phone} 格式有误");
            }
            if (!Email.IsNoneOrValidEmail())
            {
                result.Add($"邮箱:{Email} 不符合规范");
            }
            if (!Height.IsReasonableHeight(out string heightMsg))
            {
                result.Add($"身高:{Height} {heightMsg}");
            }
            if (!BloodType.IsBloodTypeValid())
            {
                result.Add($"血型:{BloodType} 不符合规范");
            }
            if (!IsMarried.IsMaritalStatusValid())
            {
                result.Add($"婚否:{IsMarried} 不符合规范");
            }
            #region CD类型
            if (CDType == "家族/亲戚" || CDType == "网络/通信" || CDType == "知人（朋友/前后辈/同事）")
            {
                if (string.IsNullOrEmpty(CDTypeCommentString))
                {
                    result.Add($"CD类型:{CDType} 必须填写批注");
                }
                else if (CDType == "家族/亲戚" && (CDTypeCommentString != "家族" && CDTypeCommentString != "亲戚"))
                {
                    result.Add($"CD类型:{CDType} 批注:{CDTypeCommentString} 不符合规范, 必须是家族或亲戚");
                }
                else if (CDType == "网络/通信" && (CDTypeCommentString != "网络" && CDTypeCommentString != "通信"))
                {
                    result.Add($"CD类型:{CDType} 批注:{CDTypeCommentString} 不符合规范, 必须是网络或通信");
                }
                else if (CDType == "知人（朋友/前后辈/同事）" && (CDTypeCommentString != "朋友" && CDTypeCommentString != "前后辈" && CDTypeCommentString != "同事"))
                {
                    result.Add($"CD类型:{CDType} 批注:{CDTypeCommentString} 不符合规范, 必须是朋友或前后辈或同事");
                }
            }
            else if (CDType == "路旁" || CDType == "田地" || CDType == "职场)")
            {
                if (!string.IsNullOrEmpty(CDTypeCommentString))
                {
                    result.Add($"CD类型:{CDType} 不需要填写批注");
                }
            }
            else
            {
                result.Add($"CD类型:{CDType} 不符合规范");
            }
            #endregion
            #region 信仰情况
            if (BirthReligion != "三自教" && BirthReligion != "无教" && BirthReligion != "天主教" && BirthReligion != "其他宗教")
            {
                result.Add($"出身教团:{BirthReligion} 不符合规范");
            }
            if (BirthReligion == "无教")
            {
                if (!string.IsNullOrEmpty(YearsOfFaith))
                {
                    result.Add($"出身教团:{BirthReligion} 不需要填写信仰年数");
                }
                if (!string.IsNullOrEmpty(ChurchName))
                {
                    result.Add($"出身教团:{BirthReligion} 不需要填写教会名");
                }
            }
            if ((BirthReligion == "三自教" || BirthReligion == "天主教" || BirthReligion == "无教") && !string.IsNullOrEmpty(OtherReligion))
            {
                result.Add($"出身教团:{BirthReligion} 不需要填写其他宗教");
            }
            if (BirthReligion == "其他宗教")
            {
                if (string.IsNullOrEmpty(OtherReligion))
                {
                    result.Add($"出身教团:{BirthReligion} 必须填写其他宗教");
                }
            }
            if (BirthReligion == "三自教" || BirthReligion == "天主教" || BirthReligion == "其他宗教")
            {
                Match yearsOfFaithMatch = new Regex(@"^(\d+)年$").Match(YearsOfFaith);
                if (yearsOfFaithMatch.Success)
                {
                    int faithYear = int.Parse(yearsOfFaithMatch.Groups[1].Value);
                    if (faithYear < 1)
                    {
                        result.Add($"信仰年数:{YearsOfFaith} 至少写1年");
                    }
                    if (faithYear > actualAge)
                    {
                        result.Add($"信仰年数:{YearsOfFaith} 大于实际岁数{actualAge}岁");
                    }
                    else
                    {
                        if (MaternalFaith == "모태신앙" && ((actualAge - faithYear) >= 1))
                        {
                            result.Add($"母胎信仰，但是信仰年数是‘{YearsOfFaith}’，应该和年龄一致");
                        }
                    }
                }
                else
                {
                    result.Add($"信仰年数:{YearsOfFaith} 格式不符合规范");
                }

                if (string.IsNullOrEmpty(ChurchName))
                {
                    result.Add($"出身教团:{BirthReligion} 必须填写教会名");
                }
                else if (ChurchName.Length < 2)
                {
                    result.Add($"教会名:{ChurchName} 字数太少");
                }
                else if (ChurchName.Length > 4)
                {
                    var churchNameMsg = $"教会名:{ChurchName} 字数太长";
                    if (ChurchName.Contains("三自教会"))
                    {
                        churchNameMsg += " 建议去掉'三自教会'";
                    }
                    if (ChurchName.Contains("基督教会"))
                    {
                        churchNameMsg += " 建议去掉'基督教会'";
                    }
                    else if (ChurchName.Contains("基督教堂"))
                    {
                        churchNameMsg += " 建议去掉'基督教堂'";
                    }
                    else if (ChurchName.Contains("基督教"))
                    {
                        churchNameMsg += " 建议去掉'基督教'";
                    }
                    else if (ChurchName.Contains("教会"))
                    {
                        churchNameMsg += " 建议去掉'教会'";
                    }
                    else if (ChurchName.Contains("教堂"))
                    {
                        churchNameMsg += " 建议去掉'教堂'";
                    }
                    result.Add(churchNameMsg);
                }
                else
                {
                    if (ChurchName.Contains("基督教会"))
                    {
                        result.Add($"教会名:{ChurchName} 建议去掉'基督教会");
                    }
                    else if (ChurchName.Contains("基督教堂"))
                    {
                        result.Add($"教会名:{ChurchName} 建议去掉'基督教堂");
                    }
                    else if (ChurchName.Contains("基督教"))
                    {
                        result.Add($"教会名:{ChurchName} 建议去掉'基督教");
                    }
                    else if (ChurchName.Contains("教会"))
                    {
                        result.Add($"教会名:{ChurchName} 建议去掉'教会");
                    }
                    else if (ChurchName.Contains("教堂"))
                    {
                        result.Add($"教会名:{ChurchName} 建议去掉'教堂");
                    }
                }
            }
            if (!string.IsNullOrEmpty(Position))
            {
                result.Add($"职分:{Position} 无需填写");
            }
            if (MaternalFaith != "모태신앙" && MaternalFaith != "非모태신앙")
            {
                result.Add($"母胎信仰:{Position} 不符合填写规范");
            }
            #endregion
            #region 兄弟姐妹关系
            try
            {
                Match shipMatch = new Regex(@"^(\d*男)?\s*(\d*女)?\s*中\s*第(\d+)$").Match(SiblingRelationship);
                if (shipMatch.Success)
                {
                    string[] shipParts = SiblingRelationship.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    if (shipParts.Length == 4)
                    {
                        var maleStr = shipParts[0];
                        var femaleStr = shipParts[1];
                        var shipZhong = shipParts[2];
                        var sortStr = shipParts[3];

                        if (maleStr.Contains('男'))
                        {
                            if (femaleStr.Contains('女'))
                            {
                                if (shipZhong == "中")
                                {
                                    if (sortStr.Contains('第'))
                                    {
                                        try
                                        {
                                            var maleCount = maleStr.Replace("男", "") == "" ? 0 : int.Parse(maleStr.Replace("男", ""));
                                            var femaleCount = femaleStr.Replace("女", "") == "" ? 0 : int.Parse(femaleStr.Replace("女", ""));
                                            var sort = int.Parse(sortStr.Replace("第", ""));
                                            if ((maleCount + femaleCount) > 0)
                                            {
                                                if (sort <= (maleCount + femaleCount))
                                                {
                                                    if (sort > 0)
                                                    {

                                                    }
                                                    else
                                                    {
                                                        result.Add($"兄弟姐妹关系:{SiblingRelationship} 排序不存在");
                                                    }
                                                }
                                                else
                                                {
                                                    result.Add($"兄弟姐妹关系:{SiblingRelationship} 排序超出总人数");
                                                }
                                            }
                                            else
                                            {
                                                result.Add($"兄弟姐妹关系:{SiblingRelationship} 至少有1人");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            result.Add($"兄弟姐妹关系:{SiblingRelationship} 数字转换错误 {ex.Message}");
                                        }
                                    }
                                    else
                                    {
                                        result.Add($"兄弟姐妹关系:{SiblingRelationship} 排序格式错误");
                                    }
                                }
                                else
                                {
                                    result.Add($"兄弟姐妹关系:{SiblingRelationship} 格式错误");
                                }
                            }
                            else
                            {
                                result.Add($"兄弟姐妹关系:{SiblingRelationship} 女性格式错误");
                            }
                        }
                        else
                        {
                            result.Add($"兄弟姐妹关系:{SiblingRelationship} 男性格式错误");
                        }
                    }
                    else
                    {
                        result.Add($"兄弟姐妹关系:{SiblingRelationship} 格式不正确");
                    }
                }
                else
                {
                    result.Add($"兄弟姐妹关系:{SiblingRelationship} 格式不符合规范");
                }
            }
            catch (Exception)
            {

                throw;
            }
            #endregion

            #region 学历

            Match finalEducationMatch = new Regex(@"^(无学|小学在读|小学辍学|小学毕业|初中在读|初中辍学|初中毕业|高中在读|高中辍学|高中毕业|专科在读|专科辍学|专科毕业|大学在读|大学辍学|大学毕业|研究生以上|神学院在读|神学院辍学|神学院毕业|神学研究生在读|神学研究生辍学|神学研究生毕业)$").Match(FinalEducation);
            if (finalEducationMatch.Success)
            {
                if (FinalEducationLevel >= 50) // 专科以上
                {
                    var lastEdu = EducationList.LastOrDefault();

                    if (lastEdu != null)
                    {
                        if (FinalEducation.Contains("专科") && lastEdu.Category != "专科")
                        {
                            result.Add($"最终学历:{FinalEducation} 和学历列表最后一个学历:{lastEdu.Category} 不相符");
                        }
                        if (FinalEducation.Contains("大学") && (lastEdu.Category != "大学" && lastEdu.Category != "专科")) // 跟蒋组长确认过, 有些大学是包含专科的
                        {
                            result.Add($"最终学历:{FinalEducation} 和学历列表最后一个学历:{lastEdu.Category} 不相符");
                        }
                        if (FinalEducation.Contains("研究生") && lastEdu.Category != "研究生")
                        {
                            result.Add($"最终学历:{FinalEducation} 和学历列表最后一个学历:{lastEdu.Category} 不相符");
                        }
                    }
                    else
                    {
                        result.Add($"最终学历:{FinalEducation} 但是学历列表为空");
                    }
                }
            }
            else
            {
                result.Add($"最终学历:{FinalEducation} 不符合规范");
            }

            if (!EducationList.CheckAsc(p => p.Level))
            {
                result.Add($"学历列表:{FinalEducation} 没有按照时间顺序从小到大排序");
            }

            for (int i = 0; i < EducationList.Count; i++)
            {
                result.AddRange(EducationList[i].Check());
            }

            #endregion

            #region 家族

            if (!FamilyMemberList.CheckDesc(p => p.Birth))
            {
                result.Add($"家族列表:{FamilyMemberList} 没有按照年龄从大到小排序");
            }

            if (FamilyMemberList.Count > 2)
            {
                result.Add($"家族人数大于2位, 需要确认一下规则");
            }

            for (int i = 0; i < FamilyMemberList.Count; i++)
            {
                result.AddRange(FamilyMemberList[i].Check(ChineseFirstName, Gender, IsMarry));
            }

            #endregion

            return result;
        }
    }

    /// <summary>
    /// 学历信息
    /// </summary>
    public class EducationInfo
    {
        /// <summary>
        /// 信息序号
        /// </summary>
        public int Index { get; set; }

        public int Level
        {
            get
            {
                return Category switch
                {
                    "小学" => 3,
                    "初中" => 5,
                    "高中" => 7,
                    "专科" => 9,
                    "大学" => 11,
                    "研究生" => 13,
                    "博士" => 15,
                    _ => 0,
                };
            }
        }

        /// <summary>
        /// 分类
        /// </summary>
        [PropertyName("分类")]
        public string Category { get; set; }

        /// <summary>
        /// 学校名
        /// </summary>
        [PropertyName("学校名")]
        public string SchoolName { get; set; }

        /// <summary>
        /// 专业
        /// </summary>
        [PropertyName("专业")]
        public string Major { get; set; }

        /// <summary>
        /// 状态
        /// </summary>
        [PropertyName("状态")]
        public string Status { get; set; }

        /// <summary>
        /// 期间
        /// </summary>
        [PropertyName("期间")]
        public string Period { get; set; }

        public List<string> Check()
        {
            var result = new List<string>();
            if (Level > 0)
            {
                if (string.IsNullOrEmpty(SchoolName.Trim()))
                {
                    result.Add($"{Category}学历 -> 学校名:{SchoolName} 不能为空");
                }
                if (string.IsNullOrEmpty(Major.Trim()))
                {
                    result.Add($"{Category}学历 -> 专业:{Major} 不能为空");
                }
                if (Status != "毕业" && Status != "缀学" && Status != "在读")
                {
                    result.Add($"{Category}学历 -> 状态:{Status} 不符合规范");
                }
                if (Level < 9)
                {
                    if (!string.IsNullOrEmpty(Period))
                    {
                        result.Add($"{Category}学历 -> 期间:{Period} 高中及以下学历不需要填写区间");
                    }
                }
                if (Level >= 9)
                {
                    if (!string.IsNullOrEmpty(Period))
                    {
                        if (Period.Contains('-'))
                        {
                            string[] parts = Period.Split('-');
                            if (parts.Length == 2)
                            {
                                try
                                {
                                    var startYear = int.Parse(parts[0].Split('.')[0]);
                                    var startMonth = int.Parse(parts[0].Split('.')[1]);
                                    var startDate = new DateTime(startYear, startMonth, 1);
                                    if (parts[1] != "") // 有结束日期
                                    {
                                        var endYear = int.Parse(parts[1].Split('.')[0]);
                                        var endMonth = int.Parse(parts[1].Split('.')[1]);
                                        var endDate = new DateTime(endYear, endMonth, 1);
                                        if (startDate >= endDate)
                                        {
                                            result.Add($"{Category}学历 -> 期间:{Period} 开始时间大于结束时间");
                                        }
                                    }
                                    else // 无结束日期
                                    {
                                        if (Status != "在读")
                                        {
                                            result.Add($"{Category}学历 -> 期间:{Period} 的状态:{Status} 不为在读");
                                        }
                                    }
                                }
                                catch
                                {
                                    result.Add($"{Category}学历 -> 期间:{Period} 格式不规范");
                                }
                            }
                            else
                            {
                                result.Add($"{Category}学历 -> 期间:{Period} 格式不正确");
                            }
                        }
                        else
                        {
                            result.Add($"{Category}学历 -> 期间:{Period} 格式不正确");
                        }
                    }
                    else
                    {
                        result.Add($"{Category}学历 -> 未填写学历期间");
                    }
                }
            }
            else
            {
                result.Add($"学历分类:{Category} 不符合规范");
            }
            return result;
        }
    }

    /// <summary>
    /// 家族事项
    /// </summary>
    public class FamilyMemberInfo
    {
        /// <summary>
        /// 事项序号
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        [PropertyName("姓名")]
        public string Name { get; set; }

        /// <summary>
        /// 姓氏
        /// </summary>
        public string FirstName
        {
            get
            {
                if (Name.Length == 4)
                {
                    return Name[..2];
                }
                if (Name.Length == 2 || Name.Length == 3)
                {
                    return Name[..1];
                }
                return "";
            }
        }

        /// <summary>
        /// 性别
        /// </summary>
        [PropertyName("性别")]
        public string Gender { get; set; }

        /// <summary>
        /// 生日
        /// </summary>
        [PropertyName("生日")]
        public string Birthday { get; set; }

        /// <summary>
        /// 生日
        /// </summary>
        public DateTime Birth { get; set; }

        /// <summary>
        /// 关系
        /// </summary>
        [PropertyName("关系")]
        public string Relationship { get; set; }

        /// <summary>
        /// 教团
        /// </summary>
        [PropertyName("教团")]
        public string Religion { get; set; }

        /// <summary>
        /// 教团-其他
        /// </summary>
        [PropertyName("教团-其他")]
        public string OtherReligion { get; set; }

        /// <summary>
        /// 出身教会
        /// </summary>
        [PropertyName("出身教会")]
        public string BirthChurch { get; set; }

        /// <summary>
        /// SCJ工号
        /// </summary>
        [PropertyName("SCJ工号")]
        public string SCJNumber { get; set; }

        public List<string> Check(string bysFirstName, string bysGender, bool bysIsMarry)
        {
            var result = new List<string>();
            Match nameMatch = new Regex(@"^[\u4e00-\u9fa5]{2,4}$").Match(Name);
            if (!nameMatch.Success)
            {
                result.Add($"家属:{Name} -> 姓名:{Name} 不符合规范");
            }

            if (Gender != "남" && Gender != "녀") // 여 녀 都代表女
            {
                result.Add($"家属:{Name} -> 性别:{Gender} 有误");
            }

            #region 关系

            Match shipMatch = new Regex(@"^(父|母|丈夫|夫人|前夫|前妻|儿子|女儿|哥哥|弟弟|姐姐|妹妹|岳父|岳母|祖父|祖母|婆婆|公公|女婿|儿媳|孙子)$").Match(Relationship);
            if (!shipMatch.Success)
            {
                result.Add($"家属:{Name} -> 关系:{shipMatch} 不符合规范");
            }

            if (bysGender == "남" && (Relationship == "父" || Relationship == "儿子" || Relationship == "女儿" || Relationship == "孙子") && bysFirstName != FirstName)
            {
                result.Add($"家属:{Name} -> 关系:{shipMatch} 的姓氏:{FirstName} 需要进一步确认, 学员姓氏:{bysFirstName},性别:{bysGender}");
            }

            if (bysGender == "여" && Relationship == "父" && bysFirstName != FirstName)
            {
                result.Add($"家属:{Name} -> 关系:{shipMatch} 的姓氏:{FirstName} 需要进一步确认, 学员姓氏:{bysFirstName},性别:{bysGender}");
            }

            if (Gender == "남" && !new Regex(@"^(父|丈夫|前夫|儿子|哥哥|弟弟|岳父|祖父|公公|女婿|孙子)$").Match(Relationship).Success)
            {
                result.Add($"家属:{Name} -> 关系:{shipMatch} 的性别:{Gender} 有误");
            }

            if (Gender == "녀" && !new Regex(@"^(母|夫人|前妻|女儿|姐姐|妹妹|岳母|祖母|婆婆|儿媳)$").Match(Relationship).Success)
            {
                result.Add($"家属:{Name} -> 关系:{shipMatch} 的性别:{Gender} 有误");
            }

            if (!bysIsMarry && new Regex(@"^(丈夫|夫人|前夫|前妻|儿子|女儿|女婿|儿媳|孙子)$").Match(Relationship).Success)
            {
                result.Add($"家属:{Name} -> 关系:{shipMatch} 有误, 与婚否状态不一致");
            }

            #endregion

            if (!string.IsNullOrEmpty(Birthday) && !Birthday.IsValidPointDate(out DateTime birth))
            {
                Birth = birth;
                result.Add($"家属:{Name} -> 生日:{Birthday} 格式错误");
            }

            Match religionMatch = new Regex(@"^(其他|三自教|scj)?$").Match(Religion);
            if (!religionMatch.Success)
            {
                result.Add($"家属:{Name} -> 教团:{Religion} 不符合规范");
            }
            else
            {
                if (string.IsNullOrEmpty(Religion))
                {
                    result.Add($"家属:{Name} -> 教团:{Religion} 未填写, 需要提醒一下是否漏填");
                }
            }
            if (Religion == "其他")
            {
                if (string.IsNullOrEmpty(OtherReligion))
                {
                    result.Add($"家属:{Name} -> 教团:{Religion} 必须填写 教团-其他");
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(OtherReligion))
                {
                    result.Add($"家属:{Name} -> 教团:{Religion} 无需填写 教团-其他");
                }
            }
            if (Religion == "三自教")
            {
                if (string.IsNullOrEmpty(BirthChurch))
                {
                    result.Add($"家属:{Name} -> 教团:{Religion} 必须填写出身教会");
                }
            }
            if (Religion == "scj")
            {
                Match scjMatch = new Regex(@"^\d{8}-\d{5}$").Match(SCJNumber);
                if (!scjMatch.Success && !SCJNumber.Contains('期'))
                {
                    result.Add($"家属:{Name} -> 教团:{Religion} 工号格式错误");
                }

                if (string.IsNullOrEmpty(BirthChurch))
                {
                    result.Add($"家属:{Name} -> 教团:{Religion} 必须填写出身教会");
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(SCJNumber))
                {
                    result.Add($"家属:{Name} -> 教团:{Religion} 无需填写工号");
                }
            }
            
            return result;
        }
    }
}
