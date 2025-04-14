using Common;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace BatchTranslateApp
{
    public class JJBMember
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
        /// 工号
        /// </summary>
        [PropertyName("工号")]
        public string ScjNumber { get; set; }

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
    }
}
