using Common;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BatchTranslateApp
{
    public static class TranslateHelper
    {
        public static string WordToKorean(string name, bool ignoreSsl = true)
        {
            if (ignoreSsl)
            {
                HttpClientHandler handler = new HttpClientHandler();
                handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;

                using (var client = new HttpClient(handler))
                {
                    var url = $"http://hanwenxingming.com/api/translate/chinese-name?name={name}";

                    var values = new Dictionary<string, string>
                    {
                        { "name", name }
                    };

                    var content = new FormUrlEncodedContent(values);

                    var response = client.PostAsync(url, content).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var responseString = response.Content.ReadAsStringAsync().Result;
                        var result = JsonConvert.DeserializeObject<TranslateResponseModel>(responseString);
                        return result.korean;
                    }
                    else
                    {
                        return "";
                    }
                }
            }
            else
            {
                var client = new RestClient("http://hanwenxingming.com/api/translate/chinese-name");
                var request = new RestRequest(Method.POST);
                request.AddOrUpdateParameter("name", name);
                var res = client.Execute(request);
                if (res.StatusCode == HttpStatusCode.OK)
                {
                    var result = JsonConvert.DeserializeObject<TranslateResponseModel>(res.Content);
                    return result.korean;
                }
                else
                {
                    return "";
                }
            }
        }

        public static string SentenceToKorean(string name, bool ignoreSsl = true)
        {
            string pattern = @"[\u4e00-\u9fa5]+";
            MatchCollection matches = Regex.Matches(name, pattern);
            string result = name;
            foreach (Match match in matches)
            {
                string chineseWords = match.Value;
                string koreanTranslation = WordToKorean(chineseWords, ignoreSsl);

                // 前面有号 后面也有号 就替换不了
                result = result.ReplaceFirst(chineseWords, koreanTranslation); // 只替换最开始的一个
            }

            return result;
        }

        public static string SentenceToKoreanIfJJB(string name, string gender)
        {
            name = name
                .Replace("家族/亲戚", "가족/친척")
                .Replace("知人（朋友/前后辈/同事）", "친구/선후배/직장동료")
                .Replace("路旁", "노방")
                .Replace("网络/通信", "인터넷/통신")
                .Replace("田地", "추수밭")
                .Replace("长老教", "장로교")
                .Replace("监理教", "감리교")
                .Replace("圣洁教", "성결교")
                .Replace("浸礼桥", "침례교")
                .Replace("纯福音", "순복음")
                .Replace("天主教", "천주교")
                .Replace("佛教", "불교")
                .Replace("无教", "무교")
                .Replace("scj(宣教教会)", "신천지(선교교회)")
                .Replace("摩门教", "몰몬교")
                .Replace("伊斯兰教", "이슬람교")
                .Replace("其他宗教", "기타종교")
                .Replace("三自教", "삼자교")
                .Replace("毕业", "졸업")
                .Replace("辍学", "중퇴")
                .Replace("在读", "재학")
                .Replace("大学在读", "대학교재학")
                .Replace("大学毕业", "대학교졸업")
                .Replace("大学辍学", "대학교중퇴")
                .Replace("研究生以上", "대학원이상")
                .Replace("神学院在读", "신학교재학")
                .Replace("神学院毕业", "신학교졸업")
                .Replace("神学院辍学", "신학교중퇴")
                .Replace("专科在读", "전문대재학")
                .Replace("专科毕业", "전문대졸업")
                .Replace("专科辍学", "전문대중퇴")
                .Replace("神学研究生在读", "신학대학원재학")
                .Replace("神学研究生毕业", "신학대학원졸업")
                .Replace("神学研究生辍学", "신학대학원중퇴")
                .Replace("大学", "대학교")
                .Replace("专科", "전문대")
                .Replace("研究生", "대학원")
                .Replace("父", "부")
                .Replace("母", "모")
                .Replace("丈夫", "남편")
                .Replace("夫人", "부인")
                .Replace("儿子", "아들")
                .Replace("女儿", "딸")
                .Replace("哥哥", gender == "남" ? "형": "오빠")
                .Replace("弟弟", "동생")
                .Replace("妹妹", "동생")
                .Replace("姐姐", gender == "남" ? "누나": "언니")
                .Replace("岳父", "장인")
                .Replace("岳母", "장모")
                .Replace("祖父", "조부")
                .Replace("祖母", "조모")
                .Replace("婆婆", "시어머니")
                .Replace("公公", "시아버지")
                .Replace("女婿", "사위")
                .Replace("儿媳", "며느리")
                .Replace("孙子", "손자");
            return SentenceToKorean(name, true);
        }
    }

    public class TranslateResponseModel
    {
        public string korean { get; set; }
    }
}