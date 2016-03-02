using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Inspector
{
    class ParseResult
    {
        public string value;
        public int ccount;
        public int jcount;

        public ParseResult(string value, int ccount, int jcount)
        {
            this.value = value;
            this.ccount = ccount;
            this.jcount = jcount;
        }
    }

    class Checker
    {
        private string fileName;
        private StringBuilder log;
        private ExcelWriter writer;
        private Dictionary<string, string> mapping;

        private static char[] punctuations_quanjiao = { '。', '、', '“', '”', '？', '！', '，', '；', '：' };
        private static char[] punctuations_banjiao = { ',', '?', '!', ';', '"', ':' };
        private static char[] punctuations_end_quanjiao = { '）' };
        private static char[] punctuations_end_banjiao = { ')' };
        
        public Checker(string fileName, ExcelWriter writer, string mappingConfFilePath)
        {
            this.fileName = fileName;
            this.writer = writer;
            this.log = new StringBuilder();
            this.mapping = LoadMappingConf(mappingConfFilePath);
        }

        public string Process()
        {
            log.AppendFormat("Processing file {0}.\n", this.fileName);
            Word.Application wordApp = null;
            Word.Document doc = null;
            object file = fileName;
            object nullobj = System.Reflection.Missing.Value;

            try 
            {
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(
                            ref file, ref nullobj, ref nullobj,
                            ref nullobj, ref nullobj, ref nullobj,
                            ref nullobj, ref nullobj, ref nullobj,
                            ref nullobj, ref nullobj, ref nullobj);
                doc.Activate();

                int tableNum = doc.Tables.Count;
                if(tableNum != 5)
                {
                    log.AppendFormat("[ERROR]There are {0} tables.\n", tableNum, this.fileName);
                    return log.ToString();
                }

                CheckFirstTable(doc.Tables[2]);
                CheckSecondTable(doc.Tables[3]);

                return log.ToString();
            }
            catch(Exception e)
            {
                log.AppendFormat("[ERROR]{0}.\n", e.ToString());
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(ref nullobj, ref nullobj, ref nullobj);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(ref nullobj, ref nullobj, ref nullobj);
                }
            }

            return log.ToString();
        }

        private Dictionary<string, string> LoadMappingConf(string mappingConfFilePath)
        {
            Dictionary<string, string> mapping = new Dictionary<string, string>();

            if (File.Exists(mappingConfFilePath))
            {
                using (var reader = new StreamReader(mappingConfFilePath, System.Text.Encoding.UTF8))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var segs = line.Split('\t');
                        if(segs.Length == 2)
                        {
                            mapping.Add(segs[0], segs[1]);            
                        }
                    }
                }
            }

            return mapping;
        }

        private void CheckFirstTable(Word.Table table)
        {
            string chinese = table.Cell(2, 2).Range.Text.ToString();
            string japanese = table.Cell(2, 3).Range.Text.ToString();

            CheckLine(chinese, japanese);
        }

        private void CheckSecondTable(Word.Table table)
        {
            for (int i = 2; i <= table.Rows.Count; i++)
            {
                string chinese = table.Cell(i, 2).Range.Text.ToString().TrimEnd(new char[] { '\r', '\a' });
                string japanese = table.Cell(i, 3).Range.Text.ToString().TrimEnd(new char[] { '\r', '\a' });

                CheckLine(chinese, japanese);
                //CheckMapping(chinese, japanese);
            }
        }

        //private void CheckMapping(string chinese, string japanese)
        //{
        //    var sb = new StringBuilder();

        //    this.mapping.Select(m =>
        //        {
        //            int chineseCount = Regex.Matches(chinese, m.Key).Count;
        //            int japaneseCount = Regex.Matches(japanese, m.Value).Count;

        //            return new Tuple<string, string, int, int>(m.Key, m.Value, chineseCount, japaneseCount);
        //        })
        //        .Where(p => p.Item3 != p.Item4 && p.Item3 > 0)
        //        .ToList()
        //        .ForEach(d =>
        //        {
        //            sb.AppendFormat("{0} => {1} {2}(中) {3} {4}(日) 不一致\n", d.Item1, d.Item2, d.Item3, d.Item3 > d.Item4 ? '>' : '<', d.Item4);
        //        });
            
        //    if (sb.Length > 0)
        //    {
        //        writer.WriteLine(new string[] { chinese, japanese, sb.ToString(), fileName });
        //    }

        //}

        private void CheckLine(string chinese, string japanese)
        {
            List<string> chineseParseResult = Extract(chinese);
            List<string> japaneseParseResult = Extract(japanese);
            List<ParseResult> result = Compare(chineseParseResult, japaneseParseResult);

            StringBuilder sb = new StringBuilder();
            foreach (var r in result)
            {
                sb.AppendFormat("{0} {1} (中) {2} {3} (日) 不一致\n", IsWhiteSpaceAll(r.value) ? "空格" : r.value, r.ccount, r.ccount > r.jcount ? '>' : '<', r.jcount);
            }

            foreach (var r in japaneseParseResult.Where(l => IsBanjiao(l)))
            {
                sb.AppendFormat("{0} (日)不是全角\n", IsWhiteSpaceAll(r) ? "空格" : r);
            }

            this.mapping.Select(m =>
                {
                    int chineseCount = Regex.Matches(chinese, m.Key).Count;
                    int japaneseCount = Regex.Matches(japanese, m.Value).Count;

                    return new Tuple<string, string, int, int>(m.Key, m.Value, chineseCount, japaneseCount);
                })
                .Where(p => p.Item3 != p.Item4 && p.Item3 > 0)
                .ToList()
                .ForEach(d =>
                {
                    sb.AppendFormat("{0} => {1} {2}(中) {3} {4}(日) 不一致\n", d.Item1, d.Item2, d.Item3, d.Item3 > d.Item4 ? '>' : '<', d.Item4);
                });

            if (sb.Length > 0)
            {
                writer.WriteLine(new string[] { chinese, japanese, sb.ToString(), fileName });
            }
        }

        private List<string> Extract(string text)
        {
            List<string> result = new List<string>();

            for (int i = 0; i < text.Length;i++)
            {
                if(text[i] <= 31)
                {
                    continue;
                }

                if (IsChinese(text[i]) || IsJapanese(text[i]))
                {
                    continue;
                }
                else if (IsPunctuation(text[i]))
                {
                    result.Add(text[i].ToString());
                }
                else if (IsWhiteSpaceAll(text[i]))
                {
                    result.Add(text[i].ToString());
                }
                else
                {
                    int end = i;
                    while (end < text.Length && !IsPhaseEnd(text[end]) && !IsWhiteSpaceAll(text[end]) && !IsPunctuation(text[end]) && !IsChinese(text[end]) && !IsJapanese(text[end]))
                    {
                        end++;
                    }
                    
                    if (IsPhaseEnd(text[end])) ++end;

                    result.Add(text.Substring(i, end - i));
                    i = end - 1;
                }
                //else if (IsDigitAll(text[i]))
                //{
                //    int end = i;
                //    while (end < text.Length && IsDigitAll(text[end])) end++;
                //    result.Add(text.Substring(i, end-i));
                //    i = end-1;
                //}
                //else
                //{
                //    result.Add(text[i].ToString());
                //}
            }

            return result;
        }

        private Dictionary<string, int> GetDict(List<string> l)
        {
            Dictionary<string, int> d = new Dictionary<string, int>();

            for (int i = 0; i < l.Count; i++)
            {
                string key = ToQuanjiao(l[i]);
                if (d.ContainsKey(key))
                {
                    var val = d[key];
                    d[key] = val+1;
                }
                else
                {
                    d[key] = 1;
                }
            }

            return d;
        }

        private List<ParseResult> Compare(List<string> chinese, List<string> japanese)
        {
            var result = new List<ParseResult>();

            var cd = GetDict(chinese);
            var jd = GetDict(japanese);

            foreach (var key in cd.Keys)
            {
                if (jd.ContainsKey(key))
                {
                    if (cd[key] != jd[key])
                    {
                        result.Add(new ParseResult(key, cd[key], jd[key]));
                    }
                }
                else
                {
                    result.Add(new ParseResult(key, cd[key], 0));
                }
            }

            foreach (var key in jd.Keys)
            {
                if (!cd.ContainsKey(key))
                {
                    result.Add(new ParseResult(key, 0, jd[key]));
                }
            }

            return result;
        }

        private bool IsChinese(char c)
        {
            return (c >= 0x4e00 && c <= 0x9fbb);
        }

        private bool IsJapanese(char c)
        {
            return (c >= 0x3040 && c <= 0x309F) || (c >= 0x30A0 && c <= 0x30FF);
        }

        private bool IsPunctuation(char c)
        {
            if(IsBanjiao(c))
            {
                return punctuations_banjiao.Any(p => p == c);
            }
            else
            {
                return punctuations_quanjiao.Any(p => p == c);
            }
        }

        private bool IsDigitAll(char c)
        {
            if(c >= '0' && c <= '9')
                return true;
            
            c = ToBanjiao(c);
            return (c >= '0' && c <= '9');
        }

        private bool IsPhaseEnd(char c)
        {
            if (IsBanjiao(c))
            {
                return punctuations_end_banjiao.Any(p => p == c);
            }
            else
            {
                return punctuations_end_quanjiao.Any(p => p == c);
            }
        }

        private bool IsQuanjiaoEqual(string cc, string jc)
        {
            string ccQuanjiao = ToQuanjiao(cc);
            string jcQuanjiao = ToQuanjiao(jc);

            return ccQuanjiao == jcQuanjiao;
        }

        private bool IsBanjiao(string line)
        {
            bool ret = true;
            foreach(char c in line)
            {
                if (c >= 127) ret = false;
            }

            return ret;
        }

        private bool IsBanjiao(char c)
        {
            return c < 127;
        }

        private bool IsWhiteSpaceAll(string s)
        {
            foreach (var c in s)
            {
                if (!IsWhiteSpaceAll(c))
                {
                    return false;
                }
            }

            return true;
        }

        private bool IsWhiteSpaceAll(char c)
        {
            return c == 32 || c == 12288;
        }

        // 半角转全角
        private string ToQuanjiao(string line)
        {
            StringBuilder r = new StringBuilder();

            foreach (char c in line)
            {
                // 空格单独处理
                if (c == 32)
                {
                    r.Append((char)12288);
                }
                else if (c < 127)
                {
                    r.Append((char)(c + 65248));
                }
                else
                {
                    r.Append(c);
                }
            }

            return r.ToString();
        }

        // 半角转全角
        private char ToQuanjiao(char c)
        {
            char ret;
            // 空格单独处理
            if (c == 32)
            {
                ret = ((char)12288);
            }
            else if (c < 127)
            {
                ret = ((char)(c + 65248));
            }
            else
            {
                ret = c;
            }

            return ret;
        }

        // 全角转半角
        private char ToBanjiao(char c)
        {
            if (c == 12288)
            {
                c = (char)32;
            }
            else if(c > 65280 && c < 65375)
            {
                c = (char)(c - 65248);
            }

            return c;
        }
    }
}
