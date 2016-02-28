using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

        public Checker(string fileName, ExcelWriter writer)
        {
            this.fileName = fileName;
            this.writer = writer;
            this.log = new StringBuilder();
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
            }
        }

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
                else
                {
                    int end = i;
                    while (end < text.Length && !IsChinese(text[end]) && !IsJapanese(text[end]))
                    {
                        end++;
                    }
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

        private bool IsDigitAll(char c)
        {
            if(c >= '0' && c <= '9')
                return true;
            
            c = ToBanjiao(c);
            return (c >= '0' && c <= '9');
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
