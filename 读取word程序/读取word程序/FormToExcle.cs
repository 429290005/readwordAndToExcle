using BanMa.Model;
using cn.bmob.io;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 读取word程序
{
    public partial class FormToExcle : BanMa.BmobBasePCForm
    {
        public FormToExcle()
        {
            
            InitializeComponent();
            bomoToexcel();
           
        }
     
        private void bomoToexcel(){
            DataTable _px = new DataTable();
           
            var query = new BmobQuery();
            query.WhereGreaterThan("indexID", 2651);      //分数大于60岁
            query.OrderBy("indexID");
            query.Limit(600);
            var childfuture = Bmob.FindTaskAsync<computerTiKu>("computerTiKu", query);
       
            if (childfuture.Result.results.Count != 0)
            {
                _px = ToDataTable<computerTiKu>(childfuture.Result.results);
              //  MessageBox.Show("总行数" + _px.Rows.Count);
            }
            
            //_dspx = AdDispatchManage.GetValue(sqlstr);

            int row;
            row = 2;
            int b;
            b = 0;
            string PathStr;
            string SourceFileName;
            string DestinationFileName;
            PathStr = Application.StartupPath.Trim();
            // MakeBigWord.MakeBigWord.KillExcelProcess();
            DestinationFileName = PathStr + @"\plantmp.xls";
            SourceFileName = PathStr + @"\yhygzpx.xls";
            try
            {
                File.Delete(DestinationFileName);
                File.Copy(SourceFileName, DestinationFileName);
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                return;
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlBook = xlApp.Workbooks.Add(DestinationFileName);
            Microsoft.Office.Interop.Excel.Worksheet xlSheet = xlBook.Worksheets[1];
            xlSheet.Activate();
            xlSheet.Application.Visible = true;
            try
            {
                for (int i = 0; i < _px.Rows.Count; i++)
                {
                    if(_px.Rows[i]["titleSubject"].ToString()!=null){
                        xlSheet.Cells[i + row, 1] = _px.Rows[i]["titleSubject"].ToString();
                    }else{
                        MessageBox.Show(i + "titleSubject");
                    }
                    if (_px.Rows[i]["titleSubject"].ToString() != null)
                    {
                        xlSheet.Cells[i + row, 1] = _px.Rows[i]["titleSubject"].ToString();
                    }
                    else
                    {
                        MessageBox.Show(i + "titleSubject");
                    }
                    if (_px.Rows[i]["optionA"].ToString() != null && _px.Rows[i]["optionB"].ToString() != null && _px.Rows[i]["optionC"].ToString() != null && _px.Rows[i]["optionD"].ToString() != null)
                    {
                        Choice[] student = new Choice[4];
                        student[0] = new Choice();
                        student[0].key = "A";
                        student[0].content = _px.Rows[i]["optionA"].ToString();
                        student[1] = new Choice();
                        student[1].key = "B";
                        student[1].content = _px.Rows[i]["optionB"].ToString();
                        student[2] = new Choice();
                        student[2].key = "C";
                        student[2].content = _px.Rows[i]["optionC"].ToString();
                        student[3] = new Choice();
                        student[3].key = "D";
                        student[3].content = _px.Rows[i]["optionD"].ToString();
                        string aa = Newtonsoft.Json.JsonConvert.SerializeObject(student);
                        xlSheet.Cells[i + row, 2] = aa;
                        //设置一个Person类
                        //Choice p = new Choice();
                        //p.key = _px.Rows[i]["optionA"].ToString();
                        //p.content = _px.Rows[i]["optionB"].ToString();
                        //p.C = _px.Rows[i]["optionC"].ToString();
                        //p.D = _px.Rows[i]["optionD"].ToString();
                     
                        //string json1 = JsonConvert.SerializeObject(p);
                        //Console.WriteLine(json1 + "\n");
                        ////缩进输出
                        
                        //string json2 = JsonConvert.SerializeObject(p, Formatting.Indented);
                        //Console.WriteLine(json2 + "\n");

                      // {"A":"+价格发现功能+","B":"+套利功能+","C":"+投机功能+","D":"+套期保值功能"}
                       
                    }
                    else
                    {
                        MessageBox.Show(i + "A=" + _px.Rows[i]["optionA"].ToString() + "B=" + _px.Rows[i]["optionB"].ToString() + "C=" + _px.Rows[i]["optionC"].ToString() + "D=" + _px.Rows[i]["optionD"].ToString());
                    }

                    xlSheet.Cells[i + row, 3] = "3RcsJ77J";
                  if (_px.Rows[i]["answer"].ToString() != null)
                  {
                      String title = _px.Rows[i]["answer"].ToString().Substring(4, _px.Rows[i]["answer"].ToString().Length - 4); ///去掉两个字符串
                      xlSheet.Cells[i + row, 4] = "[\"" + title + "\"]";
                     
                  }
                   else
                   {
                       MessageBox.Show(i + "answer");
                    }
                    if (_px.Rows[i]["explain"].ToString() != null)
                    {
                        xlSheet.Cells[i + row, 5] = _px.Rows[i]["explain"].ToString();
                    }
                    else
                    {
                        MessageBox.Show(i + "explain");
                    }
                  
                        xlSheet.Cells[i + row, 6] = 500+i;
                  
                        xlSheet.Cells[i + row, 7] = "P0Vm666E";
                   
                    b = i;
                    
                    //xlSheet.Cells[i + row, 2] = _px.Rows[i]["indexID"].ToString();
                    //xlSheet.Cells[i + row, 3] = _px.Rows[i]["titleSubject"].ToString();
                    //xlSheet.Cells[i + row, 4] = _px.Rows[i]["answer"].ToString();
                    //xlSheet.Cells[i + row, 5] = _px.Rows[i]["optionA"].ToString();
                    //xlSheet.Cells[i + row, 6] = _px.Rows[i]["chapterflag"].ToString();
                    //xlSheet.Cells[i + row, 7] = _px.Rows[i]["core"].ToString();
                    //xlSheet.Cells[i + row, 8] = _px.Rows[i]["answerCode"].ToString();
                    //xlSheet.Cells[i + row, 9] = "'" + _px.Rows[i]["optionB"].ToString() + "/" + _px.Rows[i]["optionC"].ToString();
                    //xlSheet.Cells[i + row, 10] = _px.Rows[i]["optionD"].ToString();
                    //xlSheet.Cells[i + row, 11] = _px.Rows[i]["questiontype"].ToString();
                    //xlSheet.Cells[i + row, 12] = _px.Rows[i]["subjectType"].ToString() + "/" + _px.Rows[i]["explain"].ToString();
                  
                    
                }

            }
            catch
            { }
            try
            {
                Microsoft.Office.Interop.Excel.Range r1 = xlSheet.Range[xlSheet.Cells[3, 1], xlSheet.Cells[b + row + 4, 6]];
                r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = 7;
                r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = 7;
                r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = 7;
                r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = 7;
                r1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = 7;
            }
            catch
            { }



        }
        /// <summary>
        /// 人员类
        /// </summary>
        public class Choice
        {
            public string key; //A
            public string content; // 
            //public string C; // 
            //public string D; // 

        }


        private DataTable ToDataTable<T>(List<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                Type t = GetCoreType(prop.PropertyType);
                tb.Columns.Add(prop.Name, t);
            }

            foreach (T item in items)
            {
                var values = new object[props.Length];

                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }

                tb.Rows.Add(values);
            }

            return tb;
        }


        /// <summary>
        /// Determine of specified type is nullable
        /// </summary>
        public static bool IsNullable(Type t)
        {
            return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
        }

        /// <summary>
        /// Return underlying type if type is Nullable otherwise return the type
        /// </summary>
        public static Type GetCoreType(Type t)
        {
            if (t != null && IsNullable(t))
            {
                if (!t.IsValueType)
                {
                    return t;
                }
                else
                {
                    return Nullable.GetUnderlyingType(t);
                }
            }
            else
            {
                return t;
            }
        }
        private void duWord()
        {
            try
            {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = null;
                object unknow = Type.Missing;
                app.Visible = true;
                string str = @"E:\样例.doc";
                object file = str;
                doc = app.Documents.Open(ref file,
                    ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow);
                //得到标题
                int size1 = 100000;
                for (int i = 1; i < size1; i++)
                {
                    string temp = doc.Paragraphs[i].Range.Text.Trim();
                    string title11 = temp;
                    Console.Write(title11);
                    if (!temp.Equals(""))
                    {
                        string tem1p3 = temp.Substring(0, 1);
                        if (StrIsInt(temp.Substring(0, 1)))
                        {
                            String title = temp.Substring(2, temp.Length - 2); ///去掉两个字符串
                            String A = String.Empty;
                            String B = String.Empty;
                            String C = String.Empty;
                            String D = String.Empty;
                            String Answer = String.Empty;
                            String Answercode = String.Empty;//答案编码
                            String core = String.Empty;//相关词汇
                            String codeyi = String.Empty;//参考译文
                            String coreanasyl = String.Empty;//题目分析
                            String explain = String.Empty;//考点解析

                            i++;
                            temp = doc.Paragraphs[i].Range.Text.Trim();

                            if (temp.Substring(0, 2).Equals("A)"))
                            {
                                string[] sArray1 = temp.Split(new string[] { "A)", "B)", "C)", "D)" }, StringSplitOptions.RemoveEmptyEntries);
                                A = sArray1[0];
                                B = sArray1[1];
                                C = sArray1[2];
                                D = sArray1[3];
                                i++;
                                temp = doc.Paragraphs[i].Range.Text.Trim();
                                string tempcore1 = doc.Paragraphs[i++].Range.Text.Trim();
                                string tempcore2 = doc.Paragraphs[i++].Range.Text.Trim();
                                string tempcore3 = doc.Paragraphs[i++].Range.Text.Trim();
                                string tempcore4 = doc.Paragraphs[i++].Range.Text.Trim();
                                string tempcore8 = doc.Paragraphs[i++].Range.Text.Trim();
                                string tempcore5 = doc.Paragraphs[i++].Range.Text.Trim();//参考译文
                                string tempcore6 = doc.Paragraphs[i++].Range.Text.Trim();//题目分析
                                string tempcore7 = doc.Paragraphs[i].Range.Text.Trim();//考点解析
                                if (!StrIsInt(tempcore1.Trim().Substring(0, 1)) && !StrIsInt(tempcore2.Trim().Substring(0, 1)) && !StrIsInt(tempcore3.Trim().Substring(0, 1)) && !StrIsInt(tempcore4.Trim().Substring(0, 1)))
                                {
                                    //如果不等于数字则将数据付给core
                                    core = tempcore1 + tempcore2 + tempcore3 + tempcore4;
                                    if (!tempcore5.Substring(0, 3).Equals("【参考"))
                                    {
                                        MessageBox.Show("【参考" + tempcore5);
                                    }
                                    if (!tempcore6.Substring(0, 3).Equals("【题目"))
                                    {
                                        MessageBox.Show("【题目" + tempcore6);
                                    }
                                    if (!tempcore7.Substring(0, 3).Equals("【考点"))
                                    {
                                        MessageBox.Show("【考点" + tempcore7);
                                    }
                                    EnglishWordTiKu gameObject = new EnglishWordTiKu("EnglishWordTiKu");

                                    gameObject.titleSubject = title;

                                    gameObject.answer = "【答案】" + tempcore8;
                                    if (tempcore8.Equals("A"))
                                    {
                                        Answercode = "1";
                                    }
                                    else if (tempcore8.Equals("B"))
                                    {
                                        Answercode = "2";
                                    }
                                    else if (tempcore8.Equals("C"))
                                    {
                                        Answercode = "3";
                                    }
                                    else if (tempcore8.Equals("D"))
                                    {
                                        Answercode = "4";
                                    }

                                    gameObject.core = core;//相关词汇
                                    gameObject.answerCode = Answercode;//答案编码
                                    gameObject.codeyi = tempcore5;//参考译文
                                    gameObject.coreanasyl = tempcore6;//题目分析
                                    gameObject.explain = tempcore7;//考点解析
                                    //   gameObject.chapterflag = int.Parse(dt.Rows[i][3].ToString());
                                    gameObject.optionA = A;
                                    gameObject.optionB = B;
                                    gameObject.optionC = C;
                                    gameObject.optionD = D;
                                    //gameObject.explain = dt.Rows[i][2].ToString();
                                    gameObject.chapterflag = 1;
                                    gameObject.questiontype = 1;
                                    gameObject.subjectType = "0";

                                    var future2 = Bmob.CreateTaskAsync(gameObject);


                                }
                                else
                                {
                                    MessageBox.Show("选项中有整数" + i + "///" + temp);
                                }
                            }
                            else if (StrIsInt(temp.Substring(0, 1)))
                            {

                                MessageBox.Show(temp);

                            }
                            else
                            {
                                temp = doc.Paragraphs[i].Range.Text.Trim();

                            }
                        }

                    }
                    else
                    {

                        MessageBox.Show("数据上传结束" + i + "//" +doc.Paragraphs[i-2].Range.Text.Trim());

                    }


                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }

        private void duWord2()
        {
            try
            {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = null;
                object unknow = Type.Missing;
                app.Visible = true;
                string str = @"E:\1200题.doc";
                object file = str;
                doc = app.Documents.Open(ref file,
                    ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow);
                //得到标题
                int size1 = 100000;
                for (int i = 1; i < size1; i++)
                {
                    string temp = doc.Paragraphs[i].Range.Text.Trim();
                    string title11 = temp;
                    Console.Write(title11);
                    if (!temp.Equals(""))
                    {
                        string tem1p3 = temp.Substring(0, 1);
                        if (StrIsInt(temp.Substring(0, 1)))
                        {
                            String title = temp.Substring(2, temp.Length - 2); ///去掉两个字符串
                            String A = String.Empty;
                            String B = String.Empty;
                            String C = String.Empty;
                            String D = String.Empty;
                            String Answer = String.Empty;
                            String Answercode = String.Empty;//答案编码
                            //String core = String.Empty;//相关词汇
                            // String codeyi = String.Empty;//参考译文
                            // String coreanasyl = String.Empty;//题目分析
                            //String explain = String.Empty;//考点解析

                            i++;
                            temp = doc.Paragraphs[i].Range.Text.Trim();

                            if (temp.Substring(0, 2).Equals("A)"))
                            {
                                string[] sArray1 = temp.Split(new string[] { "A)", "B)", "C)", "D)" }, StringSplitOptions.RemoveEmptyEntries);
                                A = sArray1[0];
                                B = sArray1[1];
                                C = sArray1[2];
                                D = sArray1[3];
                                i++;
                                temp = doc.Paragraphs[i].Range.Text.Trim();
                                //string tempcore1 = doc.Paragraphs[i++].Range.Text.Trim();
                                //string tempcore2 = doc.Paragraphs[i++].Range.Text.Trim();
                                //string tempcore3 = doc.Paragraphs[i++].Range.Text.Trim();
                                //string tempcore4 = doc.Paragraphs[i++].Range.Text.Trim();
                                //string tempcore8 = doc.Paragraphs[i++].Range.Text.Trim();
                                //string tempcore5 = doc.Paragraphs[i++].Range.Text.Trim();//参考译文
                                //string tempcore6 = doc.Paragraphs[i++].Range.Text.Trim();//题目分析
                                //string tempcore7 = doc.Paragraphs[i].Range.Text.Trim();//考点解析
                                //if (!StrIsInt(tempcore1.Trim().Substring(0, 1)) && !StrIsInt(tempcore2.Trim().Substring(0, 1)) && !StrIsInt(tempcore3.Trim().Substring(0, 1)) && !StrIsInt(tempcore4.Trim().Substring(0, 1)))
                                if (!StrIsInt(temp.Trim().Substring(0, 1)))
                                {

                                    //如果不等于数字则将数据付给core
                                    //core = tempcore1 + tempcore2 + tempcore3 + tempcore4;
                                    //if (!tempcore5.Substring(0, 3).Equals("【参考"))
                                    //{
                                    //    MessageBox.Show("【参考" + tempcore5);
                                    //}
                                    //if (!tempcore6.Substring(0, 3).Equals("【题目"))
                                    //{
                                    //    MessageBox.Show("【题目" + tempcore6);
                                    //}
                                    //if (!tempcore7.Substring(0, 3).Equals("【考点"))
                                    //{
                                    //    MessageBox.Show("【考点" + tempcore7);
                                    //}
                                    EnglishWordTiKu gameObject = new EnglishWordTiKu("EnglishWordTiKu");

                                    gameObject.titleSubject = title;

                                    gameObject.answer = "【答案】" + temp;
                                    if (temp.Equals("A"))
                                    {
                                        Answercode = "1";
                                    }
                                    else if (temp.Equals("B"))
                                    {
                                        Answercode = "2";
                                    }
                                    else if (temp.Equals("C"))
                                    {
                                        Answercode = "3";
                                    }
                                    else if (temp.Equals("D"))
                                    {
                                        Answercode = "4";
                                    }

                                    // gameObject.core = core;//相关词汇
                                    gameObject.answerCode = Answercode;//答案编码
                                    // gameObject.codeyi = tempcore5;//参考译文
                                    // gameObject.coreanasyl = tempcore6;//题目分析
                                    //gameObject.explain = tempcore7;//考点解析
                                    //   gameObject.chapterflag = int.Parse(dt.Rows[i][3].ToString());
                                    gameObject.optionA = A;
                                    gameObject.optionB = B;
                                    gameObject.optionC = C;
                                    gameObject.optionD = D;
                                    //gameObject.explain = dt.Rows[i][2].ToString();
                                    gameObject.chapterflag = 2;
                                    gameObject.questiontype = 1;
                                    gameObject.subjectType = "0";

                                    var future2 = Bmob.CreateTaskAsync(gameObject);


                                }
                                else
                                {
                                    MessageBox.Show("选项中有整数" + i + "///" + temp);
                                }
                            }


                        }
                        else {
                            MessageBox.Show("选项中有空" + i + "///" + doc.Paragraphs[i - 1].Range.Text.Trim());
                        }

                    }
                    else
                    {

                        MessageBox.Show("数据上传结束" + i + "//" + doc.Paragraphs[i - 2].Range.Text.Trim());

                    }


                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }
        /// <summary>
        /// 判断字符串是否为数字
        /// </summary>
        /// <param name="Str"></param>
        /// <returns></returns>
        public static bool StrIsInt(string Str)
        {
            bool flag = true;
            if (Str != "")
            {
                for (int i = 0; i < Str.Length; i++)
                {
                    if (!Char.IsNumber(Str, i))
                    {
                        flag = false;
                        break;
                    }
                }
            }
            else
            {
                flag = false;
            }
            return flag;
        }
    }
}
