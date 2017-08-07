using BanMa.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 读取word程序
{
    public partial class Form1 : BanMa.BmobBasePCForm
    {
        public Form1()
        {
            
            InitializeComponent();

            duWord2();
            ////Word.ApplicationClass doc = new Microsoft.Office.Interop.Word.ApplicationClass();  
            //try
            //{
            //    Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            //    Microsoft.Office.Interop.Word.Document doc = null;
            //    object unknow = Type.Missing;
            //    app.Visible = true;
            //    string str = @"E:\样例.doc";
            //    object file = str;
            //    doc = app.Documents.Open(ref file,
            //        ref unknow, ref unknow, ref unknow, ref unknow,
            //        ref unknow, ref unknow, ref unknow, ref unknow,
            //        ref unknow, ref unknow, ref unknow, ref unknow,
            //        ref unknow, ref unknow, ref unknow);
            //    //得到标题
            //    int size1 = 1000;
            //    for(int i=1;i<size1;i++){
            //        string temp = doc.Paragraphs[i].Range.Text.Trim();
            //        if (!temp.Equals(""))
            //        {
            //            string tem1p3 = temp.Substring(0, 1);
            //            if (StrIsInt(temp.Substring(0, 1)))
            //            {
            //                String title = temp;
            //                String A = String.Empty;
            //                String B = String.Empty;
            //                String C = String.Empty;
            //                String D = String.Empty;
            //                String core = String.Empty;
            //                i++;
            //                temp = doc.Paragraphs[i].Range.Text.Trim();

            //                if (temp.Substring(0, 2).Equals("A)"))
            //                {
            //                    string[] sArray1 = temp.Split(new string[] { "A)", "B)", "C)", "D)" }, StringSplitOptions.RemoveEmptyEntries);
            //                    A = sArray1[0];
            //                    B = sArray1[1];
            //                    C = sArray1[2];
            //                    D = sArray1[3];
            //                    i++;
            //                    temp = doc.Paragraphs[i].Range.Text.Trim();
            //                    string tempcore1 = doc.Paragraphs[i++].Range.Text.Trim();
            //                    string tempcore2 = doc.Paragraphs[i++].Range.Text.Trim();
            //                    string tempcore3 = doc.Paragraphs[i++].Range.Text.Trim();
            //                    string tempcore4 = doc.Paragraphs[i].Range.Text.Trim();
            //                    if (!StrIsInt(tempcore1.Trim().Substring(0, 1)) && !StrIsInt(tempcore2.Trim().Substring(0, 1)) && !StrIsInt(tempcore3.Trim().Substring(0, 1)) && !StrIsInt(tempcore4.Trim().Substring(0, 1)))
            //                    {
            //                        //如果不等于数字则将数据付给core
            //                        core = tempcore1 + tempcore2 + tempcore3 + tempcore4;

            //                          computerTiKu gameObject = new computerTiKu("computerTiKu");




            //                          gameObject.titleSubject = title;
            //                          gameObject.answer = "【答案】";

            //                          gameObject.core = core;
            // //   gameObject.chapterflag = int.Parse(dt.Rows[i][3].ToString());
            //                          gameObject.optionA = A;
            //                          gameObject.optionB = B;
            //                         gameObject.optionC = C;
            //                         gameObject.optionD = D;
            //                        //gameObject.explain = dt.Rows[i][2].ToString();
               
            //                        gameObject.questiontype = 0;
            //                        gameObject.subjectType ="0";
               
               

            //    //保存数据

            //                        //保存数据
            //                        var future2 = cn.bmob.api.Bmob.CreateTaskAsync(gameObject);

            //                    }
            //                    else
            //                    {
            //                        MessageBox.Show("选项中有整数" + i + "///"+temp);
            //                    }
            //                }
            //                else if (StrIsInt(temp.Substring(0, 1)))
            //                {

            //                    MessageBox.Show(temp);

            //                }
            //                else
            //                {
            //                    temp = doc.Paragraphs[i].Range.Text.Trim();

            //                }
            //            }

            //        }
            //        else {
            //            MessageBox.Show("数据上传结束"+i);
                    
            //        }
                   
                     
            //    }
              
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}  


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
