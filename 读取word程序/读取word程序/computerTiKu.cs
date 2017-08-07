using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using cn.bmob.io;
using cn.bmob.api;
using cn.bmob.json;
using cn.bmob.tools;
namespace BanMa.Model
{
    class computerTiKu : BmobTable
    {

          private String fTable;

          public BmobInt indexID { get; set; } //题号
          public String titleSubject { get; set; } //标题
          public String answer { get; set; } //答案
          public String optionA { get; set; } //选项
          public BmobInt chapterflag { get; set; }//属于哪章节
          public String core { get; set; }//考点
          public String answerCode { get; set; } //答案编码
          public String optionB { get; set; } 
          public String optionC { get; set; }  
          public String optionD { get; set; }
          public BmobInt questiontype { get; set; }//试题类型
          public String subjectType { get; set; }//题目类型
          public String explain { get; set; }//解析
         
        //构造函数
        public computerTiKu() { }

        //构造函数
        public computerTiKu(String tableName)
        {
            this.fTable = tableName;
        }

        public override string table
        {
            get
            {
                if (fTable != null)
                {
                    return fTable;
                }
                return base.table;
            }
        }

        //读字段信息
        public override void readFields(BmobInput input)
        {
            base.readFields(input);
            this.indexID = input.getInt("indexID");
            this.titleSubject = input.getString("titleSubject");
            this.answer = input.getString("answer");
            this.optionA = input.getString("optionA");
            this.chapterflag = input.getInt("chapterflag");
            this.core = input.getString("core");
            this.answerCode = input.getString("answerCode");
            this.optionB = input.getString("optionB");
            this.optionC = input.getString("optionC");
            this.optionD = input.getString("optionD");
            this.questiontype = input.getInt("questiontype");
            this.explain = input.getString("explain");

            this.subjectType = input.getString("subjectType");
            
        }

        //写字段信息
        public override void write(BmobOutput output, bool all)
        {
            base.write(output, all);
 
            output.Put("indexID", this.indexID);
            output.Put("titleSubject", this.titleSubject);
            output.Put("answer", this.answer);
            output.Put("optionA", this.optionA);
            output.Put("chapterflag", this.chapterflag);
            output.Put("core", this.core);
            output.Put("answerCode", this.answerCode);
            output.Put("optionB", this.optionB);
            output.Put("optionC", this.optionC);
            output.Put("optionD", this.optionD);
            output.Put("questiontype", this.questiontype);
            output.Put("explain", this.explain);
            output.Put("subjectType", this.subjectType);
            
        }
    }
}
