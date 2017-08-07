
using cn.bmob.api;
using cn.bmob.json;
using cn.bmob.tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BanMa
{
    public partial class BmobBasePCForm : Form
    {       //创建Bmob实例
        private BmobWindows bmob;
        private ToolTip _toolTip;
        public  BmobBasePCForm()
            : base()
        {
            bmob = new BmobWindows();
            _toolTip = new ToolTip();
            //初始化ApplicationId，这个ApplicationId需要更改为你自己的ApplicationId（ http://www.bmob.cn 上注册登录之后，创建应用可获取到ApplicationId）
            Bmob.initialize("f8e94fde865de2270860606c30e4cdaa", "a586925dc4857f8f598ca6edbee17894");

            //注册调试工具
            BmobDebug.Register(msg => { Debug.WriteLine(msg); });
        }

        public ToolTip ToolTip
        {
            get { return _toolTip; }
        }
        public BmobWindows Bmob
        {
            get { return bmob; }
        }

        //对返回结果进行显示处理
        public void FinishedCallback<T>(T data, TextBox text)
        {
            text.Text = JsonAdapter.JSON.ToDebugJsonString(data);
        }

    }
}
