using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using mshtml;

namespace IECall
{
    public partial class Form1 : Form
    {

        #region 全局变量

        // 访问地址
        string url = "http://s.gddec.net/icourse/index/doStudyIndex.do";

        // 页面文档
        IHTMLDocument2 doc = null;
        
        // 循环点击列表
        string[] todoList = new string[] {
                "课程学习",
                "通知公告",
                "小纸条",
                "互助交流",
                "学习笔记",
                "我的班级",
                "成绩统计",
                "课程资料",
                "助学评价"
            };
        // 当前点击项
        static int listPn = 0;

        #endregion



        public Form1()
        {
            InitializeComponent();
            btnStop.Enabled = false;
        }



        //根据网站地址获取页面文档
        private bool GetIHTMLDocByUrl()
        {
            doc = ieHelper.GetIHTMLDocument2ByUrl(url);

            if (doc == null)
                doc = ieHelper.GetIHTMLDocument2ByUrl(url + "###");

            return doc != null;
        }
        

        // 执行定时任务
        private void btnRun_Click(object sender, EventArgs e)
        {            
            int timeInterval = 0;

            if (!int.TryParse(textBox1.Text, out timeInterval))
            {
                MessageBox.Show("输入的秒数格式不对");
                return;
            }

            if (!GetIHTMLDocByUrl())
            {
                MessageBox.Show("未找到学习页面，无法启动定时任务");
                return;
            }

            timer1.Interval = timeInterval * 1000;
            timer1.Start();

            btnRun.Enabled = false;
            btnStop.Enabled = true;
        }


        // 取消定时任务
        private void btnStop_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            btnRun.Enabled = true;
            btnStop.Enabled = false;
        }



        // 定时任务事件内容
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!GetIHTMLDocByUrl())
            {
                MessageBox.Show("任务中断，未找到学习页面，请确定已打开");
                this.btnStop_Click(sender, e);
                return;
            }

            ieHelper.PressBtnInHTMLDocument(doc, todoList[listPn++], "a");

            if (listPn >= todoList.Length)
                listPn = 0;
        }

        
    }
}
