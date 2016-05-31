using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.IO;
using System.Xml;

namespace DocPrinter
{
    public partial class Form1 : Form
    {
        int prs_count = 0;
        int prs_on = 0;
        BackgroundWorker bgw = new BackgroundWorker();

        public Form1()
        {
            InitializeComponent();
            bgw.WorkerReportsProgress = true;
            bgw.WorkerSupportsCancellation = true;
            bgw.DoWork += bgw_DoWork;
            bgw.RunWorkerCompleted += bgw_RunWorkerCompleted;
            bgw.ProgressChanged += bgw_ProgressChanged;
        }
        void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            label1.Text = prs_on + "/" + prs_count;
            progressBar1.Value = e.ProgressPercentage;
        }

        //delegate void InitCallback(int num);
        //private void syncProcessBar(int num)
        //{
        //    prs_count = num;
        //    prs_on = 0;
        //    if (this.progressBar1.InvokeRequired || this.label1.InvokeRequired)
        //    {
        //        InitCallback d = new InitCallback(syncProcessBar);
        //        this.Invoke(d, new object[] { num });
        //    }
        //    else
        //    {
        //        progressBar1.Value = 0;
        //        label1.Text = prs_on + "/" + prs_count;
        //    }
        //}

        //delegate void SetValueCallback();
        //private void syncProcessBar()
        //{
        //    if (this.progressBar1.InvokeRequired || this.label1.InvokeRequired)
        //    {
        //        SetValueCallback d = new SetValueCallback(syncProcessBar);
        //        this.Invoke(d, new object[] {});
        //    }
        //    else
        //    {
        //        ++prs_on;
        //        progressBar1.Value = prs_on * 100 / prs_count;
        //        label1.Text = prs_on + "/" + prs_count;
        //    }
            
        //}

        private void csv_go(string name)
        {
            StreamReader sr = new StreamReader("input\\" + name + ".csv", Encoding.Default);
            XmlHelper xh = new XmlHelper(name);
            List<string> head = new List<string>(sr.ReadLine().Split(','));
            while (!sr.EndOfStream)
            {
                if (bgw.CancellationPending)
                {
                    sr.Close();
                    return;
                }
                string[] one_line = sr.ReadLine().Split(',');
                string sname = one_line[head.FindIndex(s => s == xh.Bookmark["name"])];
                WordHelper onPrintingWH = new WordHelper();
                string currentDir = System.IO.Directory.GetCurrentDirectory();
                onPrintingWH.OpenAndActive(currentDir + "\\tpls\\" + name + ".dotx", false, true);
                foreach (KeyValuePair<string, string> bmk in xh.Bookmark)
                {
                    if (bgw.CancellationPending)
                    {
                        onPrintingWH.Close();
                        sr.Close();
                        return;
                    }
                    onPrintingWH.GoToBookMark(bmk.Key);
                    onPrintingWH.InsertText(one_line[head.FindIndex(s => s == bmk.Value)]);
                }
                if (!Directory.Exists("output\\" + sname))
                    Directory.CreateDirectory("output\\" + sname);
                onPrintingWH.SaveAs(currentDir + "\\output\\" + sname + "\\"+ name +".doc");
                onPrintingWH.Close();
                //syncProcessBar();
                bgw.ReportProgress(++prs_on * 100 / prs_count);
            }
            sr.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "开始")
            {
                button1.Text = "取消";
                label2.Text = "正在计算中，请稍候……";
                bgw.RunWorkerAsync();
            }
            else
            {
                bgw.CancelAsync();
            }
        }

        void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
                label2.Text = "用户取消了操作";
            else
                label2.Text = "输出完毕";
            button1.Text = "开始";
        }

        void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            XmlDocument index = new XmlDocument();
            index.Load("tpls\\index.xml");
            XmlNode root = index.SelectSingleNode(index.DocumentElement.Name);
            int count = 0;
            foreach (XmlNode tpl in root.SelectNodes("tpl"))
            {
                StreamReader sr = new StreamReader("input\\" + tpl.InnerText + ".csv");
                while (!sr.EndOfStream)
                {
                    sr.ReadLine();
                    ++count;
                }
                --count;
            }
            //syncProcessBar(count);
            prs_count = count;
            prs_on = 0;
            bgw.ReportProgress(0);
            foreach (XmlNode tpl in root.SelectNodes("tpl"))
            {
                csv_go(tpl.InnerText);
                if (bgw.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
            }
            
        }

    }
}

//生成WORD程序对象和WORD文档对象
//Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
//Microsoft.Office.Interop.Word.Document doc = new Document();
//object oMissing = System.Reflection.Missing.Value;
////打开模板文档，并指定doc的文档类型

//object objTemplate = System.Windows.Forms.Application.StartupPath + @"\威海一中实验部学生基本信息登记表.dotx";
////路径一定要正确
//// HttpContext.Current.Server.MapPath(@"f:\tz103.doc");
//object objDocType = WdDocumentType.wdTypeDocument;
//object objfalse = false;
//object objtrue = true;
//doc = (Document)appWord.Documents.Add(ref objTemplate, ref objfalse, ref objDocType, ref objtrue);
////获取模板中所有的书签
//Bookmarks odf = doc.Bookmarks;

//string[] testTableremarks = { "name" };
//string[] testTablevalues = { "Word标题"};

////循环所有的书签，并给书签赋值
//for (int oIndex = 0; oIndex < testTableremarks.Length; oIndex++)
//{
//    object obDD_Name = "";
//    obDD_Name = testTableremarks[oIndex];
//    //doc.Bookmarks.get_Item(ref obDD_Name).Range.Text = p_TestReportTable.Rows[0][testTablevalues[oIndex]].ToString();//此处Range也是WORD中很重要的一个对象，就是当前操作参数所在的区域
//    doc.Bookmarks.get_Item(ref obDD_Name).Range.Text = testTablevalues[oIndex];
//}

////第四步 生成word，将当前的文档对象另存为指定的路径，然后关闭doc对象。关闭应用程序
//object filename = "F:/" +"1"+".doc";//HttpContext.Current.Server.MapPath("f:\\") + "Testing_" + DateTime.Now.ToShortDateString() + ".doc";
//object miss = System.Reflection.Missing.Value;
//doc.SaveAs(ref filename, ref miss, ref miss, ref miss, ref miss, ref miss,
//    ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
//object missingValue = Type.Missing;
//object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;
//doc.Close(ref doNotSaveChanges, ref missingValue, ref missingValue);
//appWord.Application.Quit(ref miss, ref miss, ref miss);
//doc = null;
//appWord = null;
//// MessageBox.Show("生成成功！");
//System.Diagnostics.Process.Start(filename.ToString());//打开文档
