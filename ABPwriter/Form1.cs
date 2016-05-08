using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;


namespace ABPwriter
{
    public partial class Form1 : Form
    {

        List<string> abpList;
        DataSet ds = new DataSet("ABP");
        public Form1()
        {
            InitializeComponent();
            ds.Tables.Add();
            ds.Tables[0].Columns.Add("Patient ID");
            ds.Tables[0].Columns.Add("昼夜节律");
            ds.Tables[0].Columns.Add("24 h");
            ds.Tables[0].Columns.Add("24 h ");
            ds.Tables[0].Columns.Add("Day");
            ds.Tables[0].Columns.Add("Day ");
            ds.Tables[0].Columns.Add("Night");
            ds.Tables[0].Columns.Add("Night ");
            ds.Tables[0].Rows.Add(new string[] { "Patient ID", "昼夜节律", "24 h", "", "Day", "", "Night" ,""});
            ds.Tables[0].Rows.Add(new string[] { " ", " ", "SBP", "DBP", "SBP", "DBP", "SBP", "DBP" });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            abpList = new List<string>();
            String[] files = Directory.GetFiles(@textBox1.Text);
            progressBar1.Maximum = files.Length+1;
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            for (int i = 0; i < files.Length; i++)
            {
                //删除文档中的名字
                List<string> lines = new List<string>(File.ReadAllLines(files[i],Encoding.Default));
                //Regex r= new Regex(@"^[b]\d+$");
                //Match m = r.Match(lines[1]); 
                //if(m.Success)
                lines[0] = "****";
                //File.WriteAllLines(files[i], lines.ToArray(), Encoding.Default);
                //生成报告
                String name = files[i].Substring(files[i].LastIndexOf('.')-6, 6).ToLower();
                FileStream fs = new FileStream(files[i], FileMode.Open);
                StreamReader m_streamReader = new StreamReader(fs,Encoding.Default);
                String tempvalue;
                int CYCLE = 0;
                string[] data = new string[8];
                data[0] =name;
                label3.Text = "handle:" + name + "...";
                label3.Refresh();
                while (CYCLE<25)
                {

                    CYCLE++;
                    tempvalue=m_streamReader.ReadLine();
                    if (tempvalue.Contains("均正常"))
                        continue;
                    if (tempvalue.Contains("昼夜节律"))
                    {
                        if (tempvalue.Contains("消失"))
                        {
                            data[1]="0";
                        }
                        else 
                            data[1]="1";
                    }
                    else if (tempvalue.Contains("24小时平均"))
                    {
                        string[] acc=new string[5];
                        int bigornot=0;
                        if (tempvalue.Contains("("))
                            bigornot += 1;
                        else if (tempvalue.Contains("（"))
                            bigornot += 10;
                        if (tempvalue.Contains(")"))
                            bigornot += 100;
                        else if (tempvalue.Contains("）"))
                            bigornot += 1000;
                        switch (bigornot){
                            case 1010:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('（') + 1, tempvalue.LastIndexOf('）') - tempvalue.LastIndexOf('（') - 5).Split('/');
                                break;
                            case 101:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('(') + 1, tempvalue.LastIndexOf(')') - tempvalue.LastIndexOf('(') - 5).Split('/');
                                break;
                            case 110:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('（') + 1, tempvalue.LastIndexOf(')') - tempvalue.LastIndexOf('（') - 5).Split('/');
                                break;
                            case 1001:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('(') + 1, tempvalue.LastIndexOf('）') - tempvalue.LastIndexOf('(') - 5).Split('/');
                                break;
                        }
                        for (int ib = 0; ib < acc.Length; ib++)
                        {
                            if (acc[ib].Contains("m"))
                                acc[ib] = acc[ib].Substring(0, acc[ib].Length - 1);
                        }
                        if (acc.Length == 2)
                        {
                            data[2] = acc[0];
                            data[3] = acc[1]; 
                        }
                        else if (acc.Length == 1)
                        {
                            if (tempvalue.Contains("收缩压"))
                                data[2] = acc[0];
                            else if(tempvalue.Contains("舒张压"))
                                data[3]=acc[0];
                        }
                    }
                    else if (tempvalue.Contains("白天平均"))
                    {
                        string[] acc = new string[5];
                        int bigornot = 0;
                        if (tempvalue.Contains("("))
                            bigornot += 1;
                        else if (tempvalue.Contains("（"))
                            bigornot += 10;
                        if (tempvalue.Contains(")"))
                            bigornot += 100;
                        else if (tempvalue.Contains("）"))
                            bigornot += 1000;
                        switch (bigornot)
                        {
                            case 1010:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('（') + 1, tempvalue.LastIndexOf('）') - tempvalue.LastIndexOf('（') - 5).Split('/');
                                break;
                            case 101:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('(') + 1, tempvalue.LastIndexOf(')') - tempvalue.LastIndexOf('(') - 5).Split('/');
                                break;
                            case 110:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('（') + 1, tempvalue.LastIndexOf(')') - tempvalue.LastIndexOf('（') - 5).Split('/');
                                break;
                            case 1001:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('(') + 1, tempvalue.LastIndexOf('）') - tempvalue.LastIndexOf('(') - 5).Split('/');
                                break;
                        }
                        for (int ib = 0; ib < acc.Length; ib++)
                        {
                            if (acc[ib].Contains("m"))
                                acc[ib] = acc[ib].Substring(0, acc[ib].Length - 1);
                        }
                        if (acc.Length == 2)
                        {
                            data[4] = acc[0];
                            data[5] = acc[1];
                        }
                        else if (acc.Length == 1)
                        {
                            if (tempvalue.Contains("收缩压"))
                                data[4] = acc[0];
                            else if (tempvalue.Contains("舒张压"))
                                data[5] = acc[0];
                        }
                    }
                    else if (tempvalue.Contains("夜间平均"))
                    {
                        string[] acc = new string[5];
                        int bigornot = 0;
                        if (tempvalue.Contains("("))
                            bigornot += 1;
                        else if (tempvalue.Contains("（"))
                            bigornot += 10;
                        if (tempvalue.Contains(")"))
                            bigornot += 100;
                        else if (tempvalue.Contains("）"))
                            bigornot += 1000;
                        switch (bigornot)
                        {
                            case 1010:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('（') + 1, tempvalue.LastIndexOf('）') - tempvalue.LastIndexOf('（') - 5).Split('/');
                                break;
                            case 101:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('(') + 1, tempvalue.LastIndexOf(')') - tempvalue.LastIndexOf('(') - 5).Split('/');
                                break;
                            case 110:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('（') + 1, tempvalue.LastIndexOf(')') - tempvalue.LastIndexOf('（') - 5).Split('/');
                                break;
                            case 1001:
                                acc = tempvalue.Substring(tempvalue.LastIndexOf('(') + 1, tempvalue.LastIndexOf('）') - tempvalue.LastIndexOf('(') - 5).Split('/');
                                break;
                        }
                        for (int ib = 0; ib < acc.Length; ib++)
                        {
                            if (acc[ib].Contains("m"))
                                acc[ib] = acc[ib].Substring(0, acc[ib].Length - 1);
                        }


                        if (acc.Length == 2)
                        {
                            data[6] = acc[0];
                            data[7] = acc[1];
                        }
                        else if (acc.Length == 1)
                        {
                            if (tempvalue.Contains("收缩压"))
                                data[6] = acc[0];
                            else if (tempvalue.Contains("舒张压"))
                                data[7] = acc[0];
                        }
                    }
                    //Console.WriteLine(tempvalue); 
                }
                //向ds中写数据
                ds.Tables[0].Rows.Add(data);
                progressBar1.PerformStep();
                progressBar1.Refresh();
            }

            label3.Text = "Writing Excel...";
            gSendGridInfoToExcel(ds, @textBox2.Text);
            progressBar1.PerformStep();
            progressBar1.Refresh();
            label3.Text="Complete";
            progressBar1.Value = 0;
        }

        public static void gSendGridInfoToExcel(DataSet ds, string excelpath)
        {
            // set culture to US
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            if (ds.Tables.Count <= 0)
            {

                return;
            }
            int count = ds.Tables[0].Rows.Count;//获取数据表中DataRow行总数
            int column = ds.Tables[0].Columns.Count;//获取数据表中列总数
            Microsoft.Office.Interop.Excel.ApplicationClass excelapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            //Make Excel Application Visible
            excelapp.Visible = false;
            //写入特定文件
            //Excel.Workbook wb = excelapp.Workbooks.Open(string filename,
            //    Type.Missing,Type.Missing,Type.Missing,Type.Missing,
            //    Type.Missing,Type.Missing,Type.Missing,Type.Missing,
            //    Type.Missing,Type.Missing,Type.Missing,Type.Missing,
            //    Type.Missing,Type.Missing);
            Microsoft.Office.Interop.Excel.Workbook wb = excelapp.Application.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Worksheet sheets = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[1];
            sheets.Name = excelpath.Substring(excelpath.LastIndexOf(@"\") + 1); ;
            for (int x = 1; x <= count; x++)
            {
                for (int y = 1; y <= column; y++)
                {
                    //excelapp.Cells[x, y] = this.archiverdbDataSet.channel.Rows[x- 1].ItemArray[y - 1];
                    excelapp.Cells[x, y] = ds.Tables[0].Rows[x - 1].ItemArray[y - 1];
                }
            }


            try
            {
                wb.SaveAs(excelpath + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel7/*Type.Missing*/, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Saved = true;
                excelapp.UserControl = false;
            }
            catch (Exception e)
            {

                MessageBox.Show(e.ToString());
            }
            finally
            {
                excelapp.Quit();
                excelapp = null;
                ds = null;
                GC.Collect();//垃圾回收 
            }
            // restore the environment
            System.Threading.Thread.CurrentThread.CurrentCulture = CurrentCI;
            return;
        }
    }
}
