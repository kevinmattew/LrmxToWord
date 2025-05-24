using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace 干部任免审批表转换
{
    public partial class Form1 : Form
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private List<String>  files = new List<String>();   
        private Stopwatch stopwatch = new Stopwatch();
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = "";
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK) {
                files.Clear();
                while (dataGridView1.Rows.Count > 0) {
                    if (dataGridView1.Rows.Count > 0) {
                        dataGridView1.Rows.RemoveAt(0);
                    }
                
                }
                int num = 0;
                string selectedPath = folderBrowserDialog1.SelectedPath;
                listDirectory(selectedPath);
                foreach (string file in files) {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[num].Cells[0].Value = file;  
                    num ++; 
                }
                this.textBox1.Text = folderBrowserDialog1.SelectedPath;
                
            }

        }
        private void listDirectory(string path)
        {
            DirectoryInfo theFolder = new DirectoryInfo(@path);

            //遍历文件
            foreach (FileInfo NextFile in theFolder.GetFiles())
            {
                 if(NextFile.Name.EndsWith("lrmx"))
                {
                    files.Add(NextFile.FullName);
                    

                }
            }

            //遍历文件夹
            foreach (DirectoryInfo NextFolder in theFolder.GetDirectories())
            {
                listDirectory(NextFolder.FullName);
            }
        }
 


            private void Form1_Load(object sender, EventArgs e)
        {

        }


        private GanBu LrmxToClass(string filename) {
            GanBu ganBu = new GanBu();
            Type type = ganBu.GetType();
            try {

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(filename);
                XmlNodeList nl = xmlDoc.DocumentElement.ChildNodes;
                foreach (XmlNode node in nl)
                {
                    //Debug.WriteLine(node.Name);
                    if (node.Name != "JianLi" && node.Name != "JiaTingChengYuan")
                    {
                        FieldInfo fieldInfo = type.GetField(node.Name);
                        fieldInfo.SetValue(ganBu, node.InnerText);

                    }

                    if (node.Name == "JianLi")
                    {
                        List<JianLi> jianLiList = new List<JianLi>();
                        string jianli = node.InnerText;
                        int rowNums = 0;
                        string[] JianLiArrayTmp = jianli.Split('\n');
                        List<String> JianLiArray = new List<string>();
                        foreach (string s in JianLiArrayTmp)
                        {
                            if (s.Replace("\n", "").Replace("\r","").Trim() != "")
                            {
                                JianLiArray.Add(s.Replace("\r",""));
                            }

                        }
                        foreach (string line in JianLiArray)
                        {
                            if (line.Trim() != "")
                            {
                                if (!Tools.isStartWithNumber(line.Trim()))
                                {
                                    jianLiList[rowNums - 1].JingLi = jianLiList[rowNums - 1].JingLi + " " + line.Trim();
                                   
                                }
                                else
                                {
                                    JianLi jl = new JianLi();
                                    String[] partArray = Regex.Split(line, "  ");
                                    int tmpNum = 0;
                                    String jingli = "";
                                    foreach (string s in partArray)
                                    {
                                        if (tmpNum == 0)
                                        {
                                            //第一列分割的用来做时间拆分
                                            string[] a = Regex.Split(s, "--|——");
                                            if (a.Length == 1)
                                            {
                                                a = Regex.Split(s, "-|—");
                                            }
                                            jl.KaiShiNianYue = a[0].Trim().Replace(" ","");
                                            if (a.Length == 1)
                                            {
                                                jl.JieSuNianYue = "";
                                            }
                                            else
                                            {
                                                jl.JieSuNianYue += a[1].Trim().Replace(" ", "");

                                            }
                                            tmpNum++;
                                        }
                                        else
                                        {
                                            jingli += s.Trim();
                                        }

                                    }
                                    jl.JingLi = jingli;

                                    jianLiList.Add(jl);
                                    rowNums++;

                                }
                            }
                        }


                        ganBu.JianLi = jianLiList;
                    }

                    if (node.Name == "JiaTingChengYuan")
                    {
                        XmlNodeList node2list = node.ChildNodes;
                        List<JiaTingChengYuan> jtcyList = new List<JiaTingChengYuan>();
                        foreach (XmlNode node2 in node2list)
                        {
                            JiaTingChengYuan jtcy = new JiaTingChengYuan();
                            XmlNodeList node3List = node2.ChildNodes;
                            Type t2 = jtcy.GetType();
                            foreach (XmlNode node3 in node3List)
                            {
                                FieldInfo fi = t2.GetField(node3.Name);
                                fi.SetValue(jtcy, node3.InnerText.Trim());
                            }
                            jtcy.ChuShengRiQi = jtcy.ChuShengRiQi.Trim().Replace("\n", "").Replace("-", "");
                            if (jtcy.ChuShengRiQi.Length == 6 || jtcy.ChuShengRiQi.Length == 8)
                            {

                                jtcy.NianLing = Tools.JiSuanNianLing(jtcy.ChuShengRiQi, Tools.NullToEmpty(ganBu.JiSuanNianLingShiJian), Tools.NullToEmpty(ganBu.TianBiaoShiJian));
                            }

                            jtcyList.Add(jtcy);


                        }
                        ganBu.JiaTingChengYuan = jtcyList;

                    }

                }

                //20240718 增加身份证提取出生日期
                if (ganBu.ChuShengNianYue.Length == 6 || ganBu.ChuShengNianYue.Length == 8)
                {
                    ganBu.NianLing = Tools.JiSuanNianLing(ganBu.ChuShengNianYue, Tools.NullToEmpty(ganBu.JiSuanNianLingShiJian), Tools.NullToEmpty(ganBu.TianBiaoShiJian));
                }
                else
                {
                    if (!string.IsNullOrEmpty(ganBu.ShenFenZheng))
                    {
                        ganBu.ChuShengNianYue = Tools.GetDate(ganBu.ShenFenZheng);
                        ganBu.NianLing = Tools.JiSuanNianLing(ganBu.ChuShengNianYue, Tools.NullToEmpty(ganBu.JiSuanNianLingShiJian), Tools.NullToEmpty(ganBu.TianBiaoShiJian));
                    }

                }


                logger.Info(filename + "转换为类成功！");

                return ganBu;
            
            }catch(Exception ex)
            {
                logger.Error(ex.ToString());
                logger.Error(filename + "转换为类失败！");
                return ganBu;
            }

        }



        private void LrmxToDocx(String filename) {
             
            try
            {
                GanBu ganBu = LrmxToClass(filename);                 

                string[]  tmpArr = filename.Split('\\');
                string newFileName = tmpArr[tmpArr.Length - 1];
                string filePath = Path.GetDirectoryName(filename);
                String targetFilePath = ".\\docx\\"  + filePath.Substring(3, filePath.Length - 3) + "\\" +  newFileName.Substring(0,  newFileName.Length -5) + "-任免审批表.docx"; 
                if (File.Exists(targetFilePath))
                {
                    File.Delete(targetFilePath);
                }

                XWPFDocument docx;
                using (FileStream stream = File.OpenRead(".\\spbmb.docx"))
                {

                    docx = new XWPFDocument(stream);
                }
                foreach (XWPFParagraph ph in docx.Paragraphs)
                {
                    foreach (XWPFRun r in ph.Runs)
                    {
                        if (r.Text == "填表人：")
                        {
                            r.SetText("填表人：" + ganBu.TianBiaoRen);

                        }
                    }
                }


                if (!string.IsNullOrEmpty(ganBu.ZhaoPian.Replace("\r", "").Replace("\n", "")))
                {
                    XWPFTableCell imageCell = docx.Tables[1].Rows[0].GetTableCells()[6];
                    XWPFParagraph pg = imageCell.Paragraphs[0];
                    XWPFRun run = pg.CreateRun();
                    byte[] b = Convert.FromBase64String(ganBu.ZhaoPian);
                    InputStream sbs = new ByteArrayInputStream(b);
                    run.AddPicture(new ByteArrayInputStream(b), (int)NPOI.SS.UserModel.PictureType.JPEG, "", 363 * 3600, 490 * 3600);
                }


                Boolean ifFirst = true;
                int num = 0;
                foreach (XWPFTable table in docx.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        for (int j = 0; j < table.Rows[i].GetTableCells().Count; j++)
                        {
                            switch (table.Rows[i].GetTableCells()[j].GetText().Trim().Replace(" ", ""))
                            {
                                case "姓名":
                                    if (num == 1)
                                    {
                                        table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.XingMing);
                                    }
                                    break;
                                case "性别":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.XingBie);
                                    break;
                                case "出生年月(岁)":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.ChuShengNianYue.Substring(0, 4) + "." + ganBu.ChuShengNianYue.Substring(4, 2) + "（" + ganBu.NianLing + "岁）");
                                    break;
                                case "民族":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.MinZu);
                                    break;
                                case "籍贯":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.JiGuan);
                                    break;
                                case "出生地":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.ChuShengDi);
                                    break;
                                case "入党时间":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.RuDangShiJian);
                                    break;
                                case "参加工作时间":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.CanJiaGongZuoShiJian);
                                    break;
                                case "健康状况":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.JianKangZhuangKuang);
                                    break;
                                case "专业技术职务":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.ZhuanYeJiShuZhiWu);
                                    break;
                                case "熟悉专业有何专长":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.ShuXiZhuanYeYouHeZhuanChang);
                                    break;
                                case "全日制教育":
                                    XWPFParagraph pg = table.Rows[i].GetTableCells()[j + 1].Paragraphs[0];
                                    XWPFRun run = pg.CreateRun();
                                    run.SetText(ganBu.QuanRiZhiJiaoYu_XueLi);
                                    run.AddBreak();
                                    run.SetText(ganBu.QuanRiZhiJiaoYu_XueWei);
                                    break;
                                case "毕业院校系及专业":
                                    if (ifFirst)
                                    {
                                        XWPFParagraph pg2 = table.Rows[i].GetTableCells()[j + 1].Paragraphs[0];
                                        XWPFRun run2 = pg2.CreateRun();
                                        run2.SetText(ganBu.QuanRiZhiJiaoYu_XueLi_BiYeYuanXiaoXi);
                                        run2.AddBreak();
                                        run2.SetText(ganBu.QuanRiZhiJiaoYu_XueWei_BiYeYuanXiaoXi);
                                        ifFirst = false;
                                    }
                                    else
                                    {
                                        XWPFParagraph pg3 = table.Rows[i].GetTableCells()[j + 1].Paragraphs[0];
                                        XWPFRun run3 = pg3.CreateRun();
                                        run3.SetText(ganBu.ZaiZhiJiaoYu_XueLi_BiYeYuanXiaoXi);
                                        run3.AddBreak();
                                        run3.SetText(ganBu.ZaiZhiJiaoYu_XueWei_BiYeYuanXiaoXi);
                                        ifFirst = true;
                                    }
                                    break;
                                case "在职教育":
                                    XWPFParagraph pg4 = table.Rows[i].GetTableCells()[j + 1].Paragraphs[0];
                                    XWPFRun run4 = pg4.CreateRun();
                                    run4.SetText(ganBu.ZaiZhiJiaoYu_XueLi);
                                    run4.AddBreak();
                                    run4.SetText(ganBu.ZaiZhiJiaoYu_XueWei);
                                    break;
                                case "现任职务":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.XianRenZhiWu);
                                    break;
                                case "拟任职务":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.NiRenZhiWu);
                                    break;
                                case "拟免职务":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.NiMianZhiWu);
                                    break;
                                case "简历":
                                    int lastLine = 1;
                                    XWPFParagraph pg5 = table.Rows[i].GetTableCells()[j + 1].Paragraphs[0];
                                    XWPFRun run5 = pg5.CreateRun();
                                    foreach (JianLi jl in ganBu.JianLi)
                                    {
                                        if (lastLine != ganBu.JianLi.Count)
                                        {
                                            if (ifFirst)
                                            {
                                                run5.SetText(jl.KaiShiNianYue + "--" + jl.JieSuNianYue + "  " + jl.JingLi);
                                                ifFirst = false;
                                            }
                                            else
                                            {
                                                table.Rows[i].GetTableCells()[j + 1].AddParagraph().CreateRun().SetText(jl.KaiShiNianYue + "--" + jl.JieSuNianYue + "  " + jl.JingLi);
                                            }
                                            lastLine++;
                                        }
                                        else
                                        {
                                            table.Rows[i].GetTableCells()[j + 1].AddParagraph().CreateRun().SetText(jl.KaiShiNianYue + "--" + "    " + jl.JieSuNianYue + "    " + jl.JingLi);
                                        }
                                    }
                                    break;
                                case "奖惩情况":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.JiangChengQingKuang);
                                    break;
                                case "年核度结考果":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.NianDuKaoHeJieGuo);
                                    break;
                                case "任免理由":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.RenMianLiYou);
                                    break;
                                case "家庭主要成员及重要社会关系":
                                    List<JiaTingChengYuan> jtcyList = ganBu.JiaTingChengYuan;
                                    for (int m = 0; m < jtcyList.Count(); m++)
                                    {
                                        XWPFTableRow tr = table.Rows[i + m + 1];
                                        tr.GetTableCells()[1].SetText(jtcyList[m].ChengWei);
                                        tr.GetTableCells()[2].SetText(jtcyList[m].XingMing);
                                        tr.GetTableCells()[3].SetText(jtcyList[m].NianLing.ToString());
                                        tr.GetTableCells()[4].SetText(jtcyList[m].ZhengZhiMianMao);
                                        tr.GetTableCells()[5].SetText(jtcyList[m].GongZuoDanWeiJiZhiWu);

                                    }
                                    break;
                                case "呈报单位":
                                    table.Rows[i].GetTableCells()[j + 1].SetText(ganBu.ChengBaoDanWei);
                                    break;
                            }
                        }
                    }
                    num++;
                }





                System.IO.FileStream output = new System.IO.FileStream(targetFilePath, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite, System.IO.FileShare.ReadWrite);
                //写入文件
                docx.Write(output);

                output.Close();
                output.Dispose();

                logger.Info(targetFilePath + "文件生成成功！");
            }
            catch (Exception ex)
            {
                logger.Error(filename + "转换不成功");
                logger.Error(ex.ToString());

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
             stopwatch.Start(); 



                for (int rownum = 0; rownum < dataGridView1.Rows.Count; rownum++) {
                    string  filename = dataGridView1.Rows[rownum].Cells[0].Value.ToString();
                    string filePath = Path.GetDirectoryName(filename);
                    if (!Directory.Exists(".\\docx\\" + filePath.Substring(3, filePath.Length-3))) {
                        Directory.CreateDirectory(".\\docx\\" + filePath.Substring(3, filePath.Length - 3));
                    }


                    LrmxToDocx(filename);
                    string[] tmpArr = filename.Split('\\');
                    string  tmpFileName = tmpArr[tmpArr.Length  -1 ];
                    string newFileName = tmpFileName.Substring(0, tmpFileName.Length - 5) + "-任免审批表.docx";
                    if (File.Exists(".\\docx\\" + filePath.Substring(3, filePath.Length - 3) + "\\" +  newFileName)){

                        dataGridView1.Rows[rownum].Cells[1].Value = System.Environment.CurrentDirectory  +  "\\docx\\" + filePath.Substring(3, filePath.Length - 3) + "\\" + newFileName;
            }


                }


                stopwatch.Stop();
                logger.Info("程序执行耗时:" + stopwatch.ElapsedMilliseconds * 0.001 + "秒");
            }
            else {
                MessageBox.Show("必须先选择目录，或者所选目录下没有任免表文件", "警告");
            
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0) {
                try
                {
          
                    HSSFWorkbook excelBook = new HSSFWorkbook(); //创建工作簿Excel                   
                    NPOI.SS.UserModel.ISheet sheet1 = excelBook.CreateSheet("简历导出");
                    NPOI.SS.UserModel.IRow row1 = sheet1.CreateRow(0);//创建第一行
                    int rowStart  = 1;
                    row1.CreateCell(0).SetCellValue("姓名");
                    row1.CreateCell(1).SetCellValue("性别");
                    row1.CreateCell(2).SetCellValue("出生年月");
                    row1.CreateCell(3).SetCellValue("年龄");
                    row1.CreateCell(4).SetCellValue("参加工作时间");
                    row1.CreateCell(5).SetCellValue("入党时间");
                    row1.CreateCell(6).SetCellValue("开始时间");
                    row1.CreateCell(7).SetCellValue("结束时间");
                    row1.CreateCell(8).SetCellValue("时间差");
                    row1.CreateCell(9).SetCellValue("时间差2");
                    row1.CreateCell(10).SetCellValue("经历");
                    for (int l = 0; l< dataGridView1.Rows.Count; l++) {

                        GanBu ganBu = LrmxToClass(dataGridView1.Rows[l].Cells[0].Value.ToString());
                   
                        foreach (JianLi jl in ganBu.JianLi) {
                       
                        NPOI.SS.UserModel.IRow row = sheet1.CreateRow( rowStart );
                        row.CreateCell(0).SetCellValue(ganBu.XingMing);
                        row.CreateCell(1).SetCellValue(ganBu.XingBie);
                        row.CreateCell(2).SetCellValue(ganBu.ChuShengNianYue);
                        row.CreateCell(3).SetCellValue(ganBu.NianLing);
                        row.CreateCell(4).SetCellValue(ganBu.CanJiaGongZuoShiJian);
                        row.CreateCell(5).SetCellValue(ganBu.RuDangShiJian);
                        row.CreateCell(6).SetCellValue(jl.KaiShiNianYue);
                        row.CreateCell(7).SetCellValue(jl.JieSuNianYue);
                        int nian = 0;
                        string chazhi = "";
                       string  shijiancha ="";
                        if (!string.IsNullOrEmpty(jl.KaiShiNianYue) && !string.IsNullOrEmpty(jl.JieSuNianYue)  &&  jl.KaiShiNianYue.Length > 6 && jl.JieSuNianYue.Length >6) {
                                try
                                {



                                    nian = int.Parse(jl.JieSuNianYue.Substring(0, 4)) - int.Parse(jl.KaiShiNianYue.Substring(0, 4));
                                    if (int.Parse(jl.JieSuNianYue.Substring(5, 2)) < int.Parse(jl.KaiShiNianYue.Substring(5, 2)))
                                    {
                                        nian--;
                                        int yue = ((12 - int.Parse(jl.KaiShiNianYue.Substring(5, 2)) + int.Parse(jl.JieSuNianYue.Substring(5, 2))));
                                        chazhi = (nian > 0 ? nian + "年" : "") + (yue > 0 ? yue + "月" : "");
                                        shijiancha = ((double)nian + yue * 1.0f / 12).ToString("F2");
                                    }
                                    else
                                    {
                                        int yue = (int.Parse(jl.JieSuNianYue.Substring(5, 2)) - int.Parse(jl.KaiShiNianYue.Substring(5, 2)));
                                        chazhi = (nian > 0 ? nian + "年" : "") + (yue > 0 ? yue + "月" : "");
                                        shijiancha = ((double)nian + yue * 1.0f / 12).ToString("F2");
                                    }
                                }catch(Exception ex)
                                {
                                    logger.Error(ganBu.XingMing +"经历处理存在问题");

                                }


                            }


                        row.CreateCell(8).SetCellValue(chazhi.ToString());
                        row.CreateCell(9).SetCellValue(shijiancha);

                        row.CreateCell(10).SetCellValue(jl.JingLi);

                            rowStart++;

                        }

                        dataGridView1.Rows[l].Cells[1].Value = "OK";

                    }
                    FileStream writeFile = new FileStream(@".\export.xlsx", FileMode.Create);

                    excelBook.Write(writeFile);
                    writeFile.Close();
                }
                catch (Exception ex) { 
                
                    logger.Error(ex.ToString ());   
                
                
                }
            
            
            }
        }
    }
}
