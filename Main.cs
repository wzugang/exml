using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Data.OleDb;
using System.Threading;
using System.Xml;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace excelconvert
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
        //要验证顺序
        public static Dictionary<string, exml_col> compiled_list = new Dictionary<string, exml_col>();

        private void button1_Click(object sender, EventArgs e)
        {
            
            //加载xml文件
            DataSet excel_ds = new DataSet("excel");
            DataSet import_ds = new DataSet("import");
            DataSet export_ds = new DataSet("export");
            System.Data.DataTable dt=null;
            string openpath = null;
            open1.FileName = ""; 
            int index = 0;
            open1.DefaultExt = ".xlsx";
            open1.Filter = "XML格式文件(*.xml)|*.xml|所有文件(*.*)|*.*";
            if (open1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                openpath = open1.FileName;
            }
            else
            {
                return;
            }
            if (openpath == "" || openpath == null)
            {
                return;
            }

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(openpath);
            }
            catch (Exception )
            {
                
                MessageBox.Show("转换脚本加载失败，请检查脚本是否符合xml语法");
            }

            //exml预处理
            XmlNode node = doc.DocumentElement;
            if (node.Name == "excel" && node.HasChildNodes)
            {
                XmlNode child = node.FirstChild;
                //导入表
                while(child != null && child.Name == "import")
                {
                    if (child.HasChildNodes)
                    {
                        XmlNode col = child.FirstChild;
                        dt = new System.Data.DataTable();
                        while (col != null && col.Name == "col")
                        {
                            if (col.Attributes != null && col.Attributes["name"] != null && col.Attributes["name"].Value != "")
                            {
                                dt.Columns.Add(col.Attributes["name"].Value);
                            }
                            else
                            {
                                MessageBox.Show("exml语法错误");
                                return;
                            }
                            col = col.NextSibling;
                        }
                        dt.TableName =  open1.FileName.Substring(0,open1.FileName.LastIndexOf('\\')+1) + child.Attributes["name"].Value;
                        import_ds.Tables.Add(dt);
                    }
                    else
                    {
                        MessageBox.Show("exml语法错误");
                        return;
                    }

                    child = child.NextSibling;
                }
                if (child != null)
                {
                    if (child.Name != "export")
                    {
                        MessageBox.Show("exml语法错误");
                        return;
                    }
                    while (child != null && child.Name == "export")
                    {
                        //导出处理
                        if (child.HasChildNodes)
                        {
                            //生成导出表结构
                            dt = new System.Data.DataTable();
                            XmlNode col = child.FirstChild;
                            while (col != null && col.Name == "col")
                            {
                                if (col.Attributes != null && col.Attributes["name"] != null && col.Attributes["name"].Value != "")
                                {
                                    dt.Columns.Add(col.Attributes["name"].Value);
                                }
                                else
                                {
                                    MessageBox.Show("exml语法错误");
                                    return;
                                }
                                col = col.NextSibling;
                            }
                            dt.TableName = open1.FileName.Substring(0, open1.FileName.LastIndexOf('\\') + 1) + child.Attributes["name"].Value;
                            export_ds.Tables.Add(dt);

                            //开始转换，保存
                            exml_get_import_table(excel_ds,import_ds.Tables[0].TableName);
                            
                            //转化到import_ds中
                            if (exml_excel_to_ds(import_ds, excel_ds.Tables[0], import_ds.Tables[0].Columns[0].ColumnName, "合计") != 0)
                            {
                                return;
                            }

                            //import_ds转化为export_ds
                            int count = 0;
                            foreach (DataRow r in import_ds.Tables[0].Rows)
                            {
                                col = child.FirstChild;
                                DataRow dr = export_ds.Tables[0].NewRow();

                                while (col != null)
                                {
                                    if (col.Attributes["type"].Value == "index")
                                    {
                                        dr[col.Attributes["name"].Value] = count + 1;
                                    }
                                    else if (col.Attributes["type"].Value == "build")
                                    {
                                        //遍历处理append子节点
                                        if (col.HasChildNodes)
                                        {
                                            XmlNode ap = col.FirstChild;
                                            int num = is_build_array(r,col);
                                            StringBuilder sb = new StringBuilder();
                                            if (num > 0)
                                            {
                                                index = 0;
                                                for (int i = 0; i < num; i++)
                                                {
                                                    ap = col.FirstChild;
                                                    while (ap != null)
                                                    {
                                                        if (ap.Attributes["value"] != null)
                                                        {
                                                            if (ap.Attributes["split"] != null)
                                                            {
                                                                sb.Append((ap.Attributes["value"].Value.ToString().Split(ap.Attributes["split"].Value.Trim()[0]))[index]);
                                                            }
                                                            else
                                                            {
                                                                if (ap.Attributes["ifend"] != null && ap.Attributes["ifend"].Value == "0")
                                                                {
                                                                    if (index != num-1)
                                                                    {
                                                                        sb.Append(ap.Attributes["value"].Value.ToString());
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    sb.Append(ap.Attributes["value"].Value.ToString());
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (ap.Attributes["split"] != null)
                                                            {
                                                                sb.Append((r[ap.Attributes["field"].Value.ToString()].ToString().Split(ap.Attributes["split"].Value.Trim()[0]))[index]);
                                                            }
                                                            else
                                                            {
                                                                sb.Append(r[ap.Attributes["field"].Value.ToString()].ToString());
                                                            }
                                                        }
                                                        ap = ap.NextSibling;
                                                    }
                                                    index++;
                                                }
                                            }
                                            else
                                            {
                                                while (ap != null)
                                                {
                                                    if (ap.Attributes["value"] != null)
                                                    {
                                                        sb.Append(ap.Attributes["value"].Value.ToString());
                                                    }
                                                    else
                                                    {
                                                        sb.Append(r[ap.Attributes["field"].Value.ToString()].ToString());
                                                    }
                                                    ap = ap.NextSibling;
                                                }
                                            }
                                            dr[col.Attributes["name"].Value] = sb.ToString();
                                        }
                                        else
                                        {
                                            MessageBox.Show("build数据类型不能为空");
                                            return;
                                        }
                                    }
                                    else
                                    {
                                       // 这里什么都不用做
                                    }
                                    col = col.NextSibling;
                                }
                                export_ds.Tables[0].Rows.Add(dr);                                
                                count++;
                            }

                            //转换完成后保存
                            exml_save_excel(export_ds.Tables[0], export_ds.Tables[0].TableName);
                        }
                        else
                        {
                            MessageBox.Show("exml语法错误");
                            return;
                        }

                        child = child.NextSibling;
                    }
                }
            }
            else
            {
                MessageBox.Show("exml脚本错误");
            }

            MessageBox.Show("转换成功");
        }
        public void exml_get_import_table(DataSet ds, string name)
        {
            System.Data.DataTable dt=null;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (app == null)
            {
                return;
            }
            app.Application.Workbooks.Open(name, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            Workbook wb = app.Workbooks[1];
            string con_cmd = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + name + ";Extended Properties = 'Excel 8.0;HDR=YES;IMEX=1;'";
            OleDbConnection cnn = new OleDbConnection(con_cmd);
            cnn.Open();
            OleDbDataAdapter dap = null;
            string sql = null;
            foreach (Worksheet ws in wb.Worksheets)
            {
                dt = new System.Data.DataTable();
                sql = "select * from [" + ws.Name.ToString() + "$]";
                dap = new OleDbDataAdapter(sql, cnn);
                dap.Fill(dt);
                ds.Tables.Add(dt);
            }
            cnn.Close();
            wb.Close(false, Missing.Value, Missing.Value);
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);//释放资源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);//释放资源
            wb = null;
            app = null;
            GC.Collect();
        }

        public void exml_save_excel(System.Data.DataTable dt,string savename)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            if (app == null)
            {
                MessageBox.Show("应用创建失败！");
                return;
            }

            //创建sheets
            Workbook wb = app.Workbooks.Add(true);
            Worksheet sheet = wb.Worksheets[1];
            sheet.Name = "sheet3";
            sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            sheet.Name = "sheet2";
            sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            sheet.Name = "sheet1";
            sheet = wb.Worksheets[1];
            //给sheet1写数据
            //MessageBox.Show(savetable.Columns.Count.ToString());

            for (int j = 0; j < dt.Columns.Count; j++)
            {
                sheet.Cells[1, j + 1] = dt.Columns[j].ColumnName;
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                }
            }

            //设置标题格式
            Microsoft.Office.Interop.Excel.Range title = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, dt.Columns.Count]];//选取单元格，选取一行或多行
            title.EntireColumn.AutoFit();     //自动调整列宽 
            //title.Merge(true);//合并单元格  
            //title.Value2 = ""; //设置单元格内文本  
            title.Font.Name = "宋体";//设置字体  
            title.Font.Size = 12;//字体大小  
            title.Font.Bold = true;//加粗显示  
            title.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
            title.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中  
            //title.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//设置边框  
            //title.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;//边框常规粗细  

            //保存
            wb.SaveAs(savename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            wb.Close(false, Missing.Value, Missing.Value);
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);//释放资源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);//释放资源
            wb = null;
            sheet = null;
            app = null;
            GC.Collect();
        }


        public int exml_excel_to_ds(DataSet ds,System.Data.DataTable dt,string start,string end)
        {
            int count = 0;
            foreach (DataRow r in dt.Rows)
            {
                if (r[0] == null || r[0].ToString().Trim() == "")
                {
                    continue;
                }
                if (r[0].ToString().Trim() == start)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (r[i].ToString().Trim() != "")
                        {
                            count++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (count != ds.Tables[0].Columns.Count)
                    {
                        MessageBox.Show(string.Format("导入表不符合要求count={0},columns={1}", count, ds.Tables[0].Columns.Count));
                        return -1;
                    }
                    else
                    {
                        continue;
                    }
                }

                if (r[0].ToString().StartsWith(end))
                {
                    break;
                }
                DataRow dr = ds.Tables[0].NewRow();

                for (int i = 0; i < count; i++)
                {
                    dr[i] = r[i].ToString().Trim();
                }
                ds.Tables[0].Rows.Add(dr);
            }
            return 0;
        }

        public int is_build_array(DataRow dr, XmlNode node)
        {
            foreach (XmlNode item in node.ChildNodes)
            {
                if (item.Attributes["type"].Value == "array")
                {
                    if (item.Attributes["value"] == null)
                    {
                        return (dr[item.Attributes["field"].Value].ToString().Split(item.Attributes["split"].Value.Trim()[0])).Length;
                    }
                    else
                    {
                        return item.Attributes["value"].Value.ToString().Split(item.Attributes["split"].Value.Trim()[0]).Length;
                    }
                }
            }
            return -1;
        }

    }

    //下面的代码暂且没用，以后升级使用

    public class exml
    {
        
        public static void exml_format_file(string file)
        {
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(file);
            XmlTextWriter writer = new XmlTextWriter(file, Encoding.GetEncoding("GB2312"));
            writer.Formatting = Formatting.Indented;
            writer.WriteStartDocument();
            //写数据到文件
            exml_node_to_file(xmldoc.DocumentElement, writer);
            //文档结束
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();
        }
        public static void exml_node_to_file(XmlNode xml, XmlTextWriter writer)
        {
            writer.WriteStartElement(xml.Name);
            if (xml.Attributes != null)
            {
                for (int j = 0; j < xml.Attributes.Count; j++)
                {
                    writer.WriteStartAttribute(xml.Attributes[j].Name);
                    writer.WriteString(xml.Attributes[j].Value);
                    writer.WriteEndAttribute();
                }
            }
            if (xml.HasChildNodes)
            {
                XmlNodeList xnl = (XmlNodeList)xml.ChildNodes;
                for (int i = 0; i < xnl.Count; i++)
                {
                    if (xnl.Item(i).Name == "#comment")
                    {
                        writer.WriteComment(xnl.Item(i).Value.ToString().Trim('\r', '\n', '\t', ' '));
                    }
                    else if (xnl.Item(i).Name == "#text")
                    {
                        writer.WriteString(xnl.Item(i).Value.ToString().Trim('\r', '\n', '\t', ' '));
                    }
                    else
                    {
                        exml_node_to_file(xnl.Item(i), writer);
                    }
                }
            }
            writer.WriteEndElement();
        }

        public static string exml_get_node_file(XmlNode xml)
        {
            return xml.BaseURI.Replace("file:///", "");
        }

        #region exml转义处理
        public static string exml_escape_words(string xmlstring)
        {
            //XML转义字符
            xmlstring = xmlstring.Replace("&", "&amp;");
            xmlstring = xmlstring.Replace("<", "&lt;");
            xmlstring = xmlstring.Replace(">", "&gt;");
            xmlstring = xmlstring.Replace("\"", "&quot;");
            xmlstring = xmlstring.Replace("\'", "&apos;");
            //字符转义
            xmlstring = xmlstring.Replace("\a", "&#x7;");
            xmlstring = xmlstring.Replace("\b", "&#x8;");
            xmlstring = xmlstring.Replace("\f", "&#xC;");
            xmlstring = xmlstring.Replace("\n", "&#xA;");
            xmlstring = xmlstring.Replace("\r", "&#xD;");
            xmlstring = xmlstring.Replace("\t", "&#x9;");
            xmlstring = xmlstring.Replace("\v", "&#xB;");
            xmlstring = xmlstring.Replace("\0", "&#x0;");

            return xmlstring;
        }
        public static string exml_escape_text(string xmlstring)
        {
            //XML转义字符
            xmlstring = xmlstring.Replace("&", "&amp;");
            xmlstring = xmlstring.Replace("<", "&lt;");
            //字符转义
            xmlstring = xmlstring.Replace("\a", "&#x7;");
            xmlstring = xmlstring.Replace("\b", "&#x8;");
            xmlstring = xmlstring.Replace("\f", "&#xC;");
            xmlstring = xmlstring.Replace("\n", "&#xA;");
            xmlstring = xmlstring.Replace("\r", "&#xD;");
            xmlstring = xmlstring.Replace("\t", "&#x9;");
            xmlstring = xmlstring.Replace("\v", "&#xB;");
            xmlstring = xmlstring.Replace("\0", "&#x0;");
            return xmlstring;
        }

        public static string exml_escape_attr(string xmlstring)
        {
            //XML转义字符
            xmlstring = xmlstring.Replace("&", "&amp;");
            xmlstring = xmlstring.Replace("<", "&lt;");
            xmlstring = xmlstring.Replace("\"", "&quot;");
            //字符转义
            xmlstring = xmlstring.Replace("\a", "&#x7;");
            xmlstring = xmlstring.Replace("\b", "&#x8;");
            xmlstring = xmlstring.Replace("\f", "&#xC;");
            xmlstring = xmlstring.Replace("\n", "&#xA;");
            xmlstring = xmlstring.Replace("\r", "&#xD;");
            xmlstring = xmlstring.Replace("\t", "&#x9;");
            xmlstring = xmlstring.Replace("\v", "&#xB;");
            xmlstring = xmlstring.Replace("\0", "&#x0;");
            return xmlstring;
        }
        #endregion

    }

    public class exml_col
    {
        //指示列表个数
        int count;
        public List<string> cols;
        //0表示string，1表示array，2表示单字符串，3表示字符串数组
        public List<int> types;
        //当前列名称
        public string name;
    }
}
