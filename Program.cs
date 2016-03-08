using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;

namespace excelconvert
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    /// <summary>
    /// 定义数据结构
    /// </summary>
    public class data_core
    {
        public data_core(data_core core)
        {
            this.name = core.name;
            this.value = core.value;
        }
        public data_core(string name, string value)
        {
            this.name = name;
            this.value = value;
        }
        public string name;//数据类型
        public string value;//实际内容
    }

    /// <summary>
    /// 一行数据，封装
    /// </summary>
    public class data_line
    {
        public data_line()
        {
            line = new List<data_core>();
        }
        private List<data_core> line;

        public data_core this[int index] 
        {
            get {
                if (index < 0 || index >= line.Count)
                {
                    return null;
                }
                return line[index];
            }
            set {
                if (index < 0 || !(value is data_core))
                {
                    return;
                }
                else
                {
                    if (index < line.Count)
                    {
                        line[index] = value;
                    }
                    else
                    {
                        //data_core new_core = new data_core(value);
                        //line.Add(new_core);
                        line.Add(value);
                    }
                }
            }
        }
    }

    public class data_convert
    {
        public data_convert()
        {
            table1 = new List<data_line>();
            table2 = new List<data_line>();
        }
        private List<data_line> table1;
        private List<data_line> table2;

        //public static void inittable1()
        //{ 
            
        //}

        public void convert()
        {
            
        }

        public void save()
        { 
            
        }

        public data_line this[int type,int index]
        {
            get
            {
                if (type == 0)
                {
                    if (index < 0 || index >= table1.Count)
                    {
                        return null;
                    }
                    return table1[index];
                }
                else
                {
                    if (index < 0 || index >= table2.Count)
                    {
                        return null;
                    }
                    return table2[index];
                }
            }
            set
            {
                if (type == 0)
                {
                    if (index < 0 || !(value is data_line))
                    {
                        return;
                    }
                    else
                    {
                        if (index < table1.Count)
                        {
                            table1[index] = value;
                        }
                        else
                        {
                            //data_line new_line = new data_line();
                            //table1.Add(new_line);
                            table1.Add(value);
                        }
                    }    
                }
                else
                {
                    if (index < 0 || !(value is data_line))
                    {
                        return;
                    }
                    else
                    {
                        if (index < table2.Count)
                        {
                            table2[index] = value;
                        }
                        else
                        {
                            table2.Add(value);
                        }
                    }    
                }
            }
        }
        
    }

    public class input_templete
    {
        public string date;                 //业务日期
        public string clientname;           //客户名称
        public string org_serial;           //原单号
        public string inland_serial;        //国内单号
        public string weight;               //重量
        public string destination;          //目的地
        public string receiver_name;        //收件人
        public string receiver_id;          //收件人身份证
        public string receiver_addr;        //收件地址
        public string receiver_phone;       //收件电话
        public string receiver_code;        //收件邮编
        public string product_name;         //商品名称
        public string product_standard;     //物品规格
        public string product_price;        //商品单价
        public string product_count;        //商品数量
        public string product_unit;         //商品单位名称
        public string tax_rate;             //关税率
        public string multiple_name_count;  //多品名件数
        public string recv_name_addr_id;    //收名址ID
        public string inner_ctl;            //内控
        public string batch;                //批次
        public string sender_name;          //发件人
        public string sender_phone;         //发件电话
        public string sender_addr;          //发件地址
        public string object_code;          //物品丙码
        public string receive_province;     //收件省州
        public string receive_city;         //收件城市
        public string express_category;     //快递类别
        public string org_place;            //出发地
        public string backup_one;           //备用一
        public string belong_site;          //所属站点
        public string refer_number;         //参考号
        public string backup_five;          //备用五
        public string recorder;             //录入
        public string status;               //状态
    }



    public class ExcelHelper
    { 
                /// <summary>
        /// Excel应用程序
        /// </summary>
        Microsoft.Office.Interop.Excel._Application appExcel = null;
        /// <summary>
        /// Excel工作簿
        /// </summary>
        Microsoft.Office.Interop.Excel._Workbook workBook = null;
        /// <summary>
        /// Excel工作表
        /// </summary>
        Microsoft.Office.Interop.Excel._Worksheet workSheet = null;
        object oMissing = System.Reflection.Missing.Value;
        /// <summary>
        /// 创建WorkBook
        /// </summary>
        /// <param name="excelPath">Excel路径</param>
        public ExcelHelper(string Path)                             //构造函数
        {
            new FileInfo(Path.ToString()).Attributes = FileAttributes.Normal;
            appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Application.Workbooks.Open(Path, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            workBook = appExcel.Workbooks[1];
        }
        /// <summary>
        /// 获得进程的标识
        /// </summary>
        /// <param name="hwnd">IntPtr</param>
        /// <param name="ID">int</param>
        /// <returns>int</returns>
        [System.Runtime.InteropServices.DllImport("User32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        /// <summary>
        /// 关闭Excel进程
        /// </summary>
        public void KillExcel()                                        
        {
            if (appExcel == null)
            {
                return;
            }
            IntPtr t = new IntPtr(appExcel.Hwnd);
            int k = 0;
            GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process ps = System.Diagnostics.Process.GetProcessById(k);
            ps.Kill();
        }
        /// <summary>
        /// 关闭Excel
        /// </summary>
        public void CloseExcelApp()
        {
            if (workBook == null) return;
            workBook.Close(true, oMissing, oMissing);
            appExcel.Quit();
        }
        /// <summary>
        /// 另存为Excel
        /// </summary>
        /// <param name="newExcelName">Excel路径</param>
        public void SaveAsExcel(string newExcelName)
        {
            if (workBook == null) return;
            appExcel.Application.DisplayAlerts = false;
            workBook.SaveAs(newExcelName, oMissing, oMissing, oMissing, oMissing, oMissing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, oMissing, oMissing, oMissing, oMissing, oMissing);
        }
        /// <summary>
        /// 保存Excel
        /// </summary>
        public void SaveExcel()
        {
            if (workBook == null) return;
            workBook.Save();
        }
        /// <summary>
        /// 把数据写入Excel
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="sheetName">工作表名</param>
        /// <param name="newIndex">工作表行索引</param>
        public void DataTableToExcel(System.Data.DataTable dt, string sheetName, int newIndex)
        {
            //存在Sheet页为true，不存在为false
            bool flag = false;
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workBook.Worksheets) //获取工作表
            {
                if (sheet.Name.Trim() == sheetName.Trim())
                {
                    workSheet = sheet;
                    flag = true;
                    break;
                }
            }
            if (flag == false)   //如果表单不存在，创建
            {
                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.Add(oMissing, oMissing, oMissing, oMissing);
                workSheet.Name = sheetName;
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    workSheet.Cells[newIndex + i + 2, j + 1] = dt.Rows[i][j].ToString();
                }
            }
            workSheet = null;
        }
    }

}
