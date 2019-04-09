using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace TestReg
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            List<string> arrayList = new List<string>();

            List<Vehicle> vehicleList = new List<Vehicle>();


            Regex reg = new Regex(@"\s{0,}\d*;\s{0,}[-+];\s{0,}\d*\.\d*;\s{0,}\d*\.\d*;\s{0,}\d*\.\d*;");
            string str = "        1;     -; 0.0;     7.0;  513.7792; ";

            Boolean flag = reg.IsMatch(str);
            if (flag)
            {
                Console.WriteLine("匹配成功");
                str = str.Trim();
                string[] strArray = Regex.Split(str, ";");

                Vehicle v = new Vehicle();
                v.VehNr = strArray[0].Trim();
                v.Queue = strArray[1].Trim();
                v.QTim = strArray[2].Trim();
                v.t = strArray[3].Trim();
                v.RworldldX = strArray[4].Trim();

                vehicleList.Add(v);

                //for (int i = 0; i < strArray.Length; i++)
                //{
                    
                //    arrayList.Add(strArray[i].Trim()); // 再次去掉左右两边的空格
                //}


                // 打印输出
                foreach(string temp in arrayList)
                {
                    Console.WriteLine(temp);
                }
            }


            // 写入到指定文件测试
            //StreamWriter sw = new StreamWriter(@"C:\Users\Summerki\Desktop\活\temp.txt");
            //foreach (string temp in arrayList)
            //{
            //    sw.Write(temp + "\r\n");
            //}
            //sw.Close();

            // 写入到excel测试
            ExportExcel(vehicleList);

            Console.ReadKey();
        }



        public static void ExportExcel(List<Vehicle> list)
        { 
            System.Data.DataTable dt = new System.Data.DataTable();
            // 列
            dt.Columns.Add("VehNr", typeof(string));
            dt.Columns.Add("Queue", typeof(string));
            dt.Columns.Add("QTim", typeof(string));
            dt.Columns.Add("t", typeof(string));
            dt.Columns.Add("RworldldX", typeof(string));
            // 行
            for(int i = 0; i < list.Count; i++)
            {
                dt.Rows.Add(list[i].VehNr, list[i].Queue, list[i].QTim, list[i].t, list[i].RworldldX);
            }

            // 写入标题
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sb.Append(dt.Columns[i].ColumnName.ToString() + "\t");
            }

            // 加入换行符
            sb.Append(Environment.NewLine);

            // 写入内容
            for(int i = 0; i < dt.Rows.Count; i++)
            {
                for(int j = 0; j < dt.Columns.Count; j++)
                {
                    sb.Append(dt.Rows[i][j].ToString() + "\t");
                }
                sb.Append(Environment.NewLine);
            }
            System.Windows.Forms.Clipboard.SetText(sb.ToString());


            // 新建excel应用
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if(excelApp == null)
            {
                return;
            }

            //设置为不可见，操作在后台执行，为 true 的话会打开 Excel
            excelApp.Visible = false;
            //初始化工作簿
            Microsoft.Office.Interop.Excel.Workbooks workbooks = excelApp.Workbooks;
            //新增加一个工作簿，Add（）方法也可以直接传入参数 true
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            //同样是新增一个工作簿，但是会弹出保存对话框
            //Excel.Workbook workbook = workbooks.Add(true);

            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets.Add();

            worksheet.Paste();

            //新建一个 Excel 文件
            //string filePath = @"C:\Users\Lenovo\Desktop\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx";
            string filePath = @"C:\Users\Summerki\Desktop\活\test.xlsx";
            //创建文件
            FileStream file = new FileStream(filePath, FileMode.CreateNew);
            //关闭释放流，不然没办法写入数据
            file.Close();
            file.Dispose();

            //保存写入的数据，这里还没有保存到磁盘
            workbook.Saved = true;
            //保存到指定的路径
            workbook.SaveCopyAs(filePath);
        }
    }
}
