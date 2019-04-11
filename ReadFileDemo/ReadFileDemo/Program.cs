using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ReadFileDemo
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            List<Vehicle> vehicleList = new List<Vehicle>();
            StreamReader sr = new StreamReader(@"C:\Users\Summerki\Desktop\活\test.fzp", Encoding.Default);
            Regex reg = new Regex(@"\s{0,}\d*;\s{0,}[-+];\s{0,}\d*\.\d*;\s{0,}\d*\.\d*;\s{0,}\d*\.\d*;");
            string line;
            Boolean flag;
            string[] strArray;
            Vehicle v;
            while ((line = sr.ReadLine()) != null)
            {
                flag = reg.IsMatch(line);
                if (flag)
                {
                    line = line.Trim();
                    strArray = Regex.Split(line, ";");
                    v = new Vehicle();
                    v.VehNr = strArray[0].Trim();
                    v.Queue = strArray[1].Trim();
                    v.QTim = strArray[2].Trim();
                    v.t = strArray[3].Trim();
                    v.RworldldX = strArray[4].Trim();

                    vehicleList.Add(v);
                }
                
            }

            sr.Close();
            sr.Dispose();

            // 写入到excel测试
            //ExportExcel(vehicleList);



            // 首次由负到正的元素写入excel测试
            ExportFirstVehicleExcel(FindBestSolutionToExcel(vehicleList));

            
            Console.WriteLine("写入成功");
            Console.ReadKey();
        }

        #region 将原始数据导出到excel，暂时用不到
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
            for (int i = 0; i < list.Count; i++)
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
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sb.Append(dt.Rows[i][j].ToString() + "\t");
                }
                sb.Append(Environment.NewLine);
            }
            System.Windows.Forms.Clipboard.SetText(sb.ToString());


            // 新建excel应用
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
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
            FileStream file = new FileStream(filePath, FileMode.Create);
            //关闭释放流，不然没办法写入数据
            file.Close();
            file.Dispose();

            //保存写入的数据，这里还没有保存到磁盘
            workbook.Saved = true;
            //保存到指定的路径
            workbook.SaveCopyAs(filePath);
        }
        #endregion



        /// <summary>
        /// 找出从负变正的那一行数据,返回一个集合
        /// </summary>
        /// <param name="list"></param>
        public static List<TargetVehicle> FindBestSolutionToExcel(List<Vehicle> list)
        {
            // 寻找首次由负变正的车辆的list集合
            List<Vehicle> targetVList = new List<Vehicle>();
            Boolean flag = false; // 标志位放在foreach循环里默认表示该元素没有重复

            // 首先将所有为+的都找出来
            for(int i = 0; i < list.Count; i++)
            {
                if(list[i].Queue == "-")
                {
                    continue;
                }

                if(targetVList.Count == 0)
                {
                    targetVList.Add(list[i]);
                    continue;
                }

                foreach(Vehicle temp in targetVList)
                {
                    if(list[i].VehNr == temp.VehNr) // 证明不是第一次/重复
                    {
                        flag = true;
                        continue;
                    }
                }
                if(flag == false)
                {
                    targetVList.Add(list[i]);
                }
                flag = false;

            }

            // 对targetVList数据进行对应的处理
            TargetVehicle targetVehicle;
            // 最终指定格式的list列表
            List<TargetVehicle> finalList = new List<TargetVehicle>();
            for(int i = 0; i < targetVList.Count; i++)
            {
                targetVehicle = new TargetVehicle();
                // 这里按照0-70为第一个cycle，70-140为第2个cycle计算
                targetVehicle.cycle = ((int)(Convert.ToDouble(targetVList[i].t) / 70) + 1) + "";
                targetVehicle.VehNr = targetVList[i].VehNr;
                targetVehicle.leftTime = (Convert.ToDouble(targetVList[i].t) - ((int.Parse(targetVehicle.cycle) - 1) * 70)) + "";
                targetVehicle.RworldldX = targetVList[i].RworldldX;
                finalList.Add(targetVehicle);
            }

            return finalList;

        }



        public static void ExportFirstVehicleExcel(List<TargetVehicle> list)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            // 列
            dt.Columns.Add("cycle", typeof(string));
            dt.Columns.Add("VehNr", typeof(string));
            dt.Columns.Add("leftTime", typeof(string));
            dt.Columns.Add("RworldldX", typeof(string));
            // 行
            for (int i = 0; i < list.Count; i++)
            {
                dt.Rows.Add(list[i].cycle, list[i].VehNr, list[i].leftTime, list[i].RworldldX);
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
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sb.Append(dt.Rows[i][j].ToString() + "\t");
                }
                sb.Append(Environment.NewLine);
            }
            System.Windows.Forms.Clipboard.SetText(sb.ToString());


            // 新建excel应用
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
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
            string filePath = @"C:\Users\Summerki\Desktop\活\target.xlsx";
            //创建文件
            FileStream file = new FileStream(filePath, FileMode.Create);
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
