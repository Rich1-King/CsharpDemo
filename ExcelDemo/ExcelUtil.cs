/**
 *
 * 读取excel需要添加com引用，引用的com为 Mirosoft.Office.16.0.Object.Libary
 * 
 **/

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelDemo
{
    public class ExcelUtil
    {

        public static void writeExcel(string fileName)
        {
            /*if (!File.Exists(fileName))
            {
                File.Create(fileName);
            }*/
            object obj = System.Reflection.Missing.Value;
            Application app = new Application();
            app.Visible = false; //后台运行
            app.Workbooks.Add(obj);
            app.DisplayAlerts = false;//设置不显示确认修改提示
            Workbook wb = app.ActiveWorkbook; //创建workbook，是内存中的excel文件
            Sheets sheets = wb.Sheets; //获取workbook中的sheet
            Worksheet ws = sheets.Add(); //添加一个sheet

            ws.Name = "表一"; //设置sheet的名字
            ws.Cells[1,1].Value = "编号";
            ws.Cells[1,2].Value = "名字";
            ws.Cells[1,3].Value = "随机数";
            int j = 1;
            Random rd = new Random();
            for (int i=2; i<10; i++)
            {           
                ws.Cells[i,1].Value = j;
                ws.Cells[i,2].Value = "张三";
                ws.Cells[i,3].Value = rd.Next(100,999);
            }
            //fileName = "@"+fileName;
            ws.SaveAs(fileName); //将内存中的文件保存到硬盘中          
            wb.Close(false, Type.Missing, Type.Missing); //关闭workbook
            app.Quit();
        }

        public static void readExcel(string fileName)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Open(fileName); //打开要读取的excel文件
            Worksheet ws = wb.Worksheets[1]; //获取sheet
            int rowCount = ws.UsedRange.Rows.Count; //获取该sheet的行数
            int columnCount = ws.UsedRange.Columns.Count; //获取该sheet的列数
            for(int i=1; i<=rowCount; i++)
            {
                Console.Write(ws.Cells[i,1].Text + "     ");
                Console.Write(ws.Cells[i,2].Text + "     ");
                Console.Write(ws.Cells[i,3].Text + "     ");
                Console.WriteLine();
            }
            wb.Close();
            app.Quit();
            Console.ReadKey();
        }
    }

       
}
