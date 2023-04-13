using ExcelIO.Net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

namespace ExcelIO.Demo.Net
{
    public class Program
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            ExcelIODemo excelIODemo = new ExcelIODemo();
            excelIODemo.ToExcel();
            excelIODemo.FromExcel();

            Console.WriteLine("Hello World!");
            Console.ReadKey();
        }

        class ExcelIODemo
        {
            private string[][] kvArr = new string[][]
            {
                new string[] { "ID", "序号"},
                new string[] { "Name", "名称"},
                new string[] { "Number", "编号"},
                new string[] { "Email", "邮箱"},
                new string[] { "PhoneNumber", "手机号"},
                new string[] { "Address", "地址"}
            };

            /// <summary>
            /// 把数据导出到 Excel
            /// </summary>
            public void ToExcel()
            {
                //采用封装的 Aspose 组件进行 Excel 操作
                AsposePlugin asposePlugin = new AsposePlugin();
                IExcelDataIO excelDataIO = new ExcelDataIO(asposePlugin);
                ExcelSheet excelSheet = ExcelSheet.Instance;

                CellProperty cellProperty = null;
                foreach (var kv in kvArr)
                {
                    //建立 Excel 表头名称与字段的映射关系
                    excelSheet.AddMapping(kv[1].Trim(), kv[0].Trim().ToLower());

                    //[可省略] 设置单元格属性
                    cellProperty = excelSheet.LastCellProperty;
                    cellProperty.setBackgroundColor("#0a7eae");
                    cellProperty.setForeColor("#ffffff");
                    cellProperty.width = 15;
                    cellProperty.height = 20;
                    cellProperty.textAlign = TextAlign.Center | TextAlign.Middle;
                    cellProperty.cellDataType = CellDataType.Text;
                    cellProperty.isBold = true;
                    cellProperty.fontFamily = "黑体";
                }

                string fPath = "D:\\abc.xlsx";
                //把列头关系、及属性信息设置到 Excel
                excelDataIO.ToExcelWithProperty(excelSheet: excelSheet, fPath);

                //获取 DataTable 数据
                DataTable dt = GetDbSource();

                ExcelRowChildren cellProperties = null;
                string key = "";
                foreach (DataRow item in dt.Rows)
                {
                    // 向 ExcelDataIO 对象获取具备<列头文本>与<字段名称>映射关系的行(列集合)
                    cellProperties = excelDataIO.ToExcelWithExcelRowChildren();
                    foreach (DataColumn dc in dt.Columns)
                    {
                        key = dc.ColumnName.ToLower();
                        if (null == cellProperties[key]) continue;
                        //根据字段名称来设置 Excel 单元格值
                        cellProperties[key].cellValue = item[dc.ColumnName].ToString();
                    }
                    //逐行向 Excel 插入数据
                    excelDataIO.ToExcelWithData(cellProperties);
                }

                //获取 Excel 的字节数据
                byte[] data = excelDataIO.ToExcelGetBody();

                //释放资源
                excelDataIO.Dispose();
            }

            /// <summary>
            /// 从 Excel 获取数据
            /// </summary>
            public void FromExcel()
            {
                string fPath = "D:\\abc.xlsx";
                //采用 Aspose 组件
                IExcelDataIO excelDataIO = new ExcelDataIO(new AsposePlugin());
                ExcelSheet excelSheet = ExcelSheet.Instance;

                foreach (var kv in kvArr)
                {
                    //建立 Excel 表头名称与字段的映射关系
                    excelSheet.AddMapping(kv[1].Trim(), kv[0].Trim());
                }

                //获取指定行序号位置行数据键值对
                Dictionary<string, string> rowKV = excelDataIO.GetRowDataKayValue(excelSheet: excelSheet, fPath, 2);

                //获取指定行位置的行数据数组集合
                string[] rows = excelDataIO.GetRowData(fPath, 2);

                excelSheet.SheetName = "Sheet1";
                //获取指定 SheetName(也可通过指定 SheetIndex)和指定行序号位置的行数据数组合集合
                string[] rowDatas = excelDataIO.GetRowData(excelSheet: excelSheet, fPath, 2);

                //获取所有 WorkSheet 的名称
                string[] sheetNames = excelDataIO.GetWorksheetNames(fPath);

                //获取 Excel 的字节数据
                byte[] data = excelDataIO.ToExcelGetBody(fPath);

                //如果数据量少的情况，可直接获取 DataTable 集合
                //DataTable dt = excelDataIO.FromExcel(excelSheet, fPath, true);

                //如果数据量少的情况，可采用数据实体直接获取 List 集合
                //List<DataObj> list = excelDataIO.FromExcel<DataObj>(excelSheet, fPath);

                //大数据量的情况，可采用如下逐行获取 DataRow 数据
                //excelDataIO.FromExcel(excelSheet, fPath, dataRow =>
                //{
                //    string fn1 = dataRow["Name"].ToString();
                //    fn1 += "";
                //    Trace.WriteLine(fn1);
                //});

                //大数据量的情况，可采用数据实体逐行获取数据
                excelDataIO.FromExcel<DataObj>(excelSheet: excelSheet, fPath, dataObj =>
                {
                    Trace.WriteLine(dataObj.Name);
                });
            }

            class DataObj
            {
                public string ID { get; set; }
                public string Name { get; set; }
                public int Number { get; set; }
                public string Email { get; set; }
                public string PhoneNumber { get; set; }
                public string Address { get; set; }
            }

            private DataTable GetDbSource()
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("ID");
                dt.Columns.Add("Name");
                dt.Columns.Add("Number");
                dt.Columns.Add("Email");
                dt.Columns.Add("PhoneNumber");
                dt.Columns.Add("Address");

                DataRow dataRow = dt.NewRow();
                dataRow["ID"] = Guid.NewGuid().ToString();
                dataRow["Name"] = "ZhanSan";
                dataRow["Number"] = 32;
                dataRow["Email"] = "ZhanSan@hotmail.com";
                dataRow["PhoneNumber"] = "15302614521";
                dataRow["Address"] = "GuangDong ShenZhenShi";
                dt.Rows.Add(dataRow);

                dataRow = dt.NewRow();
                dataRow["ID"] = Guid.NewGuid().ToString();
                dataRow["Name"] = "LiShi";
                dataRow["Number"] = 28;
                dataRow["Email"] = "LiShi@hotmail.com";
                dataRow["PhoneNumber"] = "13366552176";
                dataRow["Address"] = "GuangDong ShenZhenShi NanShanQu";
                dt.Rows.Add(dataRow);

                dataRow = dt.NewRow();
                dataRow["ID"] = Guid.NewGuid().ToString();
                dataRow["Name"] = "WangWu";
                dataRow["Number"] = 30;
                dataRow["Email"] = "WangWu@hotmail.com";
                dataRow["PhoneNumber"] = "13265418715";
                dataRow["Address"] = "YunNan KunMing";
                dt.Rows.Add(dataRow);
                return dt;
            }
        }
    }
}
