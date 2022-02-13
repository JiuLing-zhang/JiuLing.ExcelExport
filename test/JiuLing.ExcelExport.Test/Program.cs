using System;
using System.Data;
using System.IO;
using System.Linq;

namespace JiuLing.ExcelExport.Test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var ds = new DataSet();
                ds.Tables.Add(GetTable1());
                ds.Tables.Add(GetTable2());
                ds.Tables.Add(GetTableTitle());

                string s = DataTableToString(ds.Tables[0]);
                Console.WriteLine(ds.Tables[0].TableName);
                Console.WriteLine(s);
                Console.WriteLine();

                s = DataTableToString(ds.Tables[1]);
                Console.WriteLine(ds.Tables[1].TableName);
                Console.WriteLine(s);
                Console.WriteLine();

                s = DataTableToString(ds.Tables[2]);
                Console.WriteLine(ds.Tables[2].TableName);
                Console.WriteLine(s);
                Console.WriteLine();

                var templateFile = Path.Combine(System.AppContext.BaseDirectory, "Template.xlsx");
                var destinationFile = Path.Combine(System.AppContext.BaseDirectory, "test.xlsx");

                new TemplateData().Export(templateFile, destinationFile, ds);
                Console.WriteLine("导出完成");
                Console.WriteLine($"模板文件：{templateFile}");
                Console.WriteLine($"导出文件：{destinationFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"导出失败：{ex.Message}");
            }
            finally
            {
                Console.ReadLine();
            }
        }

        private static DataTable GetTable1()
        {
            var dt = new DataTable("dt1");
            dt.Columns.Add("Class");
            dt.Columns.Add("Name");
            dt.Columns.Add("Score");

            var dr = dt.NewRow();
            dr["Class"] = "1班";
            dr["Name"] = "张三";
            dr["Score"] = "90";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Class"] = "2班";
            dr["Name"] = "李四";
            dr["Score"] = "80";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Class"] = "2班";
            dr["Name"] = "王五";
            dr["Score"] = "72";
            dt.Rows.Add(dr);
            return dt;
        }

        private static DataTable GetTable2()
        {
            var dt = new DataTable("dt2");
            dt.Columns.Add("Time");
            dt.Columns.Add("Type");

            var dr = dt.NewRow();
            dr["Time"] = "早上";
            dr["Type"] = "语文";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Time"] = "中午";
            dr["Type"] = "英语";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Time"] = "下午";
            dr["Type"] = "数学";
            dt.Rows.Add(dr);
            return dt;
        }

        private static DataTable GetTableTitle()
        {
            var dt = new DataTable("dtOther");
            dt.Columns.Add("Name");
            dt.Columns.Add("Time");

            var dr = dt.NewRow();
            dr["Name"] = "课程表";
            dr["Time"] = DateTime.Now;
            dt.Rows.Add(dr);

            return dt;
        }


        private static string DataTableToString(DataTable dt)
        {
            return string.Join(Environment.NewLine, dt.Rows.OfType<DataRow>().Select(x => string.Join(" ; ", x.ItemArray)));
        }
    }
}
