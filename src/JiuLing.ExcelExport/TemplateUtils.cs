using System.Text.RegularExpressions;
using JiuLing.ExcelExport.Items;

namespace JiuLing.ExcelExport
{
    internal class TemplateUtils
    {
        /// <summary>
        /// 获取单元格绑定信息
        /// </summary>
        /// <param name="cellValue">单元格配置的值</param>
        /// <returns></returns>
        public static CellBindingInfo GetCellBindingInfo(string cellValue)
        {
            //匹配规则
            //列表绑定  %表名-列名-list%
            //值绑定    %表名-列名-0%

            MatchCollection mc = Regex.Matches(cellValue, "{(?<tableName>.*)}-{(?<columnName>.*)}-list");
            if (mc.Count == 1)
            {
                return new CellBindingInfo()
                {
                    BindingType = BindingTypeEnum.List,
                    TableName = mc[0].Groups["tableName"].Value,
                    ColumnName = mc[0].Groups["columnName"].Value
                };
            }
            mc = Regex.Matches(cellValue, "{(?<tableName>.*)}-{(?<columnName>.*)}-0");
            if (mc.Count == 1)
            {
                return new CellBindingInfo()
                {
                    BindingType = BindingTypeEnum.Cell,
                    TableName = mc[0].Groups["tableName"].Value,
                    ColumnName = mc[0].Groups["columnName"].Value
                };
            }

            return new CellBindingInfo()
            {
                BindingType = BindingTypeEnum.NotTemplate,
            };
        }
        /// <summary>
        /// 获取单元格要绑定的列名
        /// </summary>
        /// <param name="cellValue">单元格配置的值</param>
        /// <param name="tableName">单元格配置表名</param>
        /// <returns></returns>
        public static string GetCellBindingColumnName(string cellValue, string tableName)
        {
            MatchCollection mc = Regex.Matches(cellValue, $"{{{tableName}}}-{{(?<columnName>.*)}}-list");
            if (mc.Count != 1)
            {
                return "";
            }
            return mc[0].Groups["columnName"].Value;
        }

    }
}
