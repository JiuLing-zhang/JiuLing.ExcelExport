![](https://img.shields.io/github/license/JiuLing-zhang/JiuLing.ExcelExport)
![](https://img.shields.io/github/workflow/status/JiuLing-zhang/JiuLing.ExcelExport/Build)
[![](https://img.shields.io/nuget/v/JiuLing.ExcelExport)](https://www.nuget.org/packages/JiuLing.ExcelExport/)
[![](https://img.shields.io/github/v/release/JiuLing-zhang/JiuLing.ExcelExport)](https://github.com/JiuLing-zhang/JiuLing.ExcelExport/releases)  

## JiuLing.ExcelExport
.Net 5开发的一个基于Excel模板导出的组件（基于NPOI），配置完成模板和数据源即可一键导出，支持多 `Sheet`导出。  

## 安装  
* ~~通过`Nuget`直接安装。👉👉👉[`JiuLing.ExcelExport`](https://www.nuget.org/packages/JiuLing.ExcelExport)~~  
* ~~下载最新的`Release`版本自己引用到项目。👉👉👉[`下载`](https://github.com/JiuLing-zhang/JiuLing.ExcelExport/releases)~~  
* 开发中  

## 使用  
1. 将要导出的数据保存为 `DataSet` 对象。

    ```C# 
    var ds = new DataSet();
    //添加待导出数据
    ds.Tables.Add(GetTable1());
    ds.Tables.Add(GetTable2());
    ```

2. 配置Excel模板。  
* 列表形式的绑定：  
将单元格配置为如下格式：**%表名-字段名-list%**。  
例如： `%dt1-Class-list%`  
该配置会自动查找 `DataSet` 中的 `dt1` 表，并且将 `Class` 列绑定到 `Excel` 的当前列。  

* 单元格形式的绑定：  
将单元格配置为如下格式：**%表名-字段名-0%**。  
例如： `%dtOther-Name-0%`  
该配置会自动查找 `DataSet` 中的 `dtOther` 表，并且将 `Name` 列的第一行对应的值绑定到 `Excel` 的当前单元格。  

3. 导出
    ```C#
    //templateFile：模板文件的文件名
    //destinationFile：要导出的文件名
    //ds：数据源
    var templateFile = Path.Combine(AppContext.BaseDirectory, "Template.xlsx");
    var destinationFile = Path.Combine(AppContext.BaseDirectory, "test.xlsx");
    var ds = new DataSet();
    new TemplateData().Export(templateFile, destinationFile, ds);
    ```

## 已知问题  
1. 列表绑定时，如果模板中对应的部分包含合并单元格，导出后的列表不会自动合并单元格。  
2. 由于 `NPOI` 对时间的格式支持的不是很友好，因此如果导出的字段为 `DateTime` 类型，则会直接转换成 `String` 类型填充，使用 `"yyyy-MM-dd HH:mm:ss"` 进行格式化。  
## License
MIT License