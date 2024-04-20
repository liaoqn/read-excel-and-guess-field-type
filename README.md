# read-excel-and-guess-field-type

通常读取 Excel 是已知列类型的，比如使用 EasyExcel 读取会事先定义 Excel 文件对应的实体，Excel 列类型和实体字段类型对应。

但也会有某些场景，事先并不知道 Excel 文件的列类型。比如一些 BI 应用，将 Excel 文件导入数仓进行数据分析，会根据 Excel 的列类型在数仓创建表。对于这类场景事先并不知道列类型，无法直接定义 Excel
对应的实体，并且读取 Excel 时需要推测字段的类型。

本项目主要采用 EasyExcel 解析 Excel，并推测 Excel 列的类型，最后读取时将结果转换为推测得到的列类型。

使用方式
=====
1、读取 Excel sheet 页的列信息，并推测字段类型

	InputStream excelInputStream = new FileInputStream("excel file path");

    SheetColumnData sheetColumnData = ExcelReadUtil.readColumnMetaData(excelInputStream, FileType.XLSX, 0, 10);

列类型结果如下：

    {
        "columnDataList": [
            {
                "columnIndex": 0,
                "fieldType": "STRING",
                "title": "文本"
            },
            {
                "columnIndex": 1,
                "fieldType": "INTEGER",
                "title": "数字"
            },
            {
                "columnIndex": 2,
                "fieldType": "DECIMAL",
                "title": "小数"
            },
            {
                "columnIndex": 3,
                "fieldType": "DATE",
                "title": "日期"
            },
            {
                "columnIndex": 4,
                "fieldType": "DATE",
                "title": "日期时间"
            },
            {
                "columnIndex": 5,
                "fieldType": "STRING",
                "title": "时间"
            },
            {
                "columnIndex": 6,
                "fieldType": "BOOLEAN",
                "title": "布尔"
            }
        ],
        "sheetNo": 0
    }

2、读取Excel，读取时传入上一步得到的列类型

    InputStream excelInputStream = new FileInputStream("excel file path");
    
    ExcelReadContext readContext = new ExcelReadContext();
    readContext.setInputStream(excelInputStream);
    readContext.setFileType(FileType.XLSX);
    readContext.setSheetNo(0);
    readContext.setColumnDataList(sheetColumnData.getColumnDataList());
    readContext.setOffset(1);
    readContext.setSize(20);

    List<ArrayList<Object>> dataList = ExcelReadUtil.readSheet(readContext);

读取得到的 Excel 数据，与传入的 columnDataList 顺序保持一致，并按照 columnDataList 传入的列类型转换数据类型，类型转换失败取null。

支持的类型
=====

1. 整数（Integer）。推测的数据全部是整数，并且未超过 Integer 的最大值。
2. 长整数（Long）。推测的数据全部是整数，有超过 Integer 的最大值，并且未超过 Long 的最大值。
3. 小数（BigDecimal）。推测的数据都是数字，并且有数据不符合整数和长整数的条件。
4. 布尔（Boolean）。"false"和"否" 为 false，"true"和"是" 为 true。
5. 日期（Date）。Excel 自身列类型是日期，或者数据符合日期格式。支持常见日期格式，包括中文日期格式。
6. 文本（String）。非上述任一类型，并且推测的数据长度皆不超过 4000 个字符。
7. 大文本（String）。有推测的数据长度超过 4000 个字符。
