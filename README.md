# 学生随机分组系统

读取Excel表格中的学生信息，对学生随机分组后输出到Excel表格中。每组5～6人，且5人组不超过5组。每组至少有一名想当组长的学生和一名女生。特别地，如果有想当组长的女生，其余组员也可都为男生

## 注意事项

* **仅支持JDK9及以下版本，使用JDK10运行会造成输入文件的损坏**

## Excel文件格式

* 输入格式为.xls和.xlsx皆可，输出格式最好为.xlsx

* 参见[test.xlsx](./test.xlsx)

## 用法

1. 将要读取的Excel表格与放在Grouping文件夹下。

2. 打开命令行，设置当前目录为Grouping工程文件夹，输入：

```terminal
java Grouping [读取文件名.xlsx] [输出文件名.xlsx]
```

例如

```terminal
java Grouping test.xlsx result.xlsx
```
