# ExcelToScriptableObject
A tool that can convert the contents of an Excel table into multiple Unity ScriptableObject resource files according to certain rules.

If your project does not contain the ExcelDll file, please download the ExcelDll folder into your project (it is recommended to place it in the Editor directory).

Support for enumerated types (which need to be created in advance)

The format in the Excel table should be Enum.EnumName

Configuration format:

The first row is the field name

The second row is the type (supports int, float, bool, string, and enum types)

The third row is for comments and will not be read

视频教程：https://www.bilibili.com/video/BV1Nw4m1q7qt/

一个能够按一定规则将Excel表的内容转换成多个unity的scriptableobject资源文件的工具

如果您的项目里没有ExcelDll文件，请将ExcelDll文件夹下载到您的项目里(推荐放在Editor目录下) 

支持枚举类型(需要自己提前创建)

在Excel表的格式为Enum.枚举名称

配置格式为:

第一行为字段名

第二行为类型(支持int float bool string 枚举类型)

第三行作注释,不会被读取
