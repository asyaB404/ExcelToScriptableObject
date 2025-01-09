一个能够按一定规则将Excel表的内容转换成多个unity的scriptableobject资源文件的工具
如果您的项目里没有ExcelDll文件，请将ExcelDll文件夹下载到您的项目里

在Execl表中的配置格式为:
第一行为字段名
第二行为类型(支持int float bool string 枚举类型 Sprite)
第三行作注释,不会被读取

枚举类型(需要自己提前创建)
格式: Enum.{枚举名称}
例如: Enum.MyEnum    //这在对应生成sobj类型的字段为 public MyEnum {Excel表中的字段名}; 如果是其他命名空间的话需要自己导

Sprite类型
格式：Assets\{路径}
例如：Assets\AddressableAssets\GameRes\Arts\Textures\Items\山梅.png 

视频教程：https://www.bilibili.com/video/BV1Nw4m1q7qt/
博客链接：https://www.cnblogs.com/asyaB404/p/18229482
