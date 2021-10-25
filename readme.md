
# 转表工具
## 目录结构
```
| -- Excels  存放待转换的策划配置表格(.xlsm)  
| -- Output   转表输出目录  
|    | -- DotB  二进制.b文件
|    | -- ServerCsv csv文件
|    | -- Struct    tars结构文件
| -- Tools    工具目录 
|    | -- Install   Windows下环境安装文件
|    | -- Server    功能实现文件
```
## 环境
* python 2.7
* xlrd 1.2.0
### 环境配置说明
为了方便Windows下配置，我们把环境相关安装文件放在Tools/Install下。  
双击`python-2.7.14.amd64.msi`安装python2.7，然后双击`xlrd_Install.bat`安装xlrd1.2.0
## 使用
1. 将配置表格放在Excels目录下  
2. Windows下双击TableGen.bat进行转换
3. Linux下使用./TableGen.sh进行转换