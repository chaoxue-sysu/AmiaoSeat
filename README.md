# 阿喵排排座 （AmiaoSeat）
**阿喵排排座** 是用来进行考场随机排座位的程序，目前（V1.0）主要适用于鸭大北校区考场座位安排，帮您一键排好考场座位。

## 使用方法
### 安装
Windows用户可直接下载[安装包](Setup.exe)，进行安装使用。

其他系统用户可以安装Python环境（Python>3.6）和pip,并安装依赖包requirements.txt，执行main.py程序。
```
$ pip install -r requirements.txt
$ python main.py
```
### 使用
输入多个.xlsx格式的座位和学生名单表格文件，选择输入文件夹，点击排座位按钮即可排好座位。输入的表格必须满足一定格式，第一个Sheet为教室座位表，第二个Sheet为学生名单表，参见 [example/seat_example.xlsx](example/seat_example.xlsx) 。排好座位后，会在座位表Sheet的座位号下填上学生名字，在学生名单Sheet的座位号列填上对应的座位号，参见 [example/seat_example.amiu.xlsx](example/seat_example.amiu.xlsx)。注意排座位是随机的，每次执行结果都可能不相同；另外学生名单中不能存在重名情况。


## 软件开发过程记录（附）
### 软件概述
本软件基于Python实现，用PyInstaller打包Python环境成exe程序，并用MISI集成Windows安装包形式。
### PyInstaller 打包 Python 程序
1. 安装 PyInstaller:
    ```
    $ pip install pyinstaller
    ```
2. 打包 main.py（参考 [pyinstaller文档](https://pyinstaller.readthedocs.io/en/stable/usage.html)）: 
    ```
    $ pyinstaller -w -i logo.ico --add-binary logo.ico;. main.py
    ```
3. 打包后会在生成disk/main目录，在此目录下有`main.exe`程序，点击即可运行。

### MISI打包成Windows安装包
为了更专业发布和管理软件，可以把上述disk/main目录的内容打包到安装包进行发布。参考：[windows下安装包制作软件：NSIS的使用方法](https://blog.csdn.net/zhichaosong/article/details/106275151)。最终生成的安装包为 [Setup.exe](Setup.exe) 。

**！注意：** 一定要安装 NSIS 和 HM NIS Edit 两个软件，才能完成打包。





