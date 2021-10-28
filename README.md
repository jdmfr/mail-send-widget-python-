## 邮件批量私发的简单小脚本(Python)

依赖：

Pyhton3

xlrd库（用于读取excel，最新2.0.x版本不支持xlsx格式，推荐使用xls格式）

zmail库： [zmail/README-cn.md at master · zhangyunhao116/zmail · GitHub](https://github.com/zhangyunhao116/zmail)

 `pip install zmail` or `pip3 install zmail`

使用方法： 

将成绩信息按照如下形式整理在excel中，每人一行（姓名、成绩排名、成绩排名、学号，**推荐保存为xls**）。由于个人邮箱是 "学号@固定尾缀" ，通过学号可以自动得到邮箱地址。

![excel,保存为xls](material\1.png)

运行py脚本，依次输入邮箱地址、邮箱密码、excel表格的存储路径，即可。

![2](material\2.png)

随后等待发送完毕即可。

![3](material\3.png)

当然，你可以根据自己的喜好，任意修改文本内容。

Enjoy your leisure time