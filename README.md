# T4 for openlab exam

根据Openlab纳新复试建立的仓库之一，对应于T4.

## 仓库结构

src文件夹里是代码和可执行文件,rec文件夹内是使用说明

## T4.1

### 简单的对项目的介绍：  

是否想要打好校园信息战？是否想要第一时间获得学校的通知内容？来看看本项目吧.

### 概况

#### 本项目使用Python脚本完成对山东大学三个网站通知的爬取：

1.山大视点-山大要闻

2.本科生院-工作通知

3.计算机学院-本科教育

#### 同时带有控制功能：

4.在已定时间自动进行三个网站内容的爬取

5.立刻停止，同时格式化输出表格

6.删除原有表格，新建输出表格

### 输入与输出：

1-6作为操作项选择，同时可以控制爬取的数量；爬取的结果将输出为ret.xlsx文件并作一定的格式化处理，位置就在当前目录下


### 关于自动爬取：

选择4模式后，程序将在7、12、18、22时间自动爬取三个网站的各一页内容并输出，同时可Ctrl+C在下一次爬取时退出，保证你得到最新的校园新闻！

### 其他

1.本脚本每次爬取后会覆盖之前的爬取结果

2.输出结果在当前目录下，推荐使用Excel打开查看
