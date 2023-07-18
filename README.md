# HZTIKU
该工具用于对于特定格式题库word文件进行xlsx文件格式转换
## 安装

要使用此工具，您需要在系统上安装 Python 3 以及以下软件包：

- docx
- openpyxl
- re

您可以通过运行以下命令来安装这些软件包
```
pip install -r requirements.txt
```

git clone本git

```
git clone https://github.com/grandjohnson/HZTIKU.git
```
升级到新版
```
cd HZTIKU
git pull
pip install -r requirements.txt
```
## 输入格式
输入文件中，题目格式应参照如下格式：

### 单选题
1、L1现场处置方案至少每（C）演练一次。
A、星期	B、月	C、半年	D、年

### 多选题
1、L2受限空间作业场所负责人的职责包括（ABC）
A、负责审批受限空间作业风险分析；
B、负责审批受限空间进入许可证；
C、担任应急情况下的总指挥；
D、决定什么时间可以进入受限空间。

### 判断题
1、L0生产经营单位必须执行依法制定的保障安全生产的国家标准或者行业标准。（对）

### 系统
##低压输送系统##

## 输出格式
输出的xlsl文件表头包含以下内容：试题难度、试题类型、试题内容、题目选项A、题目选项B、题目选项C、题目选项D、题目选项E、题目选项F
