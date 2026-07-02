# DeepExcel Python执行功能测试用例

> 来源：用户常见Python自动化需求整理
> 日期：2026-06-29

---

## 一、基础计算

### 测试用例 PY-001：批量计算
- **用户需求**：`用Python计算A列每个数的平方，结果写入B列`
- **输入数据**：A1:A10 = {1,2,3,4,5,6,7,8,9,10}
- **预期Python**：
```python
import openpyxl
wb = openpyxl.load_workbook('current.xlsx')
ws = wb.active
for row in range(1, 11):
    ws.cell(row=row, column=2, value=ws.cell(row=row, column=1).value ** 2)
wb.save('current.xlsx')
```
- **验证结果**：B列显示平方值
- **验证点**：计算正确、写入正确

### 测试用例 PY-002：统计计算
- **用户需求**：`计算A列的标准差和方差`
- **预期Python**：
```python
import pandas as pd
df = pd.read_excel('current.xlsx')
std_val = df['A'].std()
var_val = df['A'].var()
# 返回结果
```
- **验证点**：统计值正确

---

## 二、数据清洗

### 测试用例 PY-010：pandas清洗
- **用户需求**：`用Python清洗数据，去掉空行，统一日期格式`
- **预期Python**：
```python
import pandas as pd
df = pd.read_excel('current.xlsx')
df = df.dropna(how='all')
df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
df.to_excel('cleaned.xlsx', index=False)
```
- **验证点**：空行删除、日期格式统一

### 测试用例 PY-011：批量替换
- **用户需求**：`用Python把所有"北京市"替换成"北京"`
- **预期Python**：
```python
import pandas as pd
df = pd.read_excel('current.xlsx')
df = df.replace('北京市', '北京')
df.to_excel('current.xlsx', index=False)
```
- **验证点**：替换正确

---

## 三、数据分析

### 测试用例 PY-020：分组汇总
- **用户需求**：`用Python按地区分组，计算每个地区的销售总额`
- **预期Python**：
```python
import pandas as pd
df = pd.read_excel('current.xlsx')
result = df.groupby('地区')['销售额'].sum().reset_index()
result.to_excel('summary.xlsx', index=False)
```
- **验证点**：分组正确、求和正确

### 测试用例 PY-021：数据透视
- **用户需求**：`用Python创建透视表，按地区和产品汇总`
- **预期Python**：
```python
import pandas as pd
df = pd.read_excel('current.xlsx')
pivot = pd.pivot_table(df, values='销售额', index='地区', columns='产品', aggfunc='sum')
pivot.to_excel('pivot.xlsx')
```
- **验证点**：透视表结构正确

---

## 四、文件操作

### 测试用例 PY-030：合并多个文件
- **用户需求**：`用Python把文件夹里所有Excel文件合并成一个`
- **预期Python**：
```python
import pandas as pd
import glob
files = glob.glob('data/*.xlsx')
df_list = [pd.read_excel(f) for f in files]
result = pd.concat(df_list, ignore_index=True)
result.to_excel('merged.xlsx', index=False)
```
- **验证点**：所有文件数据被合并

### 测试用例 PY-031：拆分工作表
- **用户需求**：`用Python按地区把数据拆分成多个文件`
- **预期Python**：
```python
import pandas as pd
df = pd.read_excel('current.xlsx')
for region in df['地区'].unique():
    region_df = df[df['地区'] == region]
    region_df.to_excel(f'{region}.xlsx', index=False)
```
- **验证点**：每个文件数据正确

---

## 五、可视化

### 测试用例 PY-040：生成图表
- **用户需求**：`用Python生成柱状图并保存为图片`
- **预期Python**：
```python
import pandas as pd
import matplotlib.pyplot as plt
df = pd.read_excel('current.xlsx')
plt.figure(figsize=(10, 6))
plt.bar(df['产品'], df['销量'])
plt.savefig('chart.png')
```
- **验证点**：图片生成成功、内容正确

### 测试用例 PY-041：多图表
- **用户需求**：`用Python生成包含4个子图的分析报告`
- **预期Python**：使用subplot创建2x2图表
- **验证点**：4个图表都正确生成

---

## 六、错误处理

### 测试用例 PY-050：文件不存在
- **用户需求**：`处理指定路径的文件`
- **预期行为**：文件不存在时返回友好错误信息
- **验证点**：不崩溃、错误信息清晰

### 测试用例 PY-051：数据格式错误
- **用户需求**：`处理可能包含非数字的数据`
- **预期行为**：跳过或处理异常值
- **验证点**：程序不崩溃

---

## 注意事项

1. **Python环境**：需要用户系统安装Python
2. **依赖库**：pandas, openpyxl等需要预装或自动安装
3. **文件路径**：使用相对路径或临时路径
4. **安全限制**：禁止执行危险操作（如os.system, subprocess）
5. **执行超时**：设置合理的超时时间
