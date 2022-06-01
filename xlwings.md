### 1.xlwings工作表操作

#### 1.新建或打开工作簿

``````python
import xlwings as xw

app = xw.App(visible=True, add_book=True)
wb = xw.Book(r'C:\Users\Se7en\Desktop\深圳市高郭氏精密机械有限公司_综合所得申报_202112.xls')

``````

#### 2.新建工作表（批量创建工作表）

``````python
# 在最后位置插入工作表
sht = wb.sheets.add(after=wb.sheets.count)

# 在第5个工作表后面插入新的工作表
sht = wb.sheets.add(after=5)

# 命名新建的工作表
sht.name = 'cash'
``````

``````python
for i range(1,13):
    sht = wb.sheets.add(after = wb.sheets.count)
    sht.name=str(i) + '月'
``````



#### 3.引用工作表

``````python
# 索引引用
wb.sheets[1].name

# 名称引用
wb.sheets['汇总']name
``````

#### 4.遍历工作表

``````python
sList = []
for sht in wb.sheets:
    sList.append(sht.name)
    
for i in range(wb.sheets.count):
    print(wb.sheets[i].name)
``````

#### 5.复制工作表

``````python
wb.sheets['cash'].copy()
``````

#### 6.删除工作表（批量删除工作表）

``````python
wb.sheets['cash(2)'].delete()
``````

``````python
for sht in wb.sheets:
    if '月' in sht.name:
        sht.delete()
``````

#### 7.拆分工作表为独立的工作表

``````python
for sht in wb.sheets:
    sht.api.Copy()
    wb.books[xw.books.count-1].save(r'' + sht.name + '.xlsx')
    xw.books[xw.books.count-1].close()
``````



#### 8.def getSheetName

``````python
def getSheetName(ws):
    tList = []
    for s in ws:
        sList.append(s.name)
        return tList

for a in getSheetName(wb.sheets):
    print(a)
``````

### 2.xlwings单元格操作

``````python
sht = wb.sheets['Sheet1']
# 读取单元格的值
sht.range('a1').value

# 给单元格赋值
sht.range('b1').api.Value = 'xlwings'
``````

#### 1.读取连续的单元格区域

``````python
arr = sht.range('a1').expand().value

# 获取连续单元格区域的行数
sht.range('a1').expand().rows.count
# 获取连续单元格区域的列数
sht.range('a1').expand().columns.count
``````



#### 2.读取单元格区域

``````python
sht.range('A65536').end('up').row
sht.range('A1').end('down').row

sht.range('a65536').api.End(-4162).Row
sht.range('a1').api.End(-4121).Row
``````

``````python
arr = sht.range('a1:c' + str(sht.range('a65536').end('up').row)).value
``````



#### 3.给单元格区域赋值

``````python
sht = wb.sheets[0]
rng = sht.range('e1')

rng.value=[1,2,3]           #横向赋值
rng.value=[[1], [2], [3]]   #纵向赋值
rng.value = [[1,2,3], [4,5,6], [7,,9]] #三行三列赋值
``````

#### 4.清空单元格区域

``````python
rng.resize(3,3).clear_contents()
``````

#### 5.下拉列表

``````python
sht = wb.sheets[0]
rng = sht.range('a1')
rng.api.Validation.Add(3,1,1,'1,2,3,4')
``````

#### 6.超链接

``````python
rng.add_hyperlink(address='www.baidu.com', text_to_display='百度', screen_tip=None)
``````



#### 7.demo

``````python
sht = wb.sheets.add(before=1)
sht.name = '目录'
sht.range('a1').value = '目录'

sList = []
for s in wb.sheets:
    if s.name != '目录':
        sList.append([s.name])
sht.range('a2').value = sList
``````

