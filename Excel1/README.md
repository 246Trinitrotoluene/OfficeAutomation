## Excel1

### 功能需求分析
#### Python 处理部分
1. (定期)从pixiv获取小说数据（小说名称、发布时间、**阅读量、收藏量**）
2. 将数据格式化并保存至文件

#### Excel VBA 处理部分

3. 将数据导入excel表格中
4. 使用Excel进行简单的数据分析


### 实现过程
1. 使用 pypixiv 获取小说数据
2. 使用 pandas 格式化数据，并导出xlsx
3. 使用 pywin32 打开Excel程序
4. 通过 Excel 函数引用，导入数据
5. 运行 VBA 代码，将获取的数据存放在数据源表中（【点】【赞】两个sheet）
6. 运行 VBA 代码，重新筛选【近期】sheet的数据
7. 运行 VBA 代码，更新【个人】sheet的数据
8. 使用 pywin32 将xlsm另存为xlsx
9. ~~使用 WPS Office，分享 xlsx文件~~
