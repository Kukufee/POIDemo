demo描述，及使用生成POI
Excle测试：
通过读取模板
测试类：com.my.demo.excel.Client
PPT测试：
使用是通过读取PPT模板来填充数据，可填充text、table、chart（piechart和barchart）
图表生成：暂时只写了两种BarChart（柱状图）和PieChart(饼状图)
测试类：
通过下标填充PPT：com.my.ppt.chart.client.ClientFillChartByIndex
获取当页图表下标（注意先为chart填写title，获取下标后可以删除title）：com.my.ppt.chart.client.ClientGetTempChartIndexAndTitle

生成柱状图及饼状图的过程都一样：参考com.my.ppt.chart.PieChartFillUtil
ppt-part（转成chart）——创建excel——series——seriesName填充、seriesCat（图例名）填充、seriesVal（图例对应的数字）填充
——为上述三者建立与excel的连接——输出ppt

注意点：
使用注意：ppt以页为单位填写，包括text、table、chart
		相关数据内容特别是下标一定要与Temp保持一致，不然容易出错
柱状图
1.填写excel时需要使用double数据，不然ppt点开excel是文本类型，数据会销失
2.模板系列必须数量>=填写数量，否则会出异常
3.模板图例数量可以与填写数量不一致，不会影响数据标签显示
4.char可以根据名字获取（名字需要唯一）或者下标获取
饼状图：
1.几个数据很小的时候标签会重叠，解决方法：将数据和图裂一起显示，或者将标签设置为自动匹配，优先考虑第一种

设计：每个图表数据拥有一个类，实现共同的接口，PieChartData、BarChartData implements ChartData
	 ppt以页为单位填写
	 chart填写可以通过Title填写（PieChartFillUtil、BarChartFillUtill）或者根据下标填写（com.my.ppt.chart.ChartFillByIndexUtil）
	 如果使用下标填写需要使用com.my.ppt.chart.ChartIndexPrintUtil，工具打印图表的下标，图表必须先创建表头，不用再删掉
	