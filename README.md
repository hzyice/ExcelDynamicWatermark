# ExcelDynamicWatermark
Excel 动态水印
overview:
  开发时有个需求，导出excel表格需要动态添加水印，所谓动态---即整个sheet都铺满水印。最开始的思路是，在每次需要创建sheet时，生成动态水印。做到了。找到的
  jar包在window上运行正常（开发环境是在window下），一部署到时Linux项目启动就挂掉。百度了下，jar不支持Linux系统。serach了半天依然无解。转换思路。选
  生成铺满整个sheet的样本，然后每次需要创建sheet时，读取样本的IO作为sheet,再填充表格内容 .结果: O了个K 成功!特此push到github上，以作备份。
