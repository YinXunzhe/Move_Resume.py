#移动简历
=======
##程序用途：根据excel筛选结果，从共享盘批量剪切候选人简历至工作组  
-------
  S1：读取excel表中筛选出来的候选人姓名  
  S2：根据姓名在简历文件夹中搜索并剪切到目标路径  

##前提条件：  
---------
已安装pycharm、python3.8，对文件夹有剪切权限（部分同事无法剪切共享盘文件）  

##使用步骤：
----------
	1. 把Move_Resume工程下载并解压，使用pycharm打开该工程  
	2. 把当天的excel简历信息库下载到工程文件夹下 
	3. 在信息库中根据院校和专业等信息在表格中手动筛选出想要的候选人  
	4. 将筛选结果保存到新的sheet表  
		a. 在信息库中新建一个Sheet表，并将该sheet命名为“待提取候选人”（不可更改名称）  
		b. 鼠标点击一个数据中的单元格，然后Ctrl+A全选数据  
		c. 使用Alt + ; 键来仅选中筛出来的可视单元格  
		d. Ctrl +C复制数据，在新的sheet表中选中最左上角单元格，Ctrl+V粘贴  
		e. Ctrl+S保存对excel的更改  
	5. 根据实际情况和程序注释，更改程序中的文件和路径名  
		#excel简历信息库的文件名，需根据当天的文件名更改  
		exl_file='副本附件3：【总表】xx校招信息库x.xlsx'  
		#需要检索的简历存放路径，更改路径时注意不要误删路径两头的单引号'和最前面的字母r  
		path=r'D:\Project\技术分享\MoveResume\xx春季校园招聘简历推送'  
		#筛选后的简历存放路径，更改路径时注意不要误删路径两头的单引号'和最前面的字母r  
		dst_path=r'D:\Project\技术分享\MoveResume\筛后简历'  
	6. 运行  
		成功案例：  
			已有 3 份简历被您从  
			 D:\Project\技术分享\MoveResume\xx春季校园招聘简历推送-20210412   
			 剪切至  
			 D:\Project\技术分享\MoveResume\筛后简历  
			 Congratulations!  