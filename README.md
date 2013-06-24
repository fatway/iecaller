# 网络学习页面定时激活

使用 IHTMLDocument2 对象模拟网页内部链接的点击，保持用户的登陆状态

基于 C#.Net 2.0 实现

---

主要信息

	// 访问地址
	string url = "http://s.gddec.net/icourse/index/doStudyIndex.do";

	// 页面文档
	IHTMLDocument2 doc = null;

	// 循环点击列表
	string[] todoList = new string[] {
	        "课程学习",
	        "通知公告",
	        "小纸条",
	        "互助交流",
	        "学习笔记",
	        "我的班级",
	        "成绩统计",
	        "课程资料",
	        "助学评价"
	    };
	// 当前点击项
	static int listPn = 0;


---
Love my Calin & Boyang