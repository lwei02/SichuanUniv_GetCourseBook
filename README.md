# SichuanUniv_GetCourseBook
获取四川大学本人任意学期的课程教材目录

## 使用方法
1. 确认计算机中安装有Google Chrome，若没有安装，请先[点此下载](https://www.google.com/chrome/thank-you.html?statcb=1&installdataindex=empty&defaultbrowser=0&standalone=1)（若左侧链接无法访问，[可尝试此链接](https://www.google.cn/chrome/thank-you.html?statcb=1&installdataindex=empty&defaultbrowser=0&standalone=1)）
2. 确认计算机中安装有Python3，若没有安装，请[点此下载](https://www.python.org/downloads/)
3. 下载仓库，解压后，在仓库目录中运行`python3 -m pip install -r requirements.txt`
4. 运行`python3 getChromeLib.py`，程序会自动下载对应版本的Chromedriver，或者你也可以[手动下载](https://chromedriver.storage.googleapis.com/index.html)对应版本的Chromedriver并解压后放置在`./lib/chromedriver.exe`处
5. 修改`getBooks.py`，填入你的学工号和统一身份认证密码（不是微服务密码，不是教务系统密码），同时可以选择填入[百度云OCR](https://cloud.baidu.com/product/ocr.html)的`API_KEY`和`SECRET_KEY`，如果不填百度云的认证信息，则会在过程中弹出一个验证码截图，可以通过手动识别填写该验证码
6. 运行`python3 getBooks.py`，程序会自动抓取你的课程表，选择对应学期，程序将逐个查询该学期的课程所有的教材，并保存为xlsx文档

## 原理
* 教师教务系统本身不允许学生登录，但是统一身份认证的SSO登录接口并不对学生、教师作区分，导致学生账号可以简单的通过SSO登录进教师教务系统
* 教师教务系统存在一个认证通过后不区分学生教师，也不确认所登录账号是否是负责教师的接口：
  ```
  http://zhjwjs.scu.edu.cn/teacher/comprehensiveQuery/search/textbookSpecified/show?jsh=&kxh=<课序号>&kch=<课程号>&zxjxjhh=<学期代码>
  ```
  该接口可以查询某个课程所使用的教材，通过对其返回的HTML进行分析，可以得到结构化的课程教材书目
  
## 其他学校
该应用应当可以简单修改以适配其他采用新版URP教务系统的学校（需要能够统一身份认证，且统一认证后能够SSO登录到教师系统中），欢迎fork

## 关于Issues
如果程序中出现了运行错误，请提Issue，不接受其他方面的Issue（如请求新功能等）

## 关于PR
欢迎PR
