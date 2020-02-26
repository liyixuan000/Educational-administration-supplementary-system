<%
dim siteName,logourl,indexpic,diarypic,indextxt,indextxtcenter,indextxtcolor,indextxtsize,siteUrl,email,sitereg,sitePath,databaseType,databaseServer,databaseName,databaseUser,databasepwd,accessFilePath,templateFileFolder,defaultTemplate,sitePictype,sitePicsize,siteSofttype,siteSoftsize,isCut,siteNewspicsize,sitePhotopicsize,siteCasepicsize,siteDownpicsize,siteProductpicsize,runMode,guestmode,siteVisiteJs,siteKeyWords,siteNotice,siteDes,copyRight

siteName="教务辅助管理信息系统活动管理 "'系统名称
logourl="0" '是否开启客户端 1为开启，0为不开启
indexpic="images/bg/" '首页图片（路径）,跟换图片直接传到指定位置即可，路径必须以“/”结束 图片大小为：580*320
diarypic="/images/a_1.jpg"
indextxt="<p>青春无畏，热爱无限</p>" '首页文字：
indextxtcenter="right" '首页文字对齐方式：
indextxtcolor="#999999" '首页文字颜色：
indextxtsize="14px" '首页文字大小：
siteUrl="127.0.0.0" '注册url：
email="2236370019@qq.com" '注册email：
sitereg="" '授权码：
sitePath=""
databaseType=0
databaseServer="(local)\gsql"
databaseName="eptimehomev5"
databaseUser="sa"
databasepwd="123123"
accessFilePath="/data/#eptime#%Enterprise.mdb" 'ACCESS数据库路径：
templateFileFolder="Winter" '首页动画 ""-无动画;Fish-游动的鱼儿;Spring-春天（花）;Summer-夏天（蝴蝶）;Autumn-秋天（落叶）;Winter-冬天（雪花）
defaultTemplate="blue" '系统色彩方案 blue-蓝色之恋;green-青青世界;red-红色爱情;purple-紫气东来;black-黑色诱惑
sitePictype=""
sitePicsize=""
siteSofttype=""
siteSoftsize=""
isCut="1"  '是否启用审核功能：1-启用；0-不启用
siteNewspicsize=""
sitePhotopicsize=""
siteCasepicsize=""
siteDownpicsize=""
siteProductpicsize=""
runMode="1" '登陆页消费饼状图：""-按照支出大类显示;"1"-按照支出小类显示
guestmode=""
siteVisiteJs=""
siteKeyWords=""
siteNotice=""
'苹果手机添加到主屏幕显示名称，六字以内
siteDes="活动管理"
copyRight="" '版权信息：
%>