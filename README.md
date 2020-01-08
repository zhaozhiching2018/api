# api
本文章主要查询小程序每天的用户访问情况
属性	类型	说明
page_path	string	页面路径
page_visit_pv	number	访问次数
page_visit_uv	number	访问人数
page_staytime_pv	number	次均停留时长
entrypage_pv	number	进入页次数
exitpage_pv	number	退出页次数
page_share_pv	number	转发次数
page_share_uv	number	转发人数


定时发送邮件:
注意:要下载pywin32库
其中autho = '123456a'  # 授权码
    smtpServer = 'smtp.163.com'  # smtp服务器 可以从邮箱设置中获得
