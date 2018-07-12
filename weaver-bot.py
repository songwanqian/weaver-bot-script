#!/usr/bin/python
# -*- coding: UTF-8 -*-
"""
    抓取OA昨天的考勤记录并邮件通知
    系统次日早上05:05更新考勤记录
"""
import cookielib
import re
import smtplib
import sys
import urllib
import urllib2
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

from lxml import etree

# 配置信息
oa_username = 'xxx'  # OA的用户名
oa_password = 'xxxxxx'  # OA的密码
mail_from_addr = 'xxx@188.com'  # 发送邮箱
mail_from_pass = 'xxxxxx'  # 发送邮箱密码
mail_to_addr = 'xxx@163.com'  # 接收邮箱
mail_smtp_server = 'smtp.188.com'
mail_smtp_port = 587  # 或25
mail_smtp_ssl = True  # 是否加密传输


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr.encode('utf-8') if isinstance(addr, unicode) else addr))


if __name__ == '__main__':
    # 1.登录系统
    filename = 'cookie.txt'  # cookie文件
    cookie = cookielib.MozillaCookieJar(filename)
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookie))
    loginData = urllib.urlencode({
        'loginfile': '/wui/theme/ecology8/page/login.jsp?templateId=3&logintype=1&gopage=',
        'logintype': '1',
        'fontName': '微软雅黑',
        'formmethod': 'get',
        'isie': 'false',
        'islanguid': '7',
        'loginid': oa_username,
        'userpassword': oa_password,
    })
    loginUrl = 'http://oa.xxxx.cn/login/VerifyLogin.jsp'
    result = opener.open(loginUrl, loginData)
    cookie.save(ignore_discard=True, ignore_expires=True)

    # 2.获得签到表的Hash，查询需要
    tableViewUrl = 'http://oa.xxxx.cn/formmode/search/CustomSearchBySimpleIframe.jsp?' \
                   'customid=25&e71494045129565=&mainid=0'
    result = opener.open(tableViewUrl)
    searchObj = re.search(r"__tableStringKey__='([0-9A-F]+)'", result.read(), re.M)
    if not searchObj or len(searchObj.group(1)) != 32:
        print "TableHash Not found!!"
        sys.exit()
    tableHash = searchObj.group(1)

    # 3.查询Ajax接口
    checkUrl = 'http://oa.xxxx.cn/weaver/weaver.common.util.taglib.SplitPageXmlServlet?' \
               '&orderBy=t1.indate%20desc,t1.intime%20asc,t1.workcode%20asc,t1.id%20desc%20&otype=DESC' \
               '&mode=run&customParams=null&selectedstrs=&pageId=mode_customsearch:25&formmodeFlag=1' \
               '&tableInstanceId=&pageIndex=1&tableString=' + tableHash
    result = opener.open(checkUrl).read()

    # 4.提取考勤记录
    xml = etree.XML(result)
    lastname = xml.xpath("//table/row[1]/col[@column='lastname']")[0].text
    lastname = etree.HTML(lastname).xpath("//a")[0].text
    status = xml.xpath("//table/row[1]/col[@column='status']")[0].text
    status = etree.HTML(status).xpath("//a")[0].text
    indate = xml.xpath("//table/row[1]/col[@column='indate']")[0].text
    intime = xml.xpath("//table/row[1]/col[@column='intime']")[0].text
    outdate = xml.xpath("//table/row[1]/col[@column='outdate']")[0].text
    outtime = xml.xpath("//table/row[1]/col[@column='outtime']")[0].text
    hours = xml.xpath("//table/row[1]/col[@column='hours']")[0].text

    # 5.发送邮件
    title = lastname + u"你好，昨天考勤:" + status
    content = u"签到：" + indate + " " + intime + u"\n签离：" + outdate + " " + outtime + u"\n时长：" + hours
    print title, content
    # 组装邮件
    msg = MIMEText(content, 'plain', 'utf-8')
    msg['From'] = _format_addr(u'OA通知 <%s>' % mail_from_addr)
    msg['To'] = _format_addr(lastname + ' <%s>' % mail_to_addr)
    msg['Subject'] = Header(title, 'utf-8').encode()
    # 发送邮件
    server = smtplib.SMTP(mail_smtp_server, mail_smtp_port)
    if mail_smtp_ssl:
        server.starttls()
    # server.set_debuglevel(1)
    server.login(mail_from_addr, mail_from_pass)
    server.sendmail(mail_from_addr, [mail_to_addr], msg.as_string())
    server.quit()
