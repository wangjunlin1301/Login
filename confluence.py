from jira import jira_request, LoginData, blogName
import time, re
from dateV import USDate

testDate = USDate
postingDate = time.strftime('%Y-%m-%d', time.localtime(time.time()))
date = time.strftime('%Y/%m/%d', time.localtime(time.time()))
blogTitle = '[TEST UPDATES,%s] %s' % (testDate, blogName)


def CreateBlog():
    #登陆 Conflunence
    loginUrl = 'https://confluence.blackline.corp/dologin.action'
    r1 = jira_request('POST', url=loginUrl, data=LoginData)
    # 获取token
    token = re.findall(r'atlassian-token" content="(.*?)"', r1.text)[0]

    #创建Blog
    createBlogUrl = 'https://confluence.blackline.corp/pages/createblogpost.action?spaceKey=QA'
    r2 = jira_request('GET', url=createBlogUrl)
    # print(r2.content)
    # 提交内容
    postData = 'New'
    draftId = re.findall(r'draftId" value="(.*?)"', r2.text)[0]
    draftShareId = re.findall(r'draftShareId" value="(.*?)"', r2.text)[0]
    syncRev = re.findall(r'syncRev" value="(.*?)"', r2.text)[0]

    content = {
        'title':
        blogTitle,
        'queryString':
        'spaceKey=QA',
        'spaceKey':
        'QA',
        'originalReferrer':
        'https://confluence.blackline.corp/display/QA/%s/%s' %
        (date, blogTitle),
        'PostingDate':
        postingDate,
        'wysiwygContent':
        postData,
        'atl_token':
        token,
        'draftId':
        draftId,
        'entityId':
        draftId,
        'draftShareId':
        draftShareId,
        'syncRev':
        syncRev
    }
    conHeaders = {
        'User-Agent':
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:71.0) Gecko/20100101 Firefox/71.0',
        'Content-Type':
        'application/x-www-form-urlencoded',
        'Host':
        'confluence.blackline.corp',
        'Location':
        'https://confluence.blackline.corp/display/QA/%s/%s' %
        (date, blogTitle),
        'Origin':
        'https://confluence.blackline.corp',
        'Referer':
        'https://confluence.blackline.corp/pages/resumedraft.action?draftId=%s&draftShareId=%s'
        % (draftId, draftShareId)
    }

    commitUrl = 'https://confluence.blackline.corp/pages/docreateblogpost.action'
    r3 = jira_request('POST', url=commitUrl, data=content, headers=conHeaders)
    print(r2.request.prepare_cookies)
    with open('test.html', 'wb') as f:
        for i in r3.iter_content():
            f.write(i)


CreateBlog()