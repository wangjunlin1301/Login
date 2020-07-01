from jira import jira_request,LoginData,blogName
import time,re


testDate = str.upper(time.strftime('%d-%b',time.localtime(time.time())))
postingDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))

def CreateBlog():
    #登陆 Conflunence 
    loginUrl = 'https://confluence.blackline.corp/dologin.action'
    r1 = jira_request('POST',url=loginUrl,data=LoginData)
    # 获取token
    token = re.findall(r'atlassian-token" content="(.*?)"',r1.text)[0]

    #创建Blog
    createBlogUrl = 'https://confluence.blackline.corp/pages/createblogpost.action?spaceKey=QA' 
    r2=jira_request('GET',url=createBlogUrl)

    # 提交内容
    postData = 'New'
    content = {
        'title':'[TEST UPDATES,%s] %s'%(testDate,blogName),
        'queryString':'spaceKey=QA',
        'spaceKey':'QA',
        'originalReferrer': 'https://confluence.blackline.corp/pages/viewrecentblogposts.action?key=QA',
        'PostingDate':postingDate,
        'wysiwygContent': postData,
        'atl_token':token

    }

    commitUrl = 'https://confluence.blackline.corp/pages/docreateblogpost.action'
    r3=jira_request('POST',url=commitUrl,data=content)
    print(r3.request.headers)
    # with open('test.html','wb') as f:
    #     for i in r3.request.headers.items:
    #         f.write(i)
CreateBlog()