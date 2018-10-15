# coding: utf-8
import xlwt
from jira.client import JIRA

jira = JIRA(server='http://', basic_auth=('', '123456a?'))

data = xlwt.Workbook()
table = data.add_sheet('git_issues')
table.write(0, 0, u'Sprint')
table.write(0, 1, u'Issue key')
table.write(0, 2, u'主题')
table.write(0, 3, u'状态')
table.write(0, 4, u'到期日')
table.write(0, 5, u'汇总进度')
total = 1

for project in jira.projects():
    issues_in_proj = jira.search_issues('project=' + project.key)
    problems = []
    for i, issue in enumerate(issues_in_proj):
        if issue.fields.customfield_10004:
            table.write(i + total, 1, issue.fields.customfield_10004[0].split('name=')[1].split(',')[0])
        table.write(i + total, 2, issue.key)
        table.write(i + total, 3, issue.fields.summary)
        table.write(i + total, 4, issue.fields.status.name)
        table.write(i + total, 5, issue.fields.duedate)
        if hasattr(issue.fields.progress, 'percent'):
            table.write(i + total, 6, issue.fields.progress.percent)

data.save('iop_git_issues.xls')

