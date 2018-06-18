# coding=utf-8
import gitlab
import xlwt


class Problem:
    def __init__(self, title, description, assignee, pro, due_date, state):
        self.title = title
        self.description = description
        self.assignee = assignee
        self.pro = pro
        self.due_date = due_date
        self.state = state


gl = gitlab.Gitlab('http://git.com', private_token='CahYaF3dQhamLsskb8_N')


def generate_issue_by_project(project_name):
    project = gl.projects.get(project_name)
    issues = project.issues.list(all=True)
    problems = []
    for issue in issues:
        problems.append(Problem(issue.attributes['title'],
                                issue.attributes['description'],
                                issue.attributes['assignee']['name'] if issue.attributes['assignee'] else '',
                                project.attributes['name'],
                                issue.attributes['due_date'],
                                issue.attributes['state']))
    return problems




data = xlwt.Workbook()
table = data.add_sheet('git_issues')
table.write(0, 0, u'项目组')
table.write(0, 1, u'问题')
table.write(0, 2, u'描述')
table.write(0, 3, u'责任人')
table.write(0, 4, u'project')
table.write(0, 5, u'计划完成时间')
table.write(0, 6, u'状态')
total = 1


def write_xls(project_group, project_name, is_write=False):
    global total
    pro_problems = generate_issue_by_project(project_name)
    if is_write:
        table.write(total, 0, project_group)
    for i, problem in enumerate(pro_problems):
        table.write(i + total, 1, problem.title)
        table.write(i + total, 2, problem.description)
        table.write(i + total, 3, problem.assignee)
        table.write(i + total, 4, problem.pro)
        table.write(i + total, 5, problem.due_date)
        table.write(i + total, 6, problem.state)
    total += len(pro_problems)


write_xls('DevOps', 'trident/trident-web', True)
write_xls('DevOps', 'trident/trident')
data.save('iop_git_issues.xls')
