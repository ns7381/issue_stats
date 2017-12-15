# coding=utf-8
import gitlab
import xlwt


class Member:
    def __init__(self, name):
        self._name = name
        self._bug = 0
        self._task = 0
        self._improvement = 0
        self._rejected = 0
        self._total = 0

    @property
    def name(self):
        return self._name

    @property
    def bug(self):
        return self._bug

    def bug_inc(self):
        self._bug += 1

    @property
    def task(self):
        return self._task

    def task_inc(self):
        self._task += 1

    @property
    def improvement(self):
        return self._improvement

    def improvement_inc(self):
        self._improvement += 1

    @property
    def rejected(self):
        return self._rejected

    def rejected_inc(self):
        self._rejected += 1

    @property
    def total(self):
        return self._total

    def total_inc(self):
        self._total += 1


fixed = 0


def issue_increase(member):
    if "Fixed" in issue.labels:
        global fixed
        fixed += 1
        return
    member.total_inc()
    if "Rejected" in issue.labels:
        member.rejected_inc()
        return
    elif "Bug" in issue.labels:
        member.bug_inc()
    elif "Improvements" in issue.labels:
        member.improvement_inc()
    elif "Task" in issue.labels:
        member.task_inc()


gl = gitlab.Gitlab('http://10.110.17.13', email='xxx', password='123456a?')
gl.auth()

project = gl.projects.get(112)
issues = project.issues.list(state='opened', all=True, milestone=u"里程碑3.7")
members = []
for issue in issues:
    not_exists = True
    for member in members:
        if member.name == issue.assignee.name:
            issue_increase(member)
            not_exists = False
            break
    if not_exists:
        member_new = Member(issue.assignee.name)
        issue_increase(member_new)
        members.append(member_new)

data = xlwt.Workbook()
table = data.add_sheet('sheet1')
table.write(0, 0, u'研发人员')
table.write(0, 1, 'Bug')
table.write(0, 2, 'Task')
table.write(0, 3, 'Improvements')
table.write(0, 4, 'Rejected')
table.write(0, 5, u'合计')
table.write(0, 6, 'Fixed')
issue_count = {'Bug': 0, 'Task': 0, 'Improvements': 0, 'Rejected': 0, 'Total': 0}
for i, member in enumerate(members):
    table.write(i + 1, 0, member.name)
    table.write(i + 1, 1, member.bug)
    table.write(i + 1, 2, member.task)
    table.write(i + 1, 3, member.improvement)
    table.write(i + 1, 4, member.rejected)
    table.write(i + 1, 5, member.total)
    issue_count['Bug'] += member.bug
    issue_count['Task'] += member.task
    issue_count['Improvements'] += member.improvement
    issue_count['Rejected'] += member.rejected
    issue_count['Total'] += member.total
table.write(len(members)+1, 0, u'统计')
table.write(len(members)+1, 1, issue_count['Bug'])
table.write(len(members)+1, 2, issue_count['Task'])
table.write(len(members)+1, 3, issue_count['Improvements'])
table.write(len(members)+1, 4, issue_count['Rejected'])
table.write(len(members)+1, 5, issue_count['Total'])
table.write(len(members)+1, 6, fixed)
data.save('result.xls')
