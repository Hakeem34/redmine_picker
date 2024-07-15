import os
import sys
import re
import subprocess
import errno
import time
import datetime
import shutil
import openpyxl

from pathlib  import Path
from redminelib import Redmine


g_opt_user_name    = 'tkubota'
g_opt_pass         = 'ABCD1234'
g_opt_full_issues  = 0
g_opt_url          = 'http://localhost:3000/'

g_target_projects  = ['TestProject', 'TestSubProject']
g_user_list        = []
g_issue_list       = []
g_filter_limit     = 10


class cUserData:
    def __init__(self, id, name):
        self.id         = id
        self.name       = name
        self.activities = []
        return

NONE_USER = cUserData(0, "-")


class cCustomFieled:
    def __init__(self, cf):
        self.id    = cf.id
        self.name  = cf.name
        self.value = cf.value
        return


class cIssueData:
    def __init__(self, id):
        self.id                = id
        self.project           = ""
        self.parent            = ""
        self.priority          = ""
        self.tracker           = ""
        self.subject           = ""
        self.status            = ""
        self.author            = ""
        self.assigned_to       = ""
        self.start_date        = ""
        self.created_on        = ""
        self.updated_on        = ""
        self.closed_on         = ""
        self.done_ratio        = ""
        self.due_date          = ""
        self.estimated_hours   = 0
        self.total_spent_hours = 0
        self.journals          = []
        self.time_entries      = []
        self.custom_fields     = []
        return


#/*****************************************************************************/
#/* チケット情報の読み出し                                                    */
#/*****************************************************************************/
def get_issue_data(issue_id):
    global g_issue_list

    for issue_data in g_issue_list:
        if (issue_id == issue_data.id):
            return issue_data

    issue_data = cIssueData(issue_id)
    g_issue_list.append(issue_data)
    return issue_data


#/*****************************************************************************/
#/* チケット情報の読み出し                                                    */
#/*****************************************************************************/
def print_issue_data(issue_data):
    print("--------------------------------- Issue ID : %d ---------------------------------" % (issue_data.id))
    print("  Project         : %s" % (issue_data.project))
    print("  Tracker         : %s" % (issue_data.tracker))
    print("  Subject         : %s" % (issue_data.subject))
    print("  Status          : %s" % (issue_data.status))
    print("  Parent          : %s" % (issue_data.parent))
    print("  Priority        : %s" % (issue_data.priority))
    print("  Author          : %s" % (issue_data.author.name))
    print("  AssignedTo      : %s" % (issue_data.assigned_to.name))
    print("  CreatedOn       : %s" % (issue_data.created_on))
    print("  UpdatedOn       : %s" % (issue_data.updated_on))
    print("  ClosedOn        : %s" % (issue_data.closed_on))
    print("  StartDate       : %s" % (issue_data.start_date))
    print("  DueDate         : %s" % (issue_data.due_date))
    print("  DoneRatio       : %s" % (issue_data.done_ratio))
    print("  EstimatedHours  : %s" % (issue_data.estimated_hours))
    print("  TotalSpentHours : %s" % (issue_data.total_spent_hours))
    print("  CustomFields    : ")
    for cf_data in issue_data.custom_fields:
        print("[%s][%s]:%s" % (cf_data.id, cf_data.name, cf_data.value))

    return


#/*****************************************************************************/
#/* チケット情報の読み出し                                                    */
#/*****************************************************************************/
def read_issue_data(issue):
    issue_data = get_issue_data(issue.id)
    issue_data.project = issue.project.name
    if (hasattr(issue, 'parent')):
        issue_data.parent  = issue.parent.id
#       print("type parent is : " + str(type(issue_data.parent)))
    else:
        issue_data.parent  = 0

    issue_data.subject           = issue.subject
    issue_data.priority          = issue.priority.name
#   print("type priority is : " + str(type(issue_data.priority)))
    issue_data.tracker           = issue.tracker
    issue_data.subject           = issue.subject
    issue_data.status            = issue.status.name
    issue_data.author            = get_user_data(issue.author)
#   print("type author is : " + str(type(issue_data.author)))

    if (hasattr(issue, 'assigned_to')):
        issue_data.assigned_to   = get_user_data(issue.assigned_to)
    else:
        issue_data.assigned_to   = NONE_USER

    issue_data.created_on        = issue.created_on
    issue_data.updated_on        = issue.updated_on
    issue_data.closed_on         = issue.closed_on
    issue_data.start_date        = issue.start_date
    issue_data.done_ratio        = issue.done_ratio
    issue_data.due_date          = issue.due_date
    issue_data.estimated_hours   = issue.estimated_hours
    issue_data.total_spent_hours = issue.total_spent_hours

    for cf in issue.custom_fields:
#       print(cf)
#       print(dir(cf))
        cf_data = cCustomFieled(cf)
        issue_data.custom_fields.append(cf_data)

    print_issue_data(issue_data)
    return


#/*****************************************************************************/
#/* ユーザー情報の登録                                                        */
#/*****************************************************************************/
def get_user_data(user):
    global g_user_list

    if (user == None):
        return cUserData(0, "不明なユーザー")

    for user_data in g_user_list:
        if (user_data.id == user.id) and (user_data.name == user.name):
            return user_data

    user_data = cUserData(user.id, user.name)
    g_user_list.append(user_data)
    print("New user! [%d]:%s" % (user_data.id, user_data.name))
    return user_data


#/*****************************************************************************/
#/* コマンドライン引数処理                                                    */
#/*****************************************************************************/
def check_command_line_option():
    global g_opt_user_name
    global g_opt_pass
    global g_opt_full_issues

    option = ""
    sys.argv.pop(0)
    for arg in sys.argv:
        if (option == "u"):
            g_opt_user_name = arg
        elif (option == "p"):
            g_opt_pass = arg
        elif (arg == "-u") or (arg == "--user"):
            option = "u"
        elif (arg == "-p") or (arg == "--pass"):
            option = "p"
        elif (arg == "-f") or (arg == "--full"):
            g_opt_full_issues = 1
        else:
            print("invalid arg : %s" % arg)

    return



#/*****************************************************************************/
#/* 処理開始ログ                                                              */
#/*****************************************************************************/
def log_start():
    now = datetime.datetime.now()

    time_stamp = now.strftime('%Y%m%d_%H%M%S')
    log_path = 'redmine_checker_' + time_stamp + '.txt'
    log_file = open(log_path, "w")
    sys.stdout = log_file

    start_time = time.perf_counter()
    now = datetime.datetime.now()
    print("処理開始 : " + str(now))
    print ("----------------------------------------------------------------------------------------------------------------")
    return start_time


#/*****************************************************************************/
#/* 処理終了ログ                                                              */
#/*****************************************************************************/
def log_end(start_time):
    end_time = time.perf_counter()
    now = datetime.datetime.now()
    print ("----------------------------------------------------------------------------------------------------------------")
    print("処理終了 : " + str(now))
    second = int(end_time - start_time)
    msec   = ((end_time - start_time) - second) * 1000
    minute = second / 60
    second = second % 60
    print("  %dmin %dsec %dmsec" % (minute, second, msec))
    return


#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    global g_opt_user_name
    global g_opt_pass
    global g_opt_url
    global g_target_projects
    global g_opt_full_issues

    check_command_line_option()
    start_time = log_start()

    redmine = Redmine(g_opt_url, username=g_opt_user_name, password=g_opt_pass)
    projects = redmine.project.all()

    #/* 対象プロジェクトからユーザー情報を取得 */
    for project in projects:
        print(project.name)
        print(dir(project))
        for target in g_target_projects:
            if (target == project.name):
                for member_ship in project.memberships:
                    get_user_data(member_ship.user)

#               print(project.issue_custom_fields)
#               print(len(project.issue_custom_fields))
#               for cf in project.issue_custom_fields:
#                   print(dir(cf))



    if (g_opt_full_issues):
        print("--------------------------------- Full Issue Check ---------------------------------")
        filter_offset = 0
        while(1):
            print("--------------------------------- Filter Offset %d ---------------------------------" % (filter_offset))
            issues = redmine.issue.filter(status_id = '*', limit = g_filter_limit, offset = filter_offset)
            if (len(issues) == 0):
                break

            for issue in issues:
#               print(issue.id)
#               print(dir(issue))
                read_issue_data(issue)

            filter_offset += g_filter_limit

        issue = redmine.issue.get(12)
    else:
        pass

    log_end(start_time)
    return


if __name__ == "__main__":
    main()
