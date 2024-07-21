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


g_opt_user_name       = 'tkubota'
g_opt_pass            = 'ABCD1234'
g_opt_api_key         = ""
g_opt_full_issues     = 0
g_opt_url             = 'http://localhost:3000/'
g_opt_list_attrs      = []
g_opt_list_cfs        = []
g_opt_out_file        = 'redmine_result_%ymd.xlsx'
g_opt_target_projects = []


g_user_list        = []
g_cf_type_list     = []
g_issue_list       = []
g_time_entry_list  = []
g_filter_limit     = 20


ATTR_NAME_DIC      = {
                         'id'                : '#',
                         'project'           : 'プロジェクト',
                         'tracker'           : 'トラッカー',
                         'parent'            : '親チケット',
                         'status'            : 'ステータス',
                         'subject'           : 'タイトル',
                         'author'            : '作成者',
                         'created_on'        : '作成日',
                         'priority'          : '優先度',
                         'assigned_to'       : '担当者',
                         'updated_on'        : '更新日',
                         'closed_on'         : '完了日',
                         'due_date'          : '期日',
                         'done_ratio'        : '進捗率',
                         'estimated_hours'   : '予定工数',
                         'total_spent_hours' : '作業時間',
                     }


class cTimeEntryData:
    def __init__(self, id, created_on, spent_on, hours, user_data):
        self.id           = id
        self.created_on   = created_on
        self.user         = user_data
        self.spent_on     = spent_on
        self.hours        = hours
        self.updated_on   = ""
        self.activity     = ""
        self.project_name = ""
        self.issue_id     = ""
        return


class cDetailData:
    def __init__(self, detail):
        self.property   = detail.get('property')
        self.name       = detail.get('name')
        self.old_val    = detail.get('old_value')
        self.new_val    = detail.get('new_value')
        return


class cJournalData:
    def __init__(self, id, created_on, user_data):
        self.id         = id
        self.created_on = created_on
        self.user       = user_data
        self.details    = []
        self.notes      = ""
        return


class cUserData:
    def __init__(self, id, name):
        self.id           = id
        self.name         = name
        self.time_entries = []
        return

NONE_USER = cUserData(0, "-")


class cCustomFieledData:
    def __init__(self, cf):
        self.id    = cf.id
        self.name  = cf.name
        self.value = cf.value
        return

    def get_disp_value(self):
        if (get_custom_fieled_format(self.id) == "user"):
            if (self.value == None):
                return NONE_USER.name

            if (self.value == ""):
                return NONE_USER.name

            return get_user_data_by_id(int(self.value)).name
        elif (get_custom_fieled_format(self.id) == "enumeration"):
            if (self.value == None):
                return ""

            if (self.value == ""):
                return ""

            dic = get_custom_fieled_dictionary(self.id)
            return dic[str(self.value)]

        return self.value



class cCustomFieledType:
    def __init__(self, cf):
        self.id         = cf.id
        self.name       = cf.name
        self.type       = cf.customized_type
        self.format     = cf.field_format
        self.dictionary = {}

        if (self.format == 'enumeration'):
            for value_label in cf.possible_values:
                self.dictionary[value_label['value']] = value_label['label']

#           print(self.dictionary)
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

    def get_disp_attr(self, attribute):
        if (attribute == 'author') or (attribute == 'assigned_to'):
            object = getattr(self, attribute)
            return object.name
        else:
            object = getattr(self, attribute)
            return object


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
        print("[%s][%s]:%s" % (cf_data.id, cf_data.name, cf_data.get_disp_value()))

    for te_data in issue_data.time_entries:
        print("  Time Entry[%s][%s]:%s hours in %s by %s" % (te_data.id, te_data.created_on, te_data.hours, te_data.spent_on, te_data.user.name))

    for journal_data in issue_data.journals:
        print("  Update[%s][%s]:%s" % (journal_data.id, journal_data.created_on, journal_data.user.name))
        for detail_data in journal_data.details:
            print("    Detail[%s][%s] %s -> %s" % (detail_data.property, detail_data.name, detail_data.old_val, detail_data.new_val))

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
    issue_data.tracker           = issue.tracker.name
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
        cf_data = cCustomFieledData(cf)
        issue_data.custom_fields.append(cf_data)

    for journal in issue.journals:
#       print(journal)
#       print(dir(journal))
#       print("[%s][%s]%s" % (journal.id, journal.created_on, journal.user.name))
        journal_data = cJournalData(journal.id, journal.created_on, get_user_data_by_id(journal.user.id))
        for detail in journal.details:
#           print("---detail---")
#           print(detail)
#           print(dir(detail))
#           for key in detail.keys():
#               print(key)
            detail_data = cDetailData(detail)
            journal_data.details.append(detail_data)

        journal_data.notes = journal.notes
        issue_data.journals.append(journal_data)

    for time_entry in issue.time_entries:
        te_data = find_time_entry(time_entry)
        if (te_data != None):
            issue_data.time_entries.append(te_data)
        else:
            print("time entry not found! id = %s" % (time_entry.id))

    print_issue_data(issue_data)
    return


#/*****************************************************************************/
#/* カスタムフィールド情報の取得                                              */
#/*****************************************************************************/
def get_custom_fieled_type(id):
    global g_cf_type_list

    for cf_type in g_cf_type_list:
        if (cf_type.id == id):
            return cf_type

    return None


#/*****************************************************************************/
#/* カスタムフィールドのフォーマット情報の取得                                */
#/*****************************************************************************/
def get_custom_fieled_format(id):
    global g_cf_type_list

    cf_type = get_custom_fieled_type(id)

    if (cf_type != None):
        return cf_type.format

    return ""


#/*****************************************************************************/
#/* カスタムフィールドのフォーマット情報の取得                                */
#/*****************************************************************************/
def get_custom_fieled_dictionary(id):
    global g_cf_type_list

    cf_type = get_custom_fieled_type(id)

    if (cf_type != None):
        return cf_type.dictionary

    return {}


#/*****************************************************************************/
#/* 作業時間情報の検索                                                        */
#/*****************************************************************************/
def find_time_entry(te):
    global g_time_entry_list

    for te_data in g_time_entry_list:
        if (te_data.id == te.id):
            return te_data

    return None


#/*****************************************************************************/
#/* ユーザー情報の登録                                                        */
#/*****************************************************************************/
def get_user_data(user):
    global g_user_list
    global g_time_entry_list

    if (user == None):
        return cUserData(0, "不明なユーザー")

    for user_data in g_user_list:
        if (user_data.id == user.id) and (user_data.name == user.name):
            return user_data

    user_data = cUserData(user.id, user.name)

    for time_entry in user.time_entries:
#       print(time_entry)
#       print(dir(time_entry))
        te = cTimeEntryData(time_entry.id, time_entry.created_on, time_entry.spent_on, time_entry.hours, time_entry.user)
        te.project_name = time_entry.project.name
        te.issue_id     = time_entry.issue.id
        te.updated_on   = time_entry.updated_on
        te.activity     = time_entry.activity
        print("TimeEntry[%s][%s]spent on : %s  %s hours by %s for #%s in %s, activity : %s" % (time_entry.id, time_entry.created_on, time_entry.spent_on, time_entry.hours, time_entry.user.name, time_entry.issue.id, time_entry.project.name, time_entry.activity))
        user_data.time_entries.append(te)
        g_time_entry_list.append(te)

    g_user_list.append(user_data)
    print("New user! [%d]:%s" % (user_data.id, user_data.name))
    return user_data


#/*****************************************************************************/
#/* ユーザー情報の登録                                                        */
#/*****************************************************************************/
def get_user_data_by_id(user_id):
    global g_user_list

    for user_data in g_user_list:
        if (user_data.id == user_id):
            return user_data

    print("missing user ID! [%d]" % (user_id))
    return cUserData(0, "不明なユーザー")


#/*****************************************************************************/
#/* コマンドライン引数処理                                                    */
#/*****************************************************************************/
def read_setting_file(file_path):
    global g_opt_url
    global g_opt_api_key
    global g_opt_target_projects
    global g_opt_out_file
    global g_opt_list_attrs

    f = open(file_path, 'r')
    lines = f.readlines()

    re_opt_url        = re.compile(r"URL\s+: ([^\n]+)")
    re_opt_api_key    = re.compile(r"API KEY\s+: ([^\n]+)")
    re_opt_out_file   = re.compile(r"OUT FILE NAME\s+: ([^\n]+)")
    re_opt_tgt_prj    = re.compile(r"TARGET PROJECT\s+: ([^\n]+)")
    re_opt_list_att   = re.compile(r"ISSUE LIST ATTR\s+: ([^\n]+)")

    for line in lines:
#       print ("line:%s" % line)
        if (result := re_opt_url.match(line)):
            g_opt_url = result.group(1)
        elif (result := re_opt_api_key.match(line)):
#           print("api key : %s" % (result.group(1)))
            g_opt_api_key = result.group(1)
        elif (result := re_opt_tgt_prj.match(line)):
#           print("target project : %s" % (result.group(1)))
            g_opt_target_projects.append(result.group(1))
        elif (result := re_opt_out_file.match(line)):
#           print("output file : %s" % (result.group(1)))
            g_opt_out_file = result.group(1)
        elif (result := re_opt_list_att.match(line)):
#           print("attr : %s" % (result.group(1)))
            g_opt_list_attrs.append(result.group(1))

    f.close()
    return

#/*****************************************************************************/
#/* コマンドライン引数処理                                                    */
#/*****************************************************************************/
def check_command_line_option():
    global g_opt_user_name
    global g_opt_pass
    global g_opt_full_issues
    global g_opt_api_key

    option = ""
    sys.argv.pop(0)
    for arg in sys.argv:
        if (option == "u"):
            g_opt_user_name = arg
            option = ""
        elif (option == "p"):
            g_opt_pass = arg
            option = ""
        elif (option == "k"):
            g_opt_api_key = arg
            option = ""
        elif (option == "s"):
            read_setting_file(arg)
            option = ""
        elif (arg == "-u") or (arg == "--user"):
            option = "u"
        elif (arg == "-p") or (arg == "--pass"):
            option = "p"
        elif (arg == "-k") or (arg == "--key"):
            option = "k"
        elif (arg == "-s") or (arg == "--set_file"):
            option = "s"
        elif (arg == "-f") or (arg == "--full"):
            g_opt_full_issues = 1
        else:
            print("invalid arg : %s" % arg)

    return



#/*****************************************************************************/
#/* 処理開始ログ                                                              */
#/*****************************************************************************/
def log_start():
    global g_opt_out_file
    now = datetime.datetime.now()

    time_stamp = now.strftime('%Y%m%d_%H%M%S')
    log_path = 'redmine_checker_' + time_stamp + '.txt'
    log_file = open(log_path, "w")
    sys.stdout = log_file

    start_time = time.perf_counter()
    now = datetime.datetime.now()
    print("処理開始 : " + str(now))
    print ("----------------------------------------------------------------------------------------------------------------")

    yyyymmdd = now.strftime('%Y%m%d')
    hhmmss   = now.strftime('%H%M%S')
    g_opt_out_file = g_opt_out_file.replace('%ymd', yyyymmdd)
    g_opt_out_file = g_opt_out_file.replace('%hms', hhmmss)
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
#/* ユーザー情報の取得                                                        */
#/*****************************************************************************/
def check_project_member_ships(redmine):
    global g_opt_target_projects

    print("--------------------------------- Check User Datas ---------------------------------")
    projects = redmine.project.all()

    #/* 対象プロジェクトからユーザー情報を取得 */
    for project in projects:
#       print(project.name)
#       print(dir(project))
        for target in g_opt_target_projects:
            if (target == project.name):
                for member_ship in project.memberships:
#                   print(dir(member_ship))
                    if (hasattr(member_ship, 'user')):
#                       print(dir(member_ship.user))
                        get_user_data(member_ship.user)

    return


#/*****************************************************************************/
#/* カスタムフィールド情報の取得                                              */
#/*****************************************************************************/
def check_custom_fields(redmine):
    global g_cf_type_list

    print("--------------------------------- Check Custome Fields ---------------------------------")
    fields = redmine.custom_field.all()
    for cf in fields:
#       print(cf)
#       print(dir(cf))
#       if (cf.field_format == 'enumeration'):
#           print(cf.possible_values)
#       print(cf.customized_type)
#       print(cf.field_format)
        cf_type = cCustomFieledType(cf)
        if (cf_type.type == "issue"):
            print("Custom Field[%s][%s]:%s" % (cf_type.id, cf_type.name, cf_type.format))
            g_cf_type_list.append(cf_type)

    return


#/*****************************************************************************/
#/* 結果フォーマット行出力                                                    */
#/*****************************************************************************/
def output_issue_list_format_line(ws):
    global g_opt_list_attrs

    row = 1
    col = 1
    for item in g_opt_list_attrs:
        ws.cell(row, col).value = ATTR_NAME_DIC[item]
        col += 1

    return


#/*****************************************************************************/
#/* 結果フォーマット行出力                                                    */
#/*****************************************************************************/
def output_issue_list_line(ws, row, issue_data):
    global g_opt_list_attrs

    col = 1
    for item in g_opt_list_attrs:
        ws.cell(row, col).value = issue_data.get_disp_attr(item)
        col += 1

    return


#/*****************************************************************************/
#/* 結果出力                                                                  */
#/*****************************************************************************/
def output_all_issues_list(ws):
    global g_issue_list

    ws.title = "チケット一覧"
    output_issue_list_format_line(ws)
    row = 2
    for issue_data in g_issue_list:
        output_issue_list_line(ws, row, issue_data)
        row += 1

    return


#/*****************************************************************************/
#/* 結果出力                                                                  */
#/*****************************************************************************/
def output_datas():
    global g_opt_out_file
    wb = openpyxl.Workbook()
    output_all_issues_list(wb.worksheets[0])

    wb.save(g_opt_out_file)
    return


#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    global g_opt_user_name
    global g_opt_pass
    global g_opt_url
    global g_opt_full_issues
    global g_opt_api_key

    check_command_line_option()
    start_time = log_start()

    if (g_opt_api_key == ""):
        redmine = Redmine(g_opt_url, username=g_opt_user_name, password=g_opt_pass)
    else:
        redmine = Redmine(g_opt_url, key=g_opt_api_key)

    check_project_member_ships(redmine)
    check_custom_fields(redmine)

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

    output_datas()
    log_end(start_time)
    return


if __name__ == "__main__":
    main()
