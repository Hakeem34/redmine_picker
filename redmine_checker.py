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
g_opt_include_sub_prj = 1
g_opt_journal_filters = []

g_target_project_ids = []
g_user_list          = []
g_cf_type_list       = []
g_status_type_list   = []
g_issue_list         = []
g_time_entry_list    = []
g_filter_limit       = 20


ATTR_NAME_DIC      = {
                         'id'                : '#',
                         'project'           : 'プロジェクト',
                         'tracker'           : 'トラッカー',
                         'parent'            : '親チケット',
                         'children'          : '子チケット',
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
                         'journals'          : '更新ID',
                     }



re_1st_line         = re.compile(r"^([^\n]+)\n")


#/*****************************************************************************/
#/* 作業時間クラス                                                            */
#/*****************************************************************************/
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


#/*****************************************************************************/
#/* チケット更新の詳細データクラス                                            */
#/*****************************************************************************/
class cDetailData:
    def __init__(self, detail):
        self.property   = detail.get('property')
        self.name       = detail.get('name')
        self.old_val    = enc_dec_str(detail.get('old_value'))
        self.new_val    = enc_dec_str(detail.get('new_value'))
        return


    #/*****************************************************************************/
    #/* データ保存対象フィルタのチェック                                          */
    #/*****************************************************************************/
    def filter_check(self):
        global g_opt_journal_filters

        if (self.property == 'attr'):
            for filter_attr in g_opt_journal_filters:
                if (self.name == filter_attr):
                    return 1
        elif (self.property == 'cf'):
            for filter_attr in g_opt_journal_filters:
                if (self.name == filter_attr):
                    return 1

        return 0


    #/*****************************************************************************/
    #/* 表示データ名取得                                                          */
    #/*****************************************************************************/
    def get_disp_name(self):
        name = self.name

        #/* カスタムフィールドの場合、nameにはカスタムフィールドのIDがstrで入っているので、カスタムフィールドの名前に変換する */
        if (self.property == 'cf'):
            name = get_custom_fieled_name(int(name))

        return name


    #/*****************************************************************************/
    #/* 表示データ変換                                                            */
    #/*****************************************************************************/
    def get_disp_value(self, value):
        if (value == None) or (value == ""):
            return "-"

        if (self.property == 'attr'):
            if (self.name == 'assigned_to_id'):
                value = get_user_data_by_id(int(value)).name
            elif (self.name == 'status_id'):
                value = get_issue_status_name(int(value))
        elif (self.property == 'cf'):
            value = get_custom_fieled_disp_value(int(self.name), value)

        return value

    def get_disp_old_value(self):
        return self.get_disp_value(self.old_val)

    def get_disp_new_value(self):
        return self.get_disp_value(self.new_val)


#/*****************************************************************************/
#/* チケット更新データクラス                                                  */
#/*****************************************************************************/
class cJournalData:
    def __init__(self, id, created_on, user_data):
        self.id               = id
        self.created_on       = created_on
        self.user             = user_data
        self.details          = []
        self.notes            = ""
        self.filter           = 0
        return


#/*****************************************************************************/
#/* ユーザーデータクラス                                                      */
#/*****************************************************************************/
class cUserData:
    def __init__(self, id, name):
        self.id           = id
        self.name         = name
        self.time_entries = []
        return

NONE_USER = cUserData(0, "-")


#/*****************************************************************************/
#/* カスタムフィールドデータクラス                                            */
#/*****************************************************************************/
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



#/*****************************************************************************/
#/* カスタムフィールドの型情報クラス                                          */
#/*****************************************************************************/
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


#/*****************************************************************************/
#/* チケットステータス情報クラス                                              */
#/*****************************************************************************/
class cIssueStatusType:
    def __init__(self, status):
        self.id         = status.id
        self.name       = status.name
        self.is_closed  = getattr(status, 'is_closed', 0)
        return


#/*****************************************************************************/
#/* チケットデータクラス                                                      */
#/*****************************************************************************/
class cIssueData:
    def __init__(self, id):
        self.id                = id
        self.project           = ""
        self.parent            = ""
        self.children          = []
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
    #/* 表示用の属性値取得                                                        */
    #/*****************************************************************************/
    def get_disp_attr(self, attribute):
        if (attribute == 'author') or (attribute == 'assigned_to'):
            object = getattr(self, attribute)
            return object.name
        elif (attribute == 'children'):
            text = ",".join(map(str,self.children))
            return text
        elif (attribute.isdigit()):
            for cf in self.custom_fields:
                if (cf.id == int(attribute)):
                    return cf.get_disp_value()

            return ""
        else:
            object = getattr(self, attribute)
            return object


    #/*****************************************************************************/
    #/* python-redmineからチケット情報の読み出し                                  */
    #/*****************************************************************************/
    def read_issue_data(self, issue):
        global g_opt_list_attrs

        self.project           = issue.project.name
        self.parent            = getattr_ex(issue, 'parent', 'id', 0)

        children               = getattr(issue, 'children', [])
        for child in children:
            self.children.append(child.id)

        self.subject           = issue.subject
        self.priority          = issue.priority.name
        self.tracker           = issue.tracker.name
        self.status            = issue.status.name
        self.author            = get_user_data_by_id(issue.author.id)

        assigned_user          = getattr(issue, 'assigned_to', NONE_USER)
        self.assigned_to       = get_user_data_by_id(assigned_user.id)

        self.created_on        = getattr(issue, 'created_on', "-")
        self.updated_on        = getattr(issue, 'updated_on', "-")

        self.closed_on         = getattr(issue, 'closed_on', "-")
        self.start_date        = getattr(issue, 'start_date', "-")
        self.done_ratio        = getattr(issue, 'done_ratio', 0)
        self.due_date          = getattr(issue, 'due_date', "-")
        self.estimated_hours   = getattr(issue, 'estimated_hours', 0)

        #/* カスタムフィールドの取得 */
        if (hasattr(issue, 'custom_fields')):
            for cf in issue.custom_fields:
                cf_data = cCustomFieledData(cf)
                self.custom_fields.append(cf_data)
                if (str(cf_data.id) in g_opt_list_attrs):
                    pass
                else:
                    print("CF[%d] is not in g_opt_list_attrs" % (cf_data.id))

        #/* 更新情報の取得 */
        if (hasattr(issue, 'journals')):
            for journal in issue.journals:
                journal_data = cJournalData(journal.id, journal.created_on, get_user_data_by_id(journal.user.id))
                for detail in journal.details:
                    detail_data = cDetailData(detail)
                    if (detail_data.filter_check()):
                        journal_data.filter = 1
                        journal_data.details.append(detail_data)

                journal_data.notes = omit_multi_line_str(getattr(journal, 'notes', ""))
                self.append_journal_data(journal_data)

        #/* 作業時間情報の取得 */
        total_spent_hours = 0
        if (hasattr(issue, 'time_entries')):
            for time_entry in issue.time_entries:
                te_data = find_time_entry(time_entry)
                if (te_data != None):
                    self.time_entries.append(te_data)
                    total_spent_hours += te_data.hours

        #/* total_spent_hoursがサポートされない場合は、time_entriesの合計値とする（本来は、子チケットの時間も集計するようだが・・・） */
        if (hasattr(issue, 'total_spent_hours')):
            self.total_spent_hours = issue.total_spent_hours
        else:
            self.total_spent_hours = total_spent_hours

        return


    #/*****************************************************************************/
    #/* 更新情報の登録                                                            */
    #/*****************************************************************************/
    def append_journal_data(self, journal_data):
        if (journal_data.filter > 0) or (journal_data.notes != ""):
            self.journals.append(journal_data)

        return


    #/*****************************************************************************/
    #/* チケット情報のログ出力                                                    */
    #/*****************************************************************************/
    def print_issue_data(self):
        print("--------------------------------- Issue ID : %d ---------------------------------" % (self.id))
        print("  Project         : %s" % (self.project))
        print("  Tracker         : %s" % (self.tracker))
        print("  Subject         : %s" % (self.subject))
        print("  Status          : %s" % (self.status))
        print("  Parent          : %s" % (self.parent))
        print("  Children        : %s" % (self.children))
        print("  Priority        : %s" % (self.priority))
        print("  Author          : %s" % (self.author.name))
        print("  AssignedTo      : %s" % (self.assigned_to.name))
        print("  CreatedOn       : %s" % (self.created_on))
        print("  UpdatedOn       : %s" % (self.updated_on))
        print("  ClosedOn        : %s" % (self.closed_on))
        print("  StartDate       : %s" % (self.start_date))
        print("  DueDate         : %s" % (self.due_date))
        print("  DoneRatio       : %s" % (self.done_ratio))
        print("  EstimatedHours  : %s" % (self.estimated_hours))
        print("  TotalSpentHours : %s" % (self.total_spent_hours))
        print("  CustomFields    : ")

        for cf_data in self.custom_fields:
            print("[%s][%s]:%s" % (cf_data.id, cf_data.name, cf_data.get_disp_value()))

        for te_data in self.time_entries:
            print("  Time Entry[%s][%s]:%s hours in %s by %s" % (te_data.id, te_data.created_on, te_data.hours, te_data.spent_on, te_data.user.name))

        for journal_data in self.journals:
            print("  Update[%s][%s]:%s" % (journal_data.id, journal_data.created_on, journal_data.user.name))
            for detail_data in journal_data.details:
                old_val = omit_multi_line_str(detail_data.old_val)
                new_val = omit_multi_line_str(detail_data.new_val)
                print("    Detail[%s][%s] %s -> %s" % (detail_data.property, detail_data.name, old_val, new_val))

        return


#/*****************************************************************************/
#/* S-JISでエラーとなる文字の排除                                             */
#/*****************************************************************************/
def enc_dec_str(value):
    if (type(value) is str):
        value = value.encode('cp932', 'replace').decode('cp932', 'replace')

    return value


#/*****************************************************************************/
#/* 複数行のテキストを省略して1行テキストに変換                               */
#/*****************************************************************************/
def omit_multi_line_str(value):
    if (type(value) is str):
        if (result := re_1st_line.match(value)):
            value = result.group(1).replace('\r', '') + '...'                                #/* 改行の含まれる値は無視して1行目だけを扱う */

    return value


#/*****************************************************************************/
#/* 孫attrの取得                                                              */
#/*****************************************************************************/
def getattr_ex(target, attr, sub_attr, default):
    if (hasattr(target, attr)):
        attr = getattr(target, attr)
        ret_val = getattr(attr, sub_attr, default)
    else:
        ret_val = default

    return ret_val



#/*****************************************************************************/
#/* チケット情報の読み出し                                                    */
#/*****************************************************************************/
def get_issue_data(issue):
    global g_issue_list

    for issue_data in g_issue_list:
        if (issue.id == issue_data.id):
            return issue_data

    issue_data = cIssueData(issue.id)
    issue_data.read_issue_data(issue)
    g_issue_list.append(issue_data)
    return issue_data


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
#/* カスタムフィールドの名称取得                                              */
#/*****************************************************************************/
def get_custom_fieled_name(id):
    global g_cf_type_list

    cf_type = get_custom_fieled_type(id)

    if (cf_type != None):
        return cf_type.name

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
#/* カスタムフィールドの表示値の取得                                          */
#/*****************************************************************************/
def get_custom_fieled_disp_value(id, value):
    if (get_custom_fieled_format(id) == "user"):
        if (value == None):
            return NONE_USER.name

        if (value == ""):
            return NONE_USER.name

        return get_user_data_by_id(int(value)).name
    elif (get_custom_fieled_format(id) == "enumeration"):
        if (value == None):
            return ""

        if (value == ""):
            return ""

        dic = get_custom_fieled_dictionary(id)
        return dic[str(value)]

    return value


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
def get_user_data(redmine, user):
    global g_user_list
    global g_time_entry_list
    global g_target_project_ids

    if (user == None):
        return cUserData(0, "不明なユーザー")

    for user_data in g_user_list:
        if (user_data.id == user.id) and (user_data.name == user.name):
            return user_data

    user_data = cUserData(user.id, user.name)

    for target_id in g_target_project_ids:
        time_entries = redmine.time_entry.filter(project_id = target_id, user_id = user.id)
        for time_entry in time_entries:
            te = cTimeEntryData(time_entry.id, time_entry.created_on, time_entry.spent_on, time_entry.hours, time_entry.user)
            te.project_name = time_entry.project.name
            te.issue_id     = getattr_ex(time_entry, 'issue', 'id', 0)
            te.updated_on   = time_entry.updated_on
            te.activity     = time_entry.activity
            print("TimeEntry[%s][%s]spent on : %s  %s hours by %s for #%s in %s, activity : %s" % (te.id, te.created_on, te.spent_on, te.hours, te.user.name, te.issue_id, te.project_name, te.activity))
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

    if (user_id == 0):
        return NONE_USER

    for user_data in g_user_list:
        if (user_data.id == user_id):
            return user_data

    print("missing user ID! [%d]" % (user_id))
    return cUserData(0, "不明なユーザー")


#/*****************************************************************************/
#/* 設定ファイル読み込み処理                                                  */
#/*****************************************************************************/
def read_setting_file(file_path):
    global g_opt_url
    global g_opt_api_key
    global g_opt_target_projects
    global g_opt_out_file
    global g_opt_list_attrs
    global g_opt_include_sub_prj
    global g_opt_journal_filters

    f = open(file_path, 'r')
    lines = f.readlines()

    re_opt_url        = re.compile(r"URL\s+: ([^\n]+)")
    re_opt_api_key    = re.compile(r"API KEY\s+: ([^\n]+)")
    re_opt_out_file   = re.compile(r"OUT FILE NAME\s+: ([^\n]+)")
    re_opt_tgt_prj    = re.compile(r"TARGET PROJECT\s+: ([^\n]+)")
    re_opt_list_attr  = re.compile(r"ISSUE LIST ATTR\s+: ([^\n]+)")
    re_opt_list_cf    = re.compile(r"ISSUE LIST CF\s+: ([0-9]+)")
    re_opt_sub_prj    = re.compile(r"INCLUDE SUB PRJ\s+: ([^\n]+)")
    re_opt_filter     = re.compile(r"JOURNAL FILTER\s+: ([^\n]+)")

    journal_append = 0
    for line in lines:
#       print ("line:%s" % line)
        if (result := re_opt_url.match(line)):
            g_opt_url = result.group(1)
        elif (result := re_opt_api_key.match(line)):
            g_opt_api_key = result.group(1)
        elif (result := re_opt_tgt_prj.match(line)):
            g_opt_target_projects.append(result.group(1))
        elif (result := re_opt_out_file.match(line)):
            g_opt_out_file = result.group(1)
        elif (result := re_opt_list_attr.match(line)):
            if (result.group(1) == 'journals'):
                journal_append = 1                        #/* 表示の都合上、更新情報の出力は一番末尾（右側）とする */
            else:
                g_opt_list_attrs.append(result.group(1))
        elif (result := re_opt_list_cf.match(line)):
            g_opt_list_attrs.append(result.group(1))
        elif (result := re_opt_sub_prj.match(line)):
            g_opt_include_sub_prj = int(result.group(1))
        elif (result := re_opt_filter.match(line)):
            g_opt_journal_filters.append(result.group(1))

    if (journal_append > 0):
        g_opt_list_attrs.append('journals')

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
#/* プロジェクト情報の取得                                                    */
#/*****************************************************************************/
def check_project_info(redmine):
    global g_opt_target_projects
    global g_target_project_ids

    print("--------------------------------- Check Project Datas ---------------------------------")
    projects = redmine.project.all()

    #/* 対象プロジェクトからIDを取得 */
    for project in projects:
        for target in g_opt_target_projects:
            if (target == project.name):
                g_target_project_ids.append(project.id)
                print("ID[%d] : %s" % (project.id, target))

    return


#/*****************************************************************************/
#/* ユーザー情報の取得                                                        */
#/*****************************************************************************/
def check_user_info(redmine):
    print("--------------------------------- Check User Datas ---------------------------------")
    users = redmine.user.all()

    #/* 対象プロジェクトからユーザー情報を取得 */
    for user in users:
        user.name = user.lastname + ' ' + user.firstname
        get_user_data(redmine, user)

    return


#/*****************************************************************************/
#/* カスタムフィールド情報の取得                                              */
#/*****************************************************************************/
def check_custom_fields(redmine):
    global g_cf_type_list

    print("--------------------------------- Check Custome Fields ---------------------------------")
    fields = redmine.custom_field.all()
    for cf in fields:
        cf_type = cCustomFieledType(cf)
        if (cf_type.type == "issue"):
            print("Custom Field[%s][%s]:%s" % (cf_type.id, cf_type.name, cf_type.format))
            g_cf_type_list.append(cf_type)

    return


#/*****************************************************************************/
#/* チケットステータス情報の取得                                              */
#/*****************************************************************************/
def check_issue_status(redmine):
    global g_status_type_list

    statuses = redmine.issue_status.all()

    print("--------------------------------- Check Issue Status Types ---------------------------------")
    for status in statuses:
        status_type = cIssueStatusType(status)
        print("[%s][%s] is_closed : %d" % (status_type.id, status_type.name, status_type.is_closed))
        g_status_type_list.append(status_type)


#/*****************************************************************************/
#/* チケットステータス名の取得                                                */
#/*****************************************************************************/
def get_issue_status_name(status_id):
    global g_status_type_list

    for status in g_status_type_list:
        if (status_id == status.id):
            return status.name

    return "-"


#/*****************************************************************************/
#/* 結果フォーマット行出力                                                    */
#/*****************************************************************************/
def output_issue_list_format_line(ws):
    global g_opt_list_attrs

    row = 1
    col = 1
    for item in g_opt_list_attrs:
        if (item in ATTR_NAME_DIC):
            ws.cell(row, col).value = ATTR_NAME_DIC[item]
            if (item == 'journals'):
                ws.cell(row, col + 1).value = '更新日'
                ws.cell(row, col + 2).value = '更新者'
                ws.cell(row, col + 3).value = 'コメント'
                ws.cell(row, col + 4).value = '詳細'
                ws.cell(row, col + 5).value = '更新値'
                ws.cell(row, col + 6).value = '更新前'
                ws.cell(row, col + 7).value = '更新後'
        else:
            ws.cell(row, col).value = get_custom_fieled_name(int(item))
        col += 1

    return


#/*****************************************************************************/
#/* 結果フォーマット行出力                                                    */
#/*****************************************************************************/
def output_issue_list_line(ws, row, issue_data):
    global g_opt_list_attrs

    col = 1
    offset = 0
    for item in g_opt_list_attrs:
        if (item != 'journals'):
            ws.cell(row, col).value = issue_data.get_disp_attr(item)
        else:
            for journal in issue_data.journals:
                ws.cell(row + offset, col    ).value = journal.id
                ws.cell(row + offset, col + 1).value = journal.created_on
                ws.cell(row + offset, col + 2).value = journal.user.name
                ws.cell(row + offset, col + 3).value = journal.notes
                for detail in journal.details:
                    ws.cell(row + offset, col + 4).value = detail.property
                    ws.cell(row + offset, col + 5).value = detail.get_disp_name()
                    ws.cell(row + offset, col + 6).value = detail.get_disp_old_value()
                    ws.cell(row + offset, col + 7).value = detail.get_disp_new_value()
                    offset += 1

                if (len(journal.details) == 0):
                    offset += 1                       #/* 詳細データがない場合のみ、次の行に進む */

            col += 6

        col += 1

    if (offset > 0):
        offset -= 1          #/* 1行目のjournalはカウントしないため、引いておく */

    return offset


#/*****************************************************************************/
#/* 結果出力                                                                  */
#/*****************************************************************************/
def output_all_issues_list(ws):
    global g_issue_list

    ws.title = "チケット一覧"
    output_issue_list_format_line(ws)
    row = 2
    for issue_data in g_issue_list:
        offset = output_issue_list_line(ws, row, issue_data)
        row += (1 + offset)

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
#/* チケット全確認                                                            */
#/*****************************************************************************/
def full_issue_check(redmine):
    global g_filter_limit
    global g_opt_include_sub_prj

    print("--------------------------------- Full Issue Check ---------------------------------")
    for target_id in g_target_project_ids:
        filter_offset = 0
        while(1):
            print("--------------------------------- ProjectID : %d, Filter Offset %d ---------------------------------" % (target_id, filter_offset))
            if (g_opt_include_sub_prj == 0):
                issues = redmine.issue.filter(project_id = target_id, subproject_id = '!*', status_id = '*', limit = g_filter_limit, offset = filter_offset)
            else:
                issues = redmine.issue.filter(project_id = target_id, status_id = '*', limit = g_filter_limit, offset = filter_offset)
            if (len(issues) == 0):
                break

            for issue in issues:
                issue_data = get_issue_data(issue)
                issue_data.print_issue_data()

            filter_offset += g_filter_limit

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
    global g_issue_list

    check_command_line_option()
    start_time = log_start()

    if (g_opt_api_key == ""):
        redmine = Redmine(g_opt_url, username=g_opt_user_name, password=g_opt_pass)
    else:
        redmine = Redmine(g_opt_url, key=g_opt_api_key)

    check_project_info(redmine)
    check_user_info(redmine)
    check_custom_fields(redmine)
    check_issue_status(redmine)

    if (g_opt_full_issues):
        full_issue_check(redmine)
    else:
        pass

    g_issue_list = sorted(g_issue_list, key=lambda issue: issue.id)
    output_datas()
    log_end(start_time)
    return


if __name__ == "__main__":
    main()
