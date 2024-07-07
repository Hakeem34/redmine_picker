import os
import sys
import re
import datetime
import subprocess
import errno
import time
import datetime
import shutil
from pathlib  import Path
from redminelib import Redmine


g_user_name       = 'tkubota'
g_pass            = 'ABCD1234'
g_url             = 'http://localhost:3000/'
g_target_projects = ['TestProject']
g_user_list       = []



class cUserData:
    def __init__(self, id, name):
        self.id         = id
        self.name       = name
        return


#/*****************************************************************************/
#/* ユーザー情報の登録                                                        */
#/*****************************************************************************/
def append_user(member):
    global g_user_list

    for user_data in g_user_list:
        if (user_data.id == member.user.id) and (user_data.name == member.user.name):
            return

    user_data = cUserData(member.user.id, member.user.name)
    g_user_list.append(user_data)
    print("[%d]:%s" % (user_data.id, user_data.name))
    return



#/*****************************************************************************/
#/* コマンドライン引数処理                                                    */
#/*****************************************************************************/
def check_command_line_option():
    global g_user_name
    global g_pass

    option = ""
    sys.argv.pop(0)
    for arg in sys.argv:
        if (option == "u"):
            g_user_name = arg
        elif (option == "p"):
            g_pass = arg
        elif (arg == "-u") or (arg == "--user"):
            option = "u"
        elif (arg == "-p") or (arg == "--pass"):
            option = "p"
        else:
            print("invalid arg : %s" % arg)

    return



#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    global g_user_name
    global g_pass
    global g_url
    global g_project_name

    check_command_line_option()
#   log_settings()

    redmine = Redmine(g_url, username=g_user_name, password=g_pass)
    print ("----------------------------------------------------------------------------------------------------------------")
    print (redmine)
    print (dir(redmine))
    print ("----------------------------------------------------------------------------------------------------------------")
    print (redmine.version)
    print (dir(redmine.version))
    issue = redmine.issue.get(12)
    print ("----------------------------------------------------------------------------------------------------------------")
    projects = redmine.project.all()
    print (projects)
    for project in projects:
        print('[' + str(project.id) + ']' + project.name)
        member_ships = project.memberships
        for member_ship in member_ships:
            append_user(member_ship)
            pass

    print ("----------------------------------------------------------------------------------------------------------------")
    users = redmine.user.all()
    print (users)
    for user in users:
#       print(dir(user))
        print('[' + str(user.id) + ']' + user.lastname + ' ' + user.firstname)
    print ("----------------------------------------------------------------------------------------------------------------")
    print (issue)
    print (dir(issue))
    print ('id:%d' % issue.id)
    print ('project:%s' % issue.project.name)
    print ('project_id:%d' % issue.project.id)
    print ('subject:%s' % issue.subject)
    print ('tracker:%s' % issue.tracker.name)
    print ('tracker_id:%d' % issue.tracker.id)
    print ('description:%s' % issue.description)
    print ('status:%s' % issue.status.name)
    print ('status:%d' % issue.status.id)
    print ('author:%s' % issue.author.name)
    print ('author_id:%d' % issue.author.id)
    return


if __name__ == "__main__":
    main()