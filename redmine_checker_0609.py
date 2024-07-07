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


g_user_name       = 'admin'
g_pass            = 'admin'
g_url             = 'http://localhost/redmine'
g_project_name    = []
g_target_users    = ['admin']



#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    global g_user_name
    global g_pass
    global g_url
    global g_project_name

#   check_command_line_option()
#   log_settings()

    redmine = Redmine(g_url, username=g_user_name, password=g_pass)
    print ("----------------------------------------------------------------------------------------------------------------")
    print (redmine)
    print (dir(redmine))
    print ("----------------------------------------------------------------------------------------------------------------")
    print (redmine.version)
    print (dir(redmine.version))
    issue = redmine.issue.get(1)
    print ("----------------------------------------------------------------------------------------------------------------")
    projects = redmine.project.all()
    print (projects)
    for project in projects:
        print('[' + str(project.id) + ']' + project.name)
#       member_ship = redmine.project_membership.get(project.id)
#       print(dir(member_ship))
#       project_user = member_ship.user
#       print(dir(project_user))
#       print(project_user.name)
        member_ships = project.memberships
#       print(dir(member_ships))
        for member_ship in member_ships:
#           print(dir(member_ship))
            print('  ' + member_ship.user.name)
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
