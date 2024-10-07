import os
import sys
import re
import time
import datetime
import openpyxl
import pathlib
from zipfile import BadZipFile

g_server_path  = 'c:\\src\\redmine_checker'
g_tgt_dir      = ''
g_out_path     = ''
g_out_file     = ''
g_filter_ext   =  [r'exe', r'bin']
g_filter_dir   =  [r'old', r'backup', r'.git', r'.svn']
g_filter_name  =  [r'^~.*\.(xlsx|xlsm|docx)', r'^.* - コピー\.(.+)$', r'.*Thumbs\.db$']
g_log_file     = None
g_opt_backup   = 0

#/* ファイル更新に関する設定 */
g_new_file_interactive      = 0            #/* 新規ファイルのコピーについて           （0：黙って更新、1：プロンプトで確認）           */
g_update_file_interactive   = 1            #/* サーバー側更新ファイルのコピーについて （0：黙って更新、1：プロンプトで確認）           */
g_conflict_file_interactive = 0            #/* ローカル側更新ファイルのコピーについて （0：更新しない、1：プロンプトで確認）           */
g_delete_file_interactive   = 1            #/* サーバー側削除ファイルの削除について   （0：黙って更新、1：プロンプトで確認）           */

g_target_list  =  []
g_rocal_list   =  []
g_sync_dt      =  None
g_file_list    =  None

PRE_COL_OFFSET = 10

#/*****************************************************************************/
#/* 対象ファイル情報クラス                                                    */
#/*****************************************************************************/
class cFileItem:
    def __init__(self):
        self.synced             = False
        self.file_name          = ''
        self.rel_path           = ''

        #/* サーバーのファイル情報 */
        self.update_date        = ''
        self.update_time        = ''
        self.size               = 0

        #/* ローカルコピーのファイル情報 */
        self.local_update_date  = ''
        self.local_update_time  = ''
        self.local_size         = 0

        #/* サーバーから取得した際のファイル情報 */
        self.base_update_date  = ''
        self.base_update_time  = ''
        self.base_size         = 0

        self.local_updated    = False
        self.updated_dt       = None
        self.pre_updated_dt   = None
        return

    def read_ws_row(self, ws, row):
        col = 2
        if (ws.cell(row, col).value != None) and (ws.cell(row, col).value != '-'):
            self.synced           = True
        else:
            self.synced           = False
        col += 1

        self.file_name        = ws.cell(row, col).value
        col += 1

        self.rel_path         = ws.cell(row, col).value
        col += 1

        self.update_date      = ws.cell(row, col).value
        col += 1

        self.update_time      = ws.cell(row, col).value
        col += 1

#       print_log(f'Item : {self.file_name}')
#       print_log(f'Date : {self.update_date} Type : {type(self.update_date)}')
#       print_log(f'Time : {self.update_time} Type : {type(self.update_time)}')
        if (type(self.update_date) is datetime.datetime) and (type(self.update_time) is datetime.time):
            self.updated_dt = datetime.datetime.combine(self.update_date.date(), self.update_time)
        else:
            print_log(f'ファイル情報が不正です')
            exit(-1)

        self.size             = ws.cell(row, col).value
        col += 1

        self.local_update_date  = ws.cell(row, col).value
        col += 1

        self.local_update_time  = ws.cell(row, col).value
        col += 1

        if (type(self.local_update_date) is datetime.datetime) and (type(self.local_update_time) is datetime.time):
            self.local_updated_dt = datetime.datetime.combine(self.local_update_date.date(), self.local_update_time)

        self.local_size         = ws.cell(row, col).value
        col += 1

        self.local_updated    = False
        col += 1
        return


#/*****************************************************************************/
#/* ファイルリスト情報クラス                                                  */
#/*****************************************************************************/
class cFileListInfo:
    def __init__(self):
        self.target_path   = ''
        self.sync_date     = None
        self.sync_time     = None
        self.log_file      = ''
        self.pre_sync_date = None
        self.pre_sync_time = None
        self.pre_log_file  = ''
        self.items         = []        #/* cFileItem */
        return


    def read_worksheet(self, ws):
        row = 2
        col = 3
        self.target_path   = ws.cell(row, col).value
        row += 1
        self.sync_date     = ws.cell(row, col).value
        self.pre_sync_date = ws.cell(row, col + PRE_COL_OFFSET).value
        row += 1
        self.sync_time     = ws.cell(row, col).value
        self.pre_sync_time = ws.cell(row, col + PRE_COL_OFFSET).value
        row += 1
        self.log_file      = ws.cell(row, col).value
        self.pre_log_file  = ws.cell(row, col + PRE_COL_OFFSET).value

        row += 3
        col  = 3
        while (ws.cell(row, col).value != None):
            item = cFileItem()
            item.read_ws_row(ws, row)
            self.items.append(item)
            row += 1
        return


#/*****************************************************************************/
#/* ファイル情報クラス                                                        */
#/*****************************************************************************/
class cFileInfo:
    def __init__(self, nw_path, root_path):
        self.file_name   = nw_path.name
        self.rel_path    = str(nw_path).replace(self.file_name, '')
        self.rel_path    = self.rel_path.replace(root_path, '')

        stat_info        = nw_path.stat()
        dt = datetime.datetime.fromtimestamp(stat_info.st_mtime)
        self.date        = dt.date()             #/* 更新日          */
        self.time        = dt.time()             #/* 更新時          */
        self.size        = stat_info.st_size     #/* ファイルサイズ  */
        return

    def get_time_stamp(self):
        date_time_text = str(self.date).replace('-', '') + '_' + str(self.time).replace(':', '')
        result = re.match(r'([0-9]+_[0-9]+)\.[0-9]+', date_time_text)
        if not (result):
            print_log('タイムスタンプのフォーマットが異常です {self.date} {self.time}')
            exit(-1)

        return result.group(1)


#/*****************************************************************************/
#/* ログ出力                                                                  */
#/*****************************************************************************/
def print_log(text):
    global g_log_file
    print(text)
    print(text, file=g_log_file)


#/*****************************************************************************/
#/* 処理開始ログ                                                              */
#/*****************************************************************************/
def log_start():
    global g_log_file
    global g_out_path
    global g_tgt_dir
    global g_sync_dt

    g_sync_dt = datetime.datetime.now()
    time_stamp = g_sync_dt.strftime('%Y%m%d_%H%M%S')
    date_stamp = g_sync_dt.strftime('%Y%m%d')

    log_path   = g_tgt_dir + '\\.SrvSync\\ServerSync_' + time_stamp + '.txt'
    g_log_file = open(log_path, "w")
#   sys.stdout = log_file

    start_time = time.perf_counter()
    g_sync_dt = datetime.datetime.now()
    print_log("処理開始 : " + str(g_sync_dt))
    print_log("----------------------------------------------------------------------------------------------------------------")
    return start_time


#/*****************************************************************************/
#/* 処理終了ログ                                                              */
#/*****************************************************************************/
def log_end(start_time):
    end_time = time.perf_counter()
    now = datetime.datetime.now()
    print_log("----------------------------------------------------------------------------------------------------------------")
    print_log("処理終了 : " + str(now))
    second = int(end_time - start_time)
    msec   = ((end_time - start_time) - second) * 1000
    minute = second / 60
    second = second % 60
    print_log("  %dmin %dsec %dmsec" % (minute, second, msec))
    return


#/*****************************************************************************/
#/* サブディレクトリの生成                                                    */
#/*****************************************************************************/
def make_directory(dirname):
#   print_log("make dir! %s" % dirname)
    os.makedirs(os.path.join(dirname), exist_ok = True)


#/*****************************************************************************/
#/* コマンドライン引数処理                                                    */
#/*****************************************************************************/
def check_command_line_option():
    global g_server_path
    global g_tgt_dir
    global g_out_path
    global g_out_file
    global g_opt_backup

    sys.argv.pop(0)
    for arg in sys.argv:
        if (arg == '-backup'):
            g_opt_backup = 1
        elif (os.path.isdir(arg)):
            g_server_path = arg
        else:
            print('invarid arg : %s' & arg)

    if (g_server_path == ''):
        print('server_sync.py [target_path]')
        exit(-1)

    g_server_path.rstrip('\\')                             #/* 末尾のバックスラッシュを削除 */
    g_tgt_dir  = pathlib.WindowsPath(g_server_path).name
    g_out_path = g_tgt_dir + '\\.SrvSync'
    g_out_file = g_out_path + '\\ServerSync.xlsx'
    return


#/*****************************************************************************/
#/* ディレクトリ名のフィルタ                                                  */
#/*****************************************************************************/
def dir_filter_check(file_path):
    global g_filter_dir

    if (file_path.is_dir()):
        for filter in g_filter_dir:
            if (filter == file_path.name):
                return True

    return False


#/*****************************************************************************/
#/* 拡張子のフィルタ                                                          */
#/*****************************************************************************/
def extension_filter_check(file_path):
    global g_filter_ext

#   print(f'file_path:{file_path},  suffix:{file_path.suffix}')
    if (file_path.is_file()):
        for filter in g_filter_ext:
            if ('.' + filter == file_path.suffix):
                return True

    return False


#/*****************************************************************************/
#/* ファイル名のmatchフィルタ                                                 */
#/*****************************************************************************/
def match_filter_check(file_path):
    global g_filter_name

    if (file_path.is_file()):
        for filter in g_filter_name:
            if (result := re.match(filter, str(file_path))):
                return True

    return False


#/*****************************************************************************/
#/* パス検索                                                                  */
#/*****************************************************************************/
def search_target_path(path, level):
    global g_filter_ext
    global g_filter_dir
    global g_filter_name
    global g_target_list
    global g_server_path

    print_log('search   : %s' % (path))
    network_path = pathlib.WindowsPath(path)
    if (network_path.exists()):
        files = network_path.iterdir()
        for file_name in files:
#           print_log('file_name : %s' % (file_name))
            if (dir_filter_check(file_name)):
#               print_log('Filter Directory! [%s]' % (file_name))
                continue
            elif (extension_filter_check(file_name)):
#               print_log('Filter Extension! [%s]' % (file_name))
                continue
            elif (match_filter_check(file_name)):
#               print_log('Filter match! [%s]' % (file_name))
                continue

            if (file_name.is_dir()):
                search_target_path(file_name, level + 1)
            else:
                try:
                    file_info = cFileInfo(file_name, g_server_path)
                except FileNotFoundError:
                    if len(str(file_name)) > 260:
                        print(f'ファイル[{file_name}]のパスが長すぎます')
                    else:
                        print(f'ファイル[{file_name}]が見つかりません')
                else:
                    g_target_list.append(file_info)

    else:
        print_log("network path is not available!")

    return


#/*****************************************************************************/
#/* ファイルリスト入力                                                        */
#/*****************************************************************************/
def in_file_list():
    global g_out_file
    global g_opt_backup
    global g_file_list

    if (os.path.isfile(g_out_file)):
        writable = os.access(g_out_file, os.W_OK)
        if not (writable):
            print_log('[%s]が読み取り専用です' % (g_out_file))
            exit(-1)

    wb = None
    try:
        wb = openpyxl.load_workbook(g_out_file, data_only=True)
    except PermissionError:
        pass
    except FileNotFoundError:
        print(f'ファイル[{g_out_file}]が見つかりません')
    except BadZipFile:
        print(f'ファイル[{g_out_file}]を開くことができません')
        exit(-1)
    else:
        pass

#   print_log(f'g_opt_backup is {g_opt_backup}')
    if (wb != None) and (g_opt_backup == 1):
        path = pathlib.WindowsPath(g_out_file)
        file_info = cFileInfo(path, g_server_path)
        time_stamp = file_info.get_time_stamp()
        back_up_path = g_out_file.replace(r'.xlsx', r'_' + time_stamp + r'.xlsx')
        print_log(f'リストファイルをバックアップします [{g_out_file}] -> [{back_up_path}]')
        wb.save(g_out_file.replace(r'.xlsx', r'_' + time_stamp + r'.xlsx'))

    if (wb != None):
        ws = wb['FileList']
        g_file_list = cFileListInfo()
        g_file_list.read_worksheet(ws)

        if (g_file_list.target_path != g_server_path):
            print(f'同期パスが異なります {g_server_path} vs {g_file_list.target_path}')
            exit(-1)

    return


#/*****************************************************************************/
#/* ファイルリスト出力                                                        */
#/*****************************************************************************/
def out_file_list():
    global g_sync_dt
    global g_target_list
    global g_out_file
    global g_file_list

    time_stamp = g_sync_dt.strftime('%Y%m%d_%H%M%S')
    log_path = g_out_path + '\\ServerSync_' + time_stamp + '.txt'

    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]
    ws.title = 'FileList'

    row = 2
    col = 2
    ws.cell(row, col                     ).value = '対象パス'
    ws.cell(row, col + 1                 ).value = g_server_path
    row += 1
    ws.cell(row, col                     ).value = '同期 日付'
    ws.cell(row, col + 1                 ).value = g_sync_dt.date()
    ws.cell(row, col + PRE_COL_OFFSET    ).value = '前回同期 日付'
    ws.cell(row, col + PRE_COL_OFFSET + 1).value = '-'
    row += 1
    ws.cell(row, col                     ).value = '同期 時間'
    ws.cell(row, col + 1                 ).value = g_sync_dt.time()
    ws.cell(row, col + PRE_COL_OFFSET    ).value = '前回同期 時間'
    ws.cell(row, col + PRE_COL_OFFSET + 1).value = '-'
    row += 1
    ws.cell(row, col                     ).value = 'ログファイル'
    ws.cell(row, col + 1                 ).value = log_path
    ws.cell(row, col + PRE_COL_OFFSET    ).value = '前回ログ'
    ws.cell(row, col + PRE_COL_OFFSET + 1).value = '-'
    ws.column_dimensions['B'].width              =  16
    ws.column_dimensions['C'].width              =  16
    ws.column_dimensions['L'].width              =  16
    ws.column_dimensions['M'].width              =  16

    row += 2
    col = 2
    ws.cell(row, col).value = '更新'
    col += 1
    ws.cell(row, col).value = 'ファイル名'
    col += 1
    ws.cell(row, col).value = '相対パス'
    col += 1
    ws.cell(row, col).value = '日付'
    col += 1
    ws.cell(row, col).value = '時間'
    col += 1
    ws.cell(row, col).value = 'サイズ'
    col += 1
    ws.cell(row, col).value = 'ローカル日付'
    col += 1
    ws.cell(row, col).value = 'ローカル時間'
    col += 1
    ws.cell(row, col).value = 'ローカルサイズ'
    col += 1
    ws.cell(row, col).value = '取得時日付'
    col += 1
    ws.cell(row, col).value = '取得時時間'
    col += 1
    ws.cell(row, col).value = '取得時サイズ'
    col += 1

    row += 1
    for target in g_target_list:
        print_log('[%s] : %s' % (target.rel_path, target.file_name))
        col  = 2
        ws.cell(row, col).value = '-'
        col += 1
        ws.cell(row, col).value = target.file_name
        col += 1
        ws.cell(row, col).value = target.rel_path
        col += 1
        ws.cell(row, col).value = target.date
        col += 1
        ws.cell(row, col).value = target.time
        col += 1
        ws.cell(row, col).value = target.size
        col += 1
        ws.cell(row, col).value = '-'
        col += 1

        row += 1

    while True:
        try:
            wb.save(g_out_file)
        except PermissionError:
            input(f"ServerSync.xlsxが更新できません。クローズしてから、Enterを入力してください。")
        else:
            print(f'Excel file saved successfully at {g_out_file}')
            break
    return


#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    global g_out_path
    global g_tgt_dir
    check_command_line_option()
    make_directory(g_out_path)
    start_time = log_start()

    in_file_list()
    search_target_path(g_server_path, 0)
    out_file_list()

    log_end(start_time)
    return


if __name__ == "__main__":
    main()





