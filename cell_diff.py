import os
import sys
import re
import subprocess
import errno
import time
import datetime
import shutil
import openpyxl
import unicodedata


g_left_file  = ''
g_right_file = ''
g_out_path   = '.'

#/*****************************************************************************/
#/* 全角文字の数をカウント                                                    */
#/*****************************************************************************/
def get_full_width_count_in_text(text):
    count = 0
    for character in text:
        if unicodedata.east_asian_width(character) in 'FWA':
            count += 1

    return count


#/*****************************************************************************/
#/* コマンドライン引数処理                                                    */
#/*****************************************************************************/
def check_command_line_option():
    global g_left_file
    global g_right_file
    global g_out_path

    sys.argv.pop(0)
    for arg in sys.argv:
        if (os.path.isfile(arg)):
            if (g_left_file == ''):
                g_left_file = arg
            elif (g_right_file == ''):
                g_right_file = arg
            else:
                print("Too many args! : %s" % arg)
                exit(0)
        else:
            print("invalid arg : %s" % arg)


    if (g_left_file == '') or (g_right_file == ''):
        print("usage : cell_diff.py [file A] [file B]")
        exit(0)

    g_out_path = os.path.dirname(g_right_file)
    return



#/*****************************************************************************/
#/* 処理開始ログ                                                              */
#/*****************************************************************************/
def log_start():
    global g_out_path

    now = datetime.datetime.now()

    time_stamp = now.strftime('%Y%m%d_%H%M%S')
    log_path = g_out_path + '\\cell_diff_' + time_stamp + '.txt'
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
#/* 表示用文字列取得                                                          */
#/*****************************************************************************/
def get_disp_string(text, width):
#   print("get_disp_string Width  : [%d]" % (width))
#   print("get_disp_string Input  : [%s]" % (text))
    cut_flag = 0
    if (result := re.match(r"([^\n]*)\n", text)):
        text = result.group(1)
        cut_flag = 1

    length     = len(text)
    if (length > width):
        text = text[:(width - 3)]
        cut_flag = 1

    length     = len(text)
    full_count = get_full_width_count_in_text(text)
    over_diff  = (length + full_count + (cut_flag * 3)) - width

    while (over_diff > 0):
        cut_flag   = 1
        text = text[:(length - 1)]
        length     = len(text)
        full_count = get_full_width_count_in_text(text)
        length     = len(text)
        full_count = get_full_width_count_in_text(text)
        over_diff  = (length + full_count + (cut_flag * 3)) - width

    if (cut_flag):
        text = text + '...'

    text = text.ljust(width - full_count)
#   print("get_disp_string Output : [%s]" % (text))
    return text


#/*****************************************************************************/
#/* シートの比較                                                              */
#/*****************************************************************************/
def check_lr_sheets(ws_l, ws_r):
    print("  check sheet[%s]" % (ws_l.title))

    max_row_l = ws_l.max_row
    max_col_l = ws_l.max_column
    max_row_r = ws_r.max_row
    max_col_r = ws_r.max_column

    if (max_row_l == max_row_r) and (max_col_l == max_col_r):
        print("    max_row : %d, max_col = %d" % (max_row_r, max_col_r))
        for row in range(1, max_row_l):
            for col in range(1, max_col_l):
                val_l = ws_l.cell(row, col).value
                val_r = ws_r.cell(row, col).value
                if (val_l != val_r):
                    val_l = get_disp_string(str(val_l), 40)
                    val_r = get_disp_string(str(val_r), 40)
                    print("      差異(%4d, %4d) : [%s] vs [%s]" % (row, col, val_l, val_r))
    else:
        print("    max_row : [%d] vs [%d], max_col : [%d] vs [%d]" % (max_row_l, max_row_r, max_col_l, max_col_r))
        if (max_row_l > max_row_r):
            max_row = max_row_r
        else:
            max_row = max_row_l

        if (max_col_l > max_col_r):
            max_col = max_col_r
        else:
            max_col = max_col_l

        for row in range(1, max_row):
            for col in range(1, max_col):
                val_l = ws_l.cell(row, col).value
                val_r = ws_r.cell(row, col).value
                if (val_l != val_r):
                    val_l = get_disp_string(str(val_l), 40)
                    val_r = get_disp_string(str(val_r), 40)
                    print("      差異(%4d, %4d) : [%s] vs [%s]" % (row, col, val_l, val_r))

    return


#/*****************************************************************************/
#/* ブックの比較                                                              */
#/*****************************************************************************/
def check_lr_books():
    global g_left_file
    global g_right_file

    print("check book [%s] vs [%s]" % (g_left_file, g_right_file))
    wb_l = openpyxl.load_workbook(g_left_file,  data_only=True)
    wb_r = openpyxl.load_workbook(g_right_file, data_only=True)

    for ws_l in wb_l.worksheets:
        for ws_r in wb_r.worksheets:
            if (ws_l.title == ws_r.title):
#               print("title : %s" % ws_r.title)
                check_lr_sheets(ws_l, ws_r)
    return


#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    check_command_line_option()
    start_time = log_start()

     
#   print("TEST1:%s" % (get_disp_string("あいうえおかきくけこ", 10)))
#   print("TEST1:%s" % (get_disp_string("あ\nいうえおかきくけこ", 10)))
    check_lr_books()

    log_end(start_time)
    return


if __name__ == "__main__":
    main()
