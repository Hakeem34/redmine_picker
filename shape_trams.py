import os
import sys
import re
import shutil
import zipfile


g_trans_from = 'original.xlsx'
g_trans_to   = 'modified.xlsx'
g_temp_dir   = '_temp_st'
g_keep_temp  = False

#/*****************************************************************************/
#/* サブディレクトリの生成                                                    */
#/*****************************************************************************/
def make_directory(dirname):
    os.makedirs(os.path.join(dirname), exist_ok = True)
    return


#/*****************************************************************************/
#/* コマンドライン引数処理                                                    */
#/*****************************************************************************/
def check_command_line_option():
    global g_trans_from
    global g_trans_to
    global g_keep_temp

    option = ''
    sys.argv.pop(0)
    for arg in sys.argv:
#       print("arg : %s" % arg)
        if (option == 'from'):
            g_trans_from = arg
            option = ''
        elif (option == 'to'):
            g_trans_to = arg
            option = ''
        elif (arg == '-f'):
            option = 'from'
        elif (arg == '-t'):
            option = 'to'
        elif (arg == '-k'):
            g_keep_temp = True

    if not (os.path.isfile(g_trans_from)):
        print(f'-f is not file! {g_trans_from}')
        exit(-1)

    if not (os.path.isfile(g_trans_to)):
        print(f'-t is not file! {g_trans_to}')
        exit(-1)

    from_ext = os.path.splitext(g_trans_from)
    to_ext   = os.path.splitext(g_trans_to)

    if (from_ext[1] != '.xlsx') and (from_ext[1] != '.xlsm'):
        print(f'-f is not Excel file! {from_ext}')
        exit(-1)

    if (to_ext[1] != '.xlsx') and (to_ext[1] != '.xlsm'):
        print(f'-f is not Excel file! {to_ext}')
        exit(-1)

    return


#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def copy_and_unzip_targets():
    global g_trans_from
    global g_trans_to
    global g_temp_dir

    to_ext   = os.path.splitext(g_trans_to)
    out_file = to_ext[0] + '_trans' + to_ext[1]

    if (os.path.isdir(g_temp_dir)):
        print(f'{g_temp_dir} is already exist. Remove it!')
        shutil.rmtree(g_temp_dir)

    make_directory(g_temp_dir)
    from_extract = g_temp_dir + "\\" + os.path.basename(g_trans_from)
    to_extract   = g_temp_dir + "\\" + os.path.basename(g_trans_to)
    from_zip     = from_extract + '.zip'
    to_zip       = to_extract   + '.zip'

    shutil.copy2(g_trans_from,  from_zip)
    shutil.copy2(g_trans_to,    to_zip)

    shutil.unpack_archive(from_zip, from_extract)
    shutil.unpack_archive(to_zip,   to_extract)

    from_drawing = from_extract + r'\xl\drawings'
    from_rels    = from_extract + r'\xl\worksheets\_rels'
    to_drawing   = to_extract   + r'\xl\drawings'
    to_rels      = to_extract   + r'\xl\worksheets\_rels'

#   print(f'drawing : {from_drawing}')
#   print(f'rels    : {from_rels}')

    if (os.path.isdir(to_drawing)):
        files = os.listdir(from_drawing)
        for file in files:
            print(f'copy {file} to {to_drawing}')
            shutil.copy2(from_drawing + '\\' + file, to_drawing)

    else:
        shutil.copytree(from_drawing, to_drawing)

    if (os.path.isdir(to_rels)):
        pass
    else:
        shutil.copytree(from_rels,    to_rels)

    shutil.make_archive(out_file, format='zip', root_dir=to_extract)
    return


#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    check_command_line_option()

    copy_and_unzip_targets()

#   shutil.unpack_archive(g_trans_from, 'dir_out')

#   shutil.make_archive(g_trans_to, format='zip', root_dir='dir_out')
    return


if __name__ == "__main__":
    main()
