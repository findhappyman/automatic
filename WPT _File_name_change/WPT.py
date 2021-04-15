import os
import easygui as eg

def get_file_name(dirpath=''):
    file_name = []
    list1 = os.listdir(dirpath)
    for i in range(len(list1)):
        if os.path.isfile(os.path.join(dirpath,list1[i])):
            if '~' not in list1[i]:
                file_name.append(os.path.join(dirpath, list1[i]))
    return file_name

def change_name(oldname,newname):
    '''
    ussage:change file name
    :param oldname:
    :param newname:
    :return:
    '''

    with open(dirapth)


if __name__ == '__main__':

    dirpath = eg.diropenbox(msg='请指定发票文件夹的位置： ', title="Pre-alert tracking V1.0\t   作者：Henry Xue ")
    change_name('AQKKWPT_MATNR_MAS_DATA','mmd.txt')
    change_name('AQKKWPT_SBT_DESCRIPTIO','SBTdesc.txt')
    change_name('AQKKWPT_SBT_STRATEGY','STYPEstrategy.txt')
    change_name('AQKKWPT_STOCK','stock.txt')
    change_name('AQKKWPT_STYPE_SETTINGS','STYPEsettings.txt')



    path_filename = get_file_name(dirpath)
    for i in path_filename:
