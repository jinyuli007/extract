#-*- coding:utf-8 -*-

import xlrd
import re
import time

def open_excel(file='file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)

def excel_table_byindex(file='file.xls', colindex=0, byindex=0):
    data = open_excel(file)
    table = data.sheets()[byindex]
    nrows = table.nrows
    ncols = table.ncols
    print 'nrows:%d ncols:%d' % (nrows, ncols)
    # 某一行数据
    cellvalues = []
    for rownum in range(0, nrows):
        if table.cell(rownum, colindex).ctype == 1:
            cellvalue = table.cell(rownum, colindex).value.encode('utf-8')
            cellvalues.append(cellvalue)

    return cellvalues

def fmap(a, b):
    return (a, b)

def main():
    values = excel_table_byindex('topic_old_20161104.xlsx', 1)
    names = excel_table_byindex('topic_old_20161104.xlsx', 0)
    length = len(values)
    temp = map(fmap, names, values)
    dictDatas = dict(temp)
    # print length
    # print values[1]
    # datas = re.split(',|，', values[1])
    # 去除空元素
    # if '' in datas:
    #     datas.remove('')
    # print dictDatas['运动']

    temp = []
    list = []
    for name in names:
        value = dictDatas[name]
        datas = re.split(',|，', value)
        # 去除空元素
        if '' in datas:
            datas.remove('')
        for mem in datas:
            temp.append(name)
            temp.append(mem)
            list.append(temp)
            temp = []

    length = len(list)
    print length,time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

    # 用于记住每个list成员是否被合并的状态
    memstat = []
    for i in range(length):
        mydict = {i: 0}
        memstat.append(mydict)

    mylist = []

    for index in range(length):
        for index1 in range(index + 1, length):
            if list[index][1] == list[index1][0]:
                temp = [list[index][0], list[index][1], list[index1][1]]

                # str = '%s = %s, %s\n' % (temp[2], temp[0], temp[1])
                # fh.write(str)
                memstat[index][index] = 1
                memstat[index1][index1] = 1
                mylist.append(temp)
            elif list[index1][1] == list[index][0]:
                temp = [list[index1][0], list[index1][1], list[index][1]]
                # str = '%s = %s, %s\n' % (temp[2], temp[0], temp[1])
                # fh.write(str)
                memstat[index][index] = 1
                memstat[index1][index1] = 1
                mylist.append(temp)
            else:
                pass
        if memstat[index][index] == 0:
            temp = [list[index][0], list[index][1]]
            mylist.append(temp)

    print len(mylist), time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())

    finallist = []

    tempfinal = [['hello', 'world']]
    length = len(mylist)

    # 用于记住每个list成员是否被合并的状态
    memstat = []
    for i in range(length):
        mydict = {i: 0}
        memstat.append(mydict)

    try:
        fh = open('extract.txt', 'w+')
        for index in range(length):
            for index1 in range(index + 1, length):
                if mylist[index][-1] == mylist[index1][0]:

                    memstat[index][index] = 1
                    memstat[index1][index1] = 1
                    temp = mylist[index] + mylist[index1][1:]
                    str = temp[-1] + ' '
                    for data in temp:
                        str += '{0},'.format(data)
                    str = str.rstip(',') + '\n'
                    fh.write(str)
                    # finallist.append(temp)
                elif mylist[index1][-1] == mylist[index][0]:
                    memstat[index][index] = 1
                    memstat[index1][index1] = 1

                    temp = mylist[index1] + mylist[index][1:]

                    for t in temp:
                        print t
                    print '======================'
                    # 写文件
                    str = temp[-1] + ' '
                    for data in temp:
                        str += '{0},'.format(data)
                    str = str.rstip(',') + '\n'
                    fh.write(str)
                    # finallist.append(temp)
                #组合具有同一个儿子的所有父亲
                elif mylist[index][-1] == mylist[index1][-1]:
                    memstat[index][index] = 1
                    memstat[index1][index1] = 1
                    length = len(mylist[index])
                    templist = mylist[index][0: length-1] + mylist[index1]
                    newlist = list(set(templist))
                    newlist.sort(key=templist.index)

                    for i in range(len(tempfinal)):
                        for j in range(i+1, len(tempfinal)):
                            if newlist[-1] == temp[-1]:
                                templist = newlist.pop() + temp
                                templist1 = list(set(templist))
                                templist1.sort(key=templist.index)
                                tempfinal.append(templist1)


                    finallist.append(tempfinal)
                else:
                    pass

    except IOError:
        print 'IOError 没有找到文件或文件读取失败'

if __name__ == '__main__':
    main()