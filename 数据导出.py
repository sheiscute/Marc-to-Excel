
# coding: utf-8

# In[10]:


import xlwt
import xlrd
#定义find_repeat函数，找出字符在文本中出现的所有位置，返回的为列表。
def find_repeat(content,source):
    s_index=0;e_index=len(source)
    elem_index=[]
    while(s_index<e_index):
        try:
            temp=source.index(content,s_index,e_index)
            elem_index.append(temp)
            s_index=temp+1
        except ValueError:
            break
    return elem_index
#读取文件，检测数据总数,list1为用结束符chr(29)分割的数据切片列表
print("MarctoExcel [版本 1.0]"
      "Made by 上海图书馆 采编中心报刊部")
road=input("你好，欢迎使用MarctoExcel！本程序可将txt格式的Marc数据转出至Excel格式的表格.\n请输入待转数据包路径，或直接将其拖入本窗口，按回车键继续：")
text=open(road).read()
roadlist=road.split("\\")
rootroad=""
for i in roadlist[:-1]:
    rootroad=rootroad+i+"\\"
list1=text.split(chr(29))
while "" in list1:
    list1.remove("")
while "\n" in list1:
    list1.remove("\n")
while chr(30) in list1:
    list1.remove(chr(30))
print("共计"+str(len(list1))+"条数据。")
#找出特定字段的位置n,list2为用字段分隔符chr(30)分割的字段切片列表
list2=text.split(chr(30)) 
#在第2个字段（目次区）中寻找目标字段，且返回目标字段位置
def find_position(xlist):
    plist=[]
    for i in xlist:
        if i%12==0:
            plist.append(i)
        else:
            pass
    return plist
#读取目标数据,并写入excel
field_list=("011","100","110","200","207","210","215","326","690","711","801","856","888","905")
#设定目标字段
outputa=xlwt.Workbook()
sheet1=outputa.add_sheet("sheet1")
cola=0
for field in field_list:
    wsubfield_list=[]
    for i in list1:
        i = i.lstrip("\n")
        subdata_list=[]
        try:
            base_pos=int(i[12:17])
            sublist=i.split(chr(30)) 
            while "" in sublist:
                sublist.remove("")
            number_list=find_repeat(field,sublist[0]) 
            while 0 in number_list:
                number_list.remove(0)
            if number_list==[]:
                anonumber_list=find_repeat(field,sublist[1]) 
                while 0 in anonumber_list:
                    anonumber_list.remove(0)
                if anonumber_list==[]:
                    pass
                else:
                    position_list=find_position(anonumber_list)
                    if position_list==[]:
                        pass
                    else:
                        for n in position_list:
                            n=n+len(sublist[0])+1
                            f_len=int(i[n+3:n+7])
                            f_pos=int(i[n+7:n+12])+base_pos
                            bi=i.encode("gbk")
                            middata=bi[f_pos:f_pos+f_len].decode("gbk")[3:]
                            data=middata.strip(chr(30))
                            lista=data.split(chr(31))
                            for a in lista:
                                wsubfield_list.append(a[0].lower())
            else:
                position_list=find_position(number_list)
                if position_list==[]:
                    pass
                else:
                    for n in position_list:
                        f_len=int(i[n+3:n+7])
                        f_pos=int(i[n+7:n+12])+base_pos
                        bi=i.encode("gbk")
                        middata=bi[f_pos:f_pos+f_len].decode("gbk")[3:]
                        data=middata.strip(chr(30))
                        lista=data.split(chr(31))
                        for a in lista:
                            print(a)
                            wsubfield_list.append(a[0].lower())
        except TypeError:
            print("1类型错误"+str(f_len)+"\n"+i)
        except ValueError:
            print("1找不到"+field+"字段,数据详情如下："+i)
        else:
            pass
    subfield_list=set(wsubfield_list)
    print(subfield_list)
    for i in subfield_list:
        sheet1.write(0,cola,field+"$"+i)
        cola=cola+1
route=rootroad+"表头输出试验.xls"
outputa.save(route)
ExcelFile=xlrd.open_workbook(route)
sheet=ExcelFile.sheet_by_index(0)
#生成各字段列表
field_rows=sheet.row_values(0)
#按照字段，写入字段内容
outputb=xlwt.Workbook()
sheet1=outputb.add_sheet("sheet1")
colb=0
row=1
for field in field_list:
    for i in list1:
        i = i.lstrip("\n")
        subdata_list=[]
        try:
            base_pos=int(i[12:17])
            sublist=i.split(chr(30))           
            number_list=find_repeat(field,sublist[0]) 
            while 0 in number_list:
                number_list.remove(0)
            if number_list==[]:
                anonumber_list=find_repeat(field,sublist[1]) 
                while 0 in anonumber_list:
                    anonumber_list.remove(0)
                if anonumber_list!=[]:
                    position_list=find_position(anonumber_list)
                    if position_list==[]:
                         print("缺少"+field+"字段")
                    #sheet1.write(row,colb,"none")
                    else:
                        dic={}
                        for n in position_list:
                            n=n+len(sublist[0])+1
                            f_len=int(i[n+3:n+7])
                            f_pos=int(i[n+7:n+12])+base_pos
                            bi=i.encode("gbk")
                            middata=bi[f_pos:f_pos+f_len].decode("gbk")[3:]
                            data=middata.strip(chr(30))
                            #sfdata_list为子字段内容列表
                            sfdata_list=data.split(chr(31))
                            for x in sfdata_list:
                                subn=field_rows.index(field+"$"+x[0].lower())
                                print(x)
                                #subn为该子字段在表格中的位置
                                dic.setdefault(x[0].lower(),[]).append(x[1:]+"|")
                        keys=list(dic.keys())
                        for key in keys:
                            excelnum=field_rows.index(field+"$"+key.lower())
                            dicklist=dic[key.lower()]
                            ndicklist=[]
                            for k in dicklist:
                                if k==dicklist[-1]:
                                    k=k.strip("|")
                                    ndicklist.append(k)
                                elif len(dicklist)==1:
                                    k=k.strip("|")
                                    ndicklist.append(k)
                                else:
                                    ndicklist.append(k)
                            sheet1.write(row,excelnum,ndicklist)
                else:
                    print("2错误数据：缺少"+field+"字段，或头标区存在非法分隔符，数据详情如下："+i)
                #sheet1.write(row,colb,"none")
            else:
                position_list=find_position(number_list)
                if position_list==[]:
                    print("2缺少"+field+"字段")
                    #sheet1.write(row,colb,"none")
                else:
                    dic={}
                    for n in position_list:
                        f_len=int(i[n+3:n+7])
                        f_pos=int(i[n+7:n+12])+base_pos
                        bi=i.encode("gbk")
                        middata=bi[f_pos:f_pos+f_len].decode("gbk")[3:]
                        data=middata.strip(chr(30))
                        #sfdata_list为子字段内容列表
                        sfdata_list=data.split(chr(31))
                        for x in sfdata_list:
                            subn=field_rows.index(field+"$"+x[0].lower())
                            #subn为该子字段在表格中的位置
                            dic.setdefault(x[0].lower(),[]).append(x[1:]+"|")
                    keys=list(dic.keys())
                    for key in keys:
                        excelnum=field_rows.index(field+"$"+key.lower())
                        dicklist=dic[key.lower()]
                        ndicklist=[]
                        for k in dicklist:
                            if k==dicklist[-1]:
                                k=k.strip("|")
                                ndicklist.append(k)
                            elif len(dicklist)==1:
                                k=k.strip("|")
                                ndicklist.append(k)
                            else:
                                ndicklist.append(k)
                        sheet1.write(row,excelnum,ndicklist)
        except ValueError:
            print("2错误"+str(f_len)+"\n"+i)
        else:
            pass
        row=row+1
    row=1
fieldnamecol=0
for fieldname in field_rows:
    sheet1.write(0,fieldnamecol,fieldname)
    fieldnamecol=fieldnamecol+1
outputb.save(rootroad+"导出数据.xls")


# In[9]:



# In[18]:



# In[23]:


