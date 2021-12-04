import openpyxl
import time
# =======================准备阶段==读取表格并基本提示=========================
wb = openpyxl.load_workbook("P:\makePrograms\Projects\Tennoter\WordLib.xlsx")
db = wb["Sheet1"]
print("已成功加载单词库文件,在输完一组单词后输入!!help可以查看帮助,切勿直接关掉")

#judge函数用于检测输入值是否是一个指令,若是则返回状态码或执行这个指令
def judge(a):
    if len(a)==0:
        return -1
    elif a[0]=='!' and a[1]=='!':
        if a=="!!help":
            print("!!是每一个命令的开始标志,命令必须在输完一组单词后,也就是应用显示“请输入英文”时才能生效\n\texit:退出并保存程序\n\tundo:撤销上一个单词的键入")
            return 1
        elif a=="!!exit":
            return 2
        elif a=="!!undo":
            return 3
        else:#错误的指令
            return -1
    else:
        return 0

# =========================准备阶段==建立待写入序列=========================
timE =[]#单词时间缓冲区
word = []#单词内容缓冲区
translate = []#单词翻译缓冲区


#=====================运行阶段==输入单词短语并放入缓冲区=====================
ipt = ""
while True:
    ipt = input("请输入英文>>>")
    statuecode=judge(ipt)
    if statuecode==0:#正常输入,查询是否和输入缓冲区中之前输入的单词有重复
        if word.count(ipt)==0:#不重复
            word.append(ipt)#把单词放入缓冲区
            timE.append(time.time())#记录写单词的时间
            ipt = input("请输入汉译>>>")
            translate.append(ipt)#把汉译放入缓冲区
        else:#有重复,时间取旧,汉译不重复的话就追加
            pos=word.index(ipt)
            ipt = input("请输入汉译>>>")
            if translate[pos]!=ipt:
                translate[pos]=translate[pos]+";"+ipt#把汉译追加到缓冲区
            continue
    elif statuecode==2:#退出应用
        break
    elif statuecode==3:#撤销上一个单词
        if len(word)>=1:
            print("上一个单词",word[-1],"将会被删去,确定吗?  默认(t)rue / (f)alse")
            ipt = input()
            if ipt=="f" or ipt=="false":
                continue
            else:
                del word[-1]
                del timE[-1]
                del translate[-1]
                continue
        else:
            print("指令执行失败:没有上一个单词")
            continue
    else:
        continue



#=======================收尾阶段==查重并把单词存入数据库======================
predbword = db["B"]#获取已经记录的数组

#把表格里的数据转换成给人看的东西
dbword = []#给人看的list包含所有单词
for i in range(0,len(predbword)):
    dbword.append(predbword[i].value)

#对于缓冲区的每个单词进行查重和写入,此处的查重仅限于查询当前输入缓冲区与数据库的单词是否重复
add=1#填写单词时下移的位数,即为本次程序已经创建的单词数量+1
for i in range(0,len(word)):
    if dbword.count(word[i])==0:#没有查到重复项
        db["B"+str(len(dbword)-dbword.count(None)+add)]=word[i]#把创建的单词内容写到数据库
        db["A"+str(len(dbword)-dbword.count(None)+add)]=timE[i]#把创建的单词录入时间写到数据库
        db["C"+str(len(dbword)-dbword.count(None)+add)]=translate[i]#把创建的单词汉译写到数据库
        add+=1
    else:
        pos = dbword.index(word[i])#确认重复单词下标
        db["A"+str(pos+1)]=timE[i]#把创建的单词录入时间更新到数据库
        if db["C"+str(pos+1)].value!=translate[i]:
            db["C"+str(pos+1)]=db["C"+str(pos+1)].value+";"+translate[i]#如果记录的汉译不同,就把创建的单词汉译追加到数据库




#==========================终了阶段==保存数据库=============================
wb.save("P:\makePrograms\Projects\Tennoter\WordLib.xlsx")