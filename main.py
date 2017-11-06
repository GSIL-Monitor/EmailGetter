import base64
import email
# from email import parser
import pickle
import poplib
import quopri
import time
import sys

import xlrd
import xlwt

host = "pop.exmail.qq.com"
user = "service@fundbj.com"
pwd = "fed68318067"
keyword = "仙多山"
row = 4
col = 9
top = 20
t1, t2, t3 = (0, 0, 0)


def decode_message(message):
    ret = ''
    for s in message.split('\n'):
        if len(s.split('?')) > 4:
            charset = s.split('?')[1]
            mode = s.split('?')[2]
            code = s.split('?')[3]
            if mode.upper() == 'B':
                ret += str(base64.decodebytes(
                    code.encode(encoding=charset)), encoding=charset)
            else:
                ret += str(quopri.decodestring(
                    code.encode(encoding=charset)), encoding=charset)
    return ret


def to_excel(ret_list, filename):
    ''' 生成EXCEL文件 '''
    print("创建EXCEL文件。。。")
    try:
        excel = xlwt.Workbook()
        sheet = excel.add_sheet("my_sheet")
        for row in range(0, len(ret_list)):
            for col in range(0, len(ret_list[row])):
                sheet.write(row, col, ret_list[row][col])
        excel.save(filename + ".xls")
    except:
        print("创建失败:" + sys.exc_info()[0])
    print("创建完成")


if __name__ == '__main__':
    user = input("输入邮箱地址：")
    pwd = input("输入邮箱密码：")

    print('正在登录。。。')
    t1 = time.time()
    try:
        pop = poplib.POP3_SSL(host)
        pop.user(user)
        pop.pass_(pwd)
        # print('邮件数: %d, 大小: %d bytes' % pop.stat())
        num = pop.stat()[0]
    except poplib.error_proto as e:
        print(e)
        sys.exit(0)

    t2 = time.time()
    print('耗时：%.3f 秒' % (t2 - t1))

    try:
        keyword = input("输入邮件标题关键字：")
        top = int(input("输入获取邮件数："))
        row = int(input("输入行号："))
        col = int(input("输入列号："))
    except ValueError as e:
        print('输入数据类型错误')

    print('正在获取邮件。。。')
    top = top if num > top else num
    messages = [pop.retr(i) for i in range(num, num - top, -1)]
    messages = [b"\n".join(mssg[1]) for mssg in messages]
    # messages = [parser.BytesParser().parsebytes(mssg) for mssg in messages]
    messages = [email.message_from_bytes(mssg) for mssg in messages]

    t3 = time.time()
    print('耗时：%.3f 秒' % (t3 - t2))

    pop.quit()

    # file = open('tmp.txt', 'wb')
    # pickle.dump(messages, file)
    # file = open('tmp.txt', 'rb')
    # messages = pickle.load(file)

    ret = []
    for message in messages:
        subject = message.get('Subject')
        if type(subject) == str:
            subject = decode_message(subject)
        else:
            h = email.header.decode_header(subject)
            subject = h[0][0].decode(h[0][1]) if h[0][1].upper() in ['UTF-8','GBK','GB2312'] else h[0][0].decode('gbk')
        if keyword in subject:
            # print(subject)
            # print("Date: " + message["Date"])
            # print("From: " + email.utils.parseaddr(message.get('from'))[1])
            # print("To: " + email.utils.parseaddr(message.get('to'))[1])
            for part in message.walk():
                fileName = part.get_filename()
                contentType = part.get_content_type()  # 文件类型application/octet-stream
                code = part.get_content_charset()
                if fileName:
                    fileName = decode_message(fileName)
                    ext = fileName.split('.')[1] if len(
                        fileName.split('.')) > 1 else ''
                    if ext.upper() in ['XLS', 'XLSX'] and '估值' in fileName:
                        data = part.get_payload(decode=True)
                        fEx = open("tmp." + ext, 'wb')
                        fEx.write(data)
                        fEx.close()

                        xl = xlrd.open_workbook("tmp." + ext)
                        sheet = xl.sheet_by_index(0)
                        net = sheet.cell(row - 1, col - 1).value

                        ret.append([fileName, net])
    to_excel(ret, 'net')
