import email
import pickle
import base64
import quopri
import xlrd
import xlwt

keyword = "仙多山"
row = 3
col = 8

file = open('tmp.txt', 'rb')
messages = pickle.load(file)


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
    print("\n创建EXCEL文件......")
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
    ret = []
    for message in messages[::-1]:
        print("-----------------------------------------")
        subject = message.get('Subject')
        subject = decode_message(subject)
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
                    if len(fileName.split('.')) > 1 and fileName.split('.')[1].upper() in ['XLS', 'XLSX']:
                        data = part.get_payload(decode=True)
                        fEx = open("tmp.xlsx", 'wb')
                        fEx.write(data)
                        fEx.close()

                        xl = xlrd.open_workbook("tmp.xlsx")
                        sheet = xl.sheet_by_index(0)
                        net = sheet.cell(3, 8).value

                        ret.append([fileName, net])
    to_excel(ret, 'net')
