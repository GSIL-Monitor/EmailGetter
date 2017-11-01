import email
import pickle
import base64
import quopri

file = open('tmp.txt', 'rb')
messages = pickle.load(file)

i = 1
if __name__ == '__main__':
    for message in messages:
        subject = message.get('Subject')
        charset = subject.split('?')[1]
        mode = subject.split('?')[2]
        code = subject.split('?')[3]
        if mode.upper() == 'B':
            subject = str(base64.decodebytes(
                code.encode(encoding=charset)), encoding=charset)
        else:
            subject = str(quopri.decodestring(
                code.encode(encoding=charset)), encoding=charset)
        print(i)
        print(subject)
        i += 1
        # subject = unicode(dh[0][0], dh[0][1]).encode('utf8')
        # print("Date: " + message["Date"])
        # print("From: " + email.utils.parseaddr(message.get('from'))[1])
        # print("To: " + email.utils.parseaddr(message.get('to'))[1])
        # print("Subject: " + subject)
        # for part in message.walk():
        #     fileName = part.get_filename()
        #     contentType = part.get_content_type()
        #     mycode = part.get_content_charset()
        #     if fileName:
        #         data = part.get_payload(decode=True)
        #         h = email.Header.Header(fileName)
        #         dh = email.Header.decode_header(h)
        #         fname = dh[0][0]
        #         encodeStr = dh[0][1]
        #         if encodeStr != None:
        #             fname = fname.decode(encodeStr, mycode)
