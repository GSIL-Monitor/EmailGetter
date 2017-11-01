import poplib
import email

host = "pop.exmail.qq.com"
user = "service@fundbj.com"
pwd = "fed68390036"

if __name__ == '__main__':
    pop = poplib.POP3_SSL(host)
    pop.user(user)
    pop.pass_(pwd)

    print(pop.stat())

    messages = [pop.retr(i) for i in range(1, len(pop.list()[1]) + 1)]
    messages = [b"\n".join(mssg[1]) for mssg in messages]
    messages = [parser.Parser().parsestr(mssg) for mssg in messages]

    for message in messages:
        subject = message.get('subject')     
        h = email.Header.Header(subject)  
        dh = email.Header.decode_header(h)  
        subject = unicode(dh[0][0], dh[0][1]).encode('utf8')
        print("Date: " + message["Date"])
        print("From: " + email.utils.parseaddr(message.get('from'))[1])
        print("To: " + email.utils.parseaddr(message.get('to'))[1])
        print("Subject: " + subject)
        for part in message.walk():
            fileName = part.get_filename()
            contentType = part.get_content_type()
            mycode = part.get_content_charset()
            if fileName:
                data = part.get_payload(decode=True)
                h = email.Header.Header(fileName)
                dh = email.Header.decode_header(h)
                fname = dh[0][0]
                encodeStr = dh[0][1]
                if encodeStr != None:
                    fname = fname.decode(encodeStr, mycode)
    pop.quit()  
