import email
# from email import parser
import pickle
import poplib

host = "pop.exmail.qq.com"
user = "service@fundbj.com"
pwd = "fed68318067"

if __name__ == '__main__':
    pop = poplib.POP3_SSL(host)
    pop.user(user)
    pop.pass_(pwd)

    #(邮件数, 大小bytes)
    print(pop.stat())

    messages = [pop.retr(i) for i in range(1, len(pop.list()[1]) + 1)]
    messages = [b"\r\n".join(mssg[1]) for mssg in messages]
    # messages = [parser.BytesParser().parsebytes(mssg) for mssg in messages]
    messages = [email.message_from_bytes(mssg) for mssg in messages]

    file = open('tmp.txt', 'wb')
    pickle.dump(messages, file)

    
    pop.quit()
