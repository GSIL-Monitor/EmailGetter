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
