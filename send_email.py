import zmail 
import xlrd 

def read_excel( filepath) :
    workbook = xlrd.open_workbook(filepath)
    sheet = workbook.sheets()[0]
    return  sheet.get_rows()


if __name__ == '__main__':
    
    print("\033[7;37;40m请输入你的邮箱地址\033[0m")
    source_email = input()
    print("\033[7;37;40m请输入你的邮箱密码\033[0m")
    source_token = input()
    print("\033[7;37;40m请输入成绩单excel的绝对地址（应为xls格式，第一列为姓名，第二列为综合名次，第三列为成绩名次，第四列为学号)\033[0m") 
    filepath = input()

    server = zmail.server(source_email , source_token)
    # server = zmail.server('email_address' , 'email_password')
    if server.smtp_able():
        print (" SMTP service ready")
        pass
    else : 
        print ("你的邮箱没有开通smtp服务，可能发送失败")
    
    rows = read_excel(filepath)

    for row in rows :
        if row[0].value == '' or row[0].value =='姓名' or row[0].value =='NAME' :
            continue 
        name = str(row[0].value)    # get name
        integrative_rank = str(int(row[1].value)) # get integrative_rank
        score_rank = str(int(row[2].value)) # get score_rank
        if row[3].ctype == 2 and row[3].value % 1 == 0.0 : #get email_address 
            email = str(int(row[3].value))
        else :
            email = str(row[3].value)


        Subject = name + "的成绩排名"
        Content_text = "综合成绩排名为" + integrative_rank + "" + '\n' + \
                       "成绩排名为" + score_rank + "" + '\n'
        
        mail = {
            'subject' : Subject ,
            'content_text' : Content_text ,
        }     
        server.send_mail(email +'@buaa.edu.cn',mail)

    print("your task is finished ---\nCopyRight by : Dormosue")
    
    