def EmailExcel(sendr,password,recvr,attachmentFilePath,fileName,sub="Subject of the Mail",body="body_of_mail"):
    # libraries to be imported 
    import smtplib 
    from email.mime.multipart import MIMEMultipart 
    from email.mime.text import MIMEText 
    from email.mime.base import MIMEBase
    from email import encoders

    fromaddr = sendr#"EMAIL address of the sender"
    toaddr = recvr#"EMAIL address of the receiver"

    msg = MIMEMultipart() 

    msg['From'] = fromaddr 
    msg['To'] = toaddr

    msg['Subject'] = sub
    body = body
    msg.attach(MIMEText(body, 'plain')) 

    # open the file to be sent  
    filename = fileName #"File_name_with_extension"
    attachment = open(attachmentFilePath, "rb") #open("Path of the file", "rb") 

    p = MIMEBase('application', 'octet-stream') 
    p.set_payload((attachment).read()) 
    encoders.encode_base64(p) 

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 

    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 

    s = smtplib.SMTP('smtp.gmail.com:587')
    s.ehlo()
    s.starttls() 
    # Authentication 
    s.login(fromaddr, password)
    text = msg.as_string() 
    s.sendmail(fromaddr, toaddr, text) 
    s.quit() 

    
    
def EmailUpdate(sendr,password,recvr,sub="Subject of the Mail",body="body_of_mail"):
    # libraries to be imported 
    import smtplib 
    from email.mime.multipart import MIMEMultipart 
    from email.mime.text import MIMEText 
    from email.mime.base import MIMEBase
    from email import encoders

    fromaddr = sendr#"EMAIL address of the sender"
    toaddr = recvr#"EMAIL address of the receiver"

    msg = MIMEMultipart() 

    msg['From'] = fromaddr 
    msg['To'] = toaddr

    msg['Subject'] = sub
    body = body
    msg.attach(MIMEText(body, 'plain')) 

    
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.ehlo()
    s.starttls() 
    # Authentication 
    s.login(fromaddr, password)
    text = msg.as_string() 
    s.sendmail(fromaddr, toaddr, text) 
    s.quit()
    
    
def SqlToExcel(sqlQuery,excelFile,sendMail=0,database="Output"):
    import pyodbc
    import pandas as pd
    
    try :
        pyodbc.autocommit = True
        
        conn = pyodbc.connect("DSN=hive;UID=userID;PWD=Qwerty@1234",autocommit=True)
        
        # Update Mail
        body = f"Connection established : for the query : {sqlQuery} "
        sub="Update for Auto query"
        EmailUpdate(sendr,password,sendr,sub,body)

        try:
            result = pd.read_sql(sqlQuery,conn)
            
            try:
                df=pd.DataFrame(result)
                df.to_csv(excelFile,index=False)

                print(f"Results are saved to file : {excelFile} : for the query : {sqlQuery} ")
            
            
                if sendMail ==1:
                    import os
                    attachmentFilePath = os.path.join(os.getcwd(),excelFile)

                    EmailExcel(sendr,password,recvr,attachmentFilePath,excelFile,sub_Mail,body_Mail)

                else:
                    # Update Mail
                    body = f"Results are saved to file : {excelFile} : for the query : {sqlQuery} "
                    sub="Update for Auto query"
                    EmailUpdate(sendr,password,sendr,sub,body)
                    
            except:
                # Update Mail
                body = f"Result could not be saved in :{excelFile} for the : {sqlQuery} "
                sub="Alert : Update for Auto query"
                EmailUpdate(sendr,password,sendr,sub,body)
            
                    
        except:
            # Update Mail
            body = f"Query incorrect : {sqlQuery} "
            sub="Alert : Update for Auto query"
            EmailUpdate(sendr,password,sendr,sub,body)
            
            print("Query incorrect : " + sqlQuery)

        finally:
            conn.close()
    except:
        # Update Mail
        body = f"Connection NOT established : for the query : {sqlQuery} "
        sub="Alert : Update for Auto query"
        EmailUpdate(sendr,password,sendr,sub,body)

        print("Database connection not established : ")
        
        
def InitialMailData():
    workbookname = "AutomatedSqlRunner.xlsx"
    sheetMail = "MailDetails"
    sheetSql = "SQLQuery"

    global sendr,password,recvr,sub_Mail,body_Mail, no_sqlQuery
    import xlrd

    workbook = xlrd.open_workbook(workbookname)
    worksheet = workbook.sheet_by_name(sheetMail)
    num_rows = worksheet.nrows - 1
    
    sendr = (worksheet.cell(0, 1)).value
    password = (worksheet.cell(1, 1)).value
    recvr = (worksheet.cell(2, 1)).value
    sub_Mail = (worksheet.cell(3, 1)).value
    body_Mail = (worksheet.cell(4, 1)).value
    
    no_sqlQuery = (workbook.sheet_by_name(sheetSql)).nrows-1
    return no_sqlQuery
    
def InitialSQL():
    workbookname = "AutomatedSqlRunner.xlsx"
    sheetMail = "MailDetails"
    sheetSql = "SQLQuery"

    import xlrd

    workbook = xlrd.open_workbook(workbookname)
    worksheet = workbook.sheet_by_name(sheetSql)
    num_rows = worksheet.nrows - 1
    curr_row = 0 # As first row is to be excluded
    
    import threading
    list_thread=[]
    while curr_row < num_rows:
            curr_row += 1
            filename = (worksheet.cell(curr_row,0)).value
            sendMail = (worksheet.cell(curr_row,1)).value
            sqlQuery = (worksheet.cell(curr_row,2)).value
            
            t = threading.Thread(target=SqlToExcel, args=(sqlQuery,filename,sendMail,))
            list_thread.append(t)
    return list_thread
    