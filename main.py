import pandas as pd
import datetime 
import smtplib#used to read today's date
#Reading excel sheet daata
#enter your authentication details
GMAIL_ID='sinchanaac00@gmail.com'
GMAIL_PSWD='sinch151'
def sendemail(to,sub,msg):
    print(f"Email sent to {to} sent with subject:{sub} and message {msg}")
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_PSWD,to,f"Subject:{sub}\n\n{msg}")
    s.quit()
    




if __name__=="__main__":
    sendemail(GMAIL_ID,"subject","test message")
    df=pd.read_excel("main.xlsx")#REads the date saved excel sheet
    today=datetime.datetime.now().strftime("%d-%m")#takes today's date as date and month
    #print(today)#It is in string  format
    writeind=[]
    for index,item in df.iterrows(): #iterates over dataframe in python
        #print(index,item['Birthday'])#Gives the  index for birthday column in excel sheet
        bday=item['Birthday'].strftime("%d-%m")
        yearnow=datetime.datetime.now().strftime("%Y")
        #print(bday)
        if(today==bday) and yearnow not in str(item['Year']):
            sendemail(item['Email'],"HAPPY BIRTHDAY",item['Dialogue'])
            writeind.append(index)
    #print(writeind)
    for i in writeind:
        yr=df.loc[i,'Year']
        df.loc[i,'Year']=str(yr)+',' +str(yearnow)#change in year column
        #print(df.loc[i,'Year'])
    #print(df)
    df.to_excel('main.xlsx',index=False)
        
            
        
    
    