import requests
import pandas as pd
import subprocess
from pathlib import Path
from datetime import datetime, timedelta
from pyad import adquery
from exchangelib import Account, Credentials, Message, Mailbox, HTMLBody

####################################################################
# Set the API endpoint URL
url_get_list = 'https://api.wizer-training.com/api/v1/external/reports/master_report'
Url_adduser = ' https://api.wizer-training.com/api/v1/external/user'
Url_delluser = 'https://api.wizer-training.com/api/v1/external/user/by_email/'
Url_check_mail = 'https://api.wizer-training.com/api/v1/external/partner_portal/validations/email/'
# Set the API key
api_key = 'xxxxxxxxxxxxx'
csv_file = '\\\\xxx\xxx-share\AD-SAP\PRD\EMP2AD\ZHR_EMP2AD.CSV'
csv_user = 'F:\\cyber\\user.csv'
send_mailbox_username = "DailyReports@xxxx"  # for the mailbox that send the mail
send_mailbox_password = "xxxx"

####################################################################

def read_csv(csv_file):
        csv = pd.read_csv(Path(csv_file),encoding='ISO-8859-8')
        csv['תאריך תחילת עבודה'] = csv['תאריך תחילת עבודה'].apply(lambda x: '{0:0>8}'.format(x)) # fix colum that mising 0
        csv['תאריך תחילת עבודה'] = pd.to_datetime(csv['תאריך תחילת עבודה'], format='%d%m%Y')

        csv2 = csv.loc[(csv['שם קוד מצב'] == 'עזב')].copy() # fix '0' on active users
        csv2['תאריך סיום עבודה'] = csv['תאריך סיום עבודה'].apply(
            lambda x: '{0:0>8}'.format(x))  # fix colum that mising 0
        csv2['תאריך סיום עבודה'] = pd.to_datetime(csv2['תאריך סיום עבודה'], format='%d%m%Y')
        csv['תאריך סיום עבודה'] = csv2['תאריך סיום עבודה'] # merge df
        pd.set_option('display.max_columns', None)
        return csv

def user_tracking(email, csv_user):
    df = pd.read_csv(Path(csv_user))
    if email in df['user'].values:
        pass
    else:
        df2 = pd.DataFrame({"user": [email],
                            "date": [f"{datetime.now():%Y-%m-%d}"],
                            "status": ['notcompleted']})
        df3 = pd.concat([df, df2])
        df3.to_csv(Path(csv_user), index=False)

def selectday(df,numd): # filter num day users start work from csv back to list
    # df = read_csv(csv_file)
    days_ago = datetime.now() - timedelta(days=numd)
    df = df.loc[df['תאריך תחילת עבודה'] >= days_ago]
    df = df.loc[df['שם קוד מצב'] == 'פעיל']
    df = df.loc[df['זכאי ליוזר'] == 'X']
    return df.copy()

def selectday_left(df,numd): # filter num day users start work from csv back to list
    df = df.loc[df['תאריך סיום עבודה'] >= datetime.now() - timedelta(days=numd)]
    df = df.loc[df['שם קוד מצב'] == 'עזב']
    df = df.loc[df['זכאי ליוזר'] == 'X']
    return df.copy()

def send_manager_mail(email,line,username,password):  # send mail to the manager
    print(line)
    # Connect to Microsoft Exchange
    account = Account(primary_smtp_address=username,
                      credentials=Credentials(username=username, password=password),
                      autodiscover=True)

    # Create the email message
    to_address = email
    user = line['שם פרטי'] + ' ' + line['שם משפחה']
    subject = f"חשוב !! - הדרכה בנושא מודעות לעובד {user} "

    html = """
    <html>
      <body style="text-align: right;">
        <p>,מנהל יקר</p>
        <p>!בהמשך לתהליך הקליטה בחברה של העובד הוקצתה לו הדרכה בנושא מודעות. חשוב להשלים אותה בהקדם האפשרי</p>
        <p>.יש לבצע את  ההדרכה בתוך 3 ימים לאחר תקופה זו באם לא תבוצע המשימה המשתמש שלו יחסם במערכות</p>
        <br>
        <br>
        <p>לשאלות ניתן לפנות למרכז התמיכה</p>
        <br>  
      </body>
    </html>
    """
    html_body = HTMLBody(html)
    to_recipient = Mailbox(email_address=to_address)
    message = Message(account=account,
                      subject=subject,
                      body=html_body,
                      to_recipients=[to_recipient])
    message.send()
    print('send mail to: ' + email)


def send_worker_mail(email,username,password):  # send mail to the manager
    # Connect to Microsoft Exchange
    account = Account(primary_smtp_address=username,
                      credentials=Credentials(username=username, password=password),
                      autodiscover=True)

    # Create the email message
    to_address = email
    subject = f"תזכורת - הדרכה בנושא אבטחת מידע"

    html = """
    <html>
      <body style="text-align: right;">
        <p>,עובד/ת סמלת יקר/ה</p>
        <p>.בהמשך לתהליך הקליטה בחברה עליך להשלים את ההדרכה בנושא מודעות אבטחת מידע והטרדה מינית</p>
        <p> .והמשך לפי ההוראות 'cyber' אנא חפש בתיבת המייל שלך דואר שנשלח אליך משולח בשם </p>
        <p>.יש לבצע את  ההדרכה בתוך 3 ימים לאחר תקופה זו באם לא תבוצע המשימה המשתמש שלך יחסם במערכות</p>
        <br>
        <p>לשאלות ניתן לפנות למרכז התמיכה</p>
      </body>
    </html>
    """
    html_body = HTMLBody(html)
    to_recipient = Mailbox(email_address=to_address)
    message = Message(account=account,
                      subject=subject,
                      body=html_body,
                      to_recipients=[to_recipient])
    message.send()
    print('send mail to: ' + email)

def selectdayand(df,numd): # filter num day users stop work from csv back to list
    # df = read_csv(csv_file)
    days_ago = datetime.now() - timedelta(days=numd)
    df = df.loc[df['תאריך סיום עבודה'] >= days_ago]
    df = df.loc[df['שם קוד מצב'] == 'עזב']
    return df.copy()

def disable_user_AD(username):  # diable user in AD
    string = '-Replace @{info = "Script auto disabled this user by cyber"}'
    ps_command = f" Set-ADUser {username} {string}"
    subprocess.run(["powershell.exe", ps_command], capture_output=True) # write the command to user in AD
    ps_command = f"Disable-ADAccount -Identity {username}"
    # Execute the PowerShell command
    subprocess.run(["powershell.exe", ps_command], capture_output=True) # run disable command



def get_from_ad(idn): # check employeeNumber in AD end return mail
    q = adquery.ADQuery()
    q.execute_query(
        attributes=["distinguishedName", "description", "mail"],
        where_clause=f"objectClass = 'user' and employeeNumber = {idn}",
        base_dn="OU=xxx,DC=xxx,DC=local"
    )
    for row in q.get_results():
        return row["mail"]


def get_from_ad_my_mail(email): # find employeeNumber ID by mail
    q = adquery.ADQuery()
    q.execute_query(
        attributes=["distinguishedName", "description", "mail", "employeeNumber"],
        where_clause=f"objectClass = 'user' and mail = '{email}'",
        base_dn="OU=xxx,DC=xxx,DC=local"
    )
    for row in q.get_results():
        return row["employeeNumber"]

def get_from_ad_sAMAccountName(email): # find sAMAccountName ID by mail
    q = adquery.ADQuery()
    q.execute_query(
        attributes=["sAMAccountName,distinguishedName", "description", "mail", "employeeNumber"],
        where_clause=f"objectClass = 'user' and mail = '{email}'",
        base_dn="OU=xxx,DC=xxx,DC=local"
    )
    for row in q.get_results():
        return row["sAMAccountName"]


def get_data(url, api_key): # Makes a GET request to the specified URL and returns the response data.
  headers = {'apiKey': api_key}
  response = requests.get(url, headers=headers)
  return response

def ckeck_email_availability(Url_check_mail,email,api_key):
    url = Url_check_mail + email
    return get_data(url, api_key)


def post_data(url, api_key, data): # Makes a POST request to the api to get users status
  headers = {'Content-Type': 'application/json', 'apiKey': api_key}
  response = requests.post(url, json=data, headers=headers)
  return response.json()



def add_user(csvfile,response,Url_adduser,api_key,csv_user):
    ## get list of new users mail from wiz site
    rlist = []
    department = ['New_Worker02']
    for i in response.json()['userProgress']:
        rlist.append(i['email'].lower())
    slist = set(rlist)

    ## get list of new users mail from HR & AD and comper them to wiz site

    for index, row  in selectday(csvfile, 10).iterrows(): # scan for new users in last 50 days
        email = None
        email = get_from_ad(row['מספר עובד'])
        if email:
            email = email.lower()
            try:
               if email in slist: # check if thay allredy exsit
                   pass
               else:
                   user_tracking(email, csv_user) # write the user to csv file
                   print(email, 'no') # if no crate the new user
                   data = {"invites": [
                       {"email": get_from_ad(row['מספר עובד']).lower(), "departments": ['New_Worker02'], "role": "user", "name": row['שם פרטי'], "lastName": row['שם משפחה']}]}
                   try: # the data return empty so pass the error
                      result =  post_data(Url_adduser, api_key, data)
                      if 'True' in str(result.items()):
                          send_manager_mail(get_from_ad(row["מס' עובד - מנהל ישיר"]).lower(), row,
                                            send_mailbox_username,
                                            send_mailbox_password)  # send mail to the manager if no error
                          send_worker_mail(email, send_mailbox_username, send_mailbox_password)
                   except:
                       pass
            except:
                pass # no mail in ad


def disable_user_api(csvfile,Url_delluser):
    headers = {'Content-Type': 'application/json', 'apiKey': api_key}
    for index, row  in selectdayand(csvfile, 14).iterrows(): # scan for left users in last 2 days
        email = None
        email = get_from_ad(row['מספר עובד'])
        if email:
            print(email)
            Url_delluserper = Url_delluser + email.lower() + '/disable' #generet the user Url
            response = requests.patch(Url_delluserper, headers=headers)
            print(response.json())



def manage(response,send_mailbox_username,send_mailbox_password,df,csv_user):
    df = pd.read_csv(Path(csv_user))
    for index, row in df.iterrows():
        if row['status'] == 'notcompleted':
            for i in response.json()['userProgress']:
                if i['email'].lower() == row['user'].lower():
                    if i['status']!= 'Completed':
                        if datetime.strptime(df.loc[df['user'] == i['email']]['date'].values[0], '%d/%m/%Y') >= datetime.now() - timedelta(days=4): # if user did not completed he will be disabled
                            print('disable user: ' + i['email'].lower())
                            # disable_user_AD(get_from_ad_sAMAccountName(i['email'].lower())) # disable AD account
                        else:
                            send_worker_mail(i['email'].lower(), send_mailbox_username, send_mailbox_password)
                    else:
                        df.loc[df['user'] == i['email'].lower(), 'status'] = 'Completed'
    df.to_csv(Path(csv_user), index=False)




#################################################################################
csvfile = read_csv(csv_file)  #read csv file
userlist_wiz = get_data(url_get_list, api_key) #get list of users from the site api

################################################################################
# disable_user_api(csvfile,Url_delluser) # disable user from the site api
# add_user(csvfile,userlist_wiz,Url_adduser,api_key) # add user to the site api
manage(userlist_wiz,send_mailbox_username,send_mailbox_password,csvfile,csv_user) # to inform the manager only
