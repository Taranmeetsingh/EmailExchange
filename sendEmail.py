import xlwt
import datetime
import timedelta

from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version, FolderCollection, Folder, folders


credentials = Credentials('taranmeet.singh@healthec.com', 'welcome@789')
account = Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True)

qs = account.inbox.all()

m = Message(
    account=Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True),
    subject='Test email',
    body='Taran yikes beongly dow',
    to_recipients=[
        Mailbox(email_address='taranmeet_singh@mckinsey.com')
    ],
    cc_recipients=['taranmeet.oberoi28@gmail.com'],  # Simple strings work, too
    bcc_recipients=[
        Mailbox(email_address='toberoi@trackbackinc.com')
    ]
)
m.send()