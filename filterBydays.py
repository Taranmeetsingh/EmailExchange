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

tz = EWSTimeZone.localzone()
current_date = tz.localize(EWSDateTime.now())
last30days = current_date - datetime.timedelta(days=30)

items_for_last30days = qs.filter(datetime_received__range=(
    last30days,
    current_date
))
count_emails = items_for_last30days.count()
print(count_emails)
for item in items_for_last30days:
    print(item.subject, item.body, item.datetime_created, item.datetime_received, item.sender)
