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

my_folder = Folder(name='Last30days', parent=account.inbox)

# [uncomment if new folder to be created]
# Folder name for last 30 days - 'Last30days'
# Folder name for last 30 days - 'Last60days'
# Folder name for last 30 days - 'Last90days'

# my_folder.save()

for tm in items_for_last30days:
    tm.copy(to_folder=account.inbox/'Last30days')

# print(account.root.tree())

fold_emails = account.inbox/'Last30days'

for item in fold_emails.all():
    print(item.subject, item.body, item.datetime_created, item.datetime_received, item.sender)
    print("#" * 40)

