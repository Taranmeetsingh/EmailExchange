import xlwt
import datetime
import timedelta
import json

from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version, FolderCollection, Folder, folders


credentials = Credentials('taranmeet.singh@healthec.com', 'welcome@789')
account = Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True)

qs = account.inbox.all()
# my_folder = Folder(name='Last90days', parent=account.inbox) - folders have been created [uncomment if new folder to
#  be created]
# Folder name for last 30 days - 'Last30days'
# Folder name for last 30 days - 'Last60days'
# Folder name for last 30 days - 'Last90days'

# my_folder.save()

# a = Account(...)
tz = EWSTimeZone.localzone()
current_date = tz.localize(EWSDateTime.now())
last30days = current_date - datetime.timedelta(days=30)

items_for_last30days = qs.filter(datetime_received__range=(
    last30days,
    current_date
))

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

# print(items)
# print(qs.filter(body=searchbar))

sheet1.write(0, 0, "Subject")
sheet1.write(0, 1, "Sender")
sheet1.write(0, 2, "Date")
sheet1.write(0, 3, "Mail Body")

i = 1
j = 0
# for item in items_for_last30days.order_by('-datetime_received'):
for item in items_for_last30days.order_by('sender'):
    # print(item.subject, item.body, item.datetime_created, item.datetime_received, item.sender)
    j = 0
    sheet1.write(i, j, item.subject)
    j += 1
    sheet1.write(i, j, item.sender.email_address)
    j += 1
    datetime = str(item.datetime_received)
    sheet1.write(i, j, datetime)
    j += 1
    mailbody = str(item.body)
    sheet1.write(i, j, mailbody)
    i += 1


book.save("export_" + "last30days" + ".xls")
# import os
# # os.startfile('export_last30days.xls')