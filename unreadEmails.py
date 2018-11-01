# import xlwt
# import datetime
# import timedelta
import json

from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version, FolderCollection, Folder, folders

credentials = Credentials('taranmeet.singh@healthec.com', 'welcome@789')
account = Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True)


count_unread_emails = account.inbox.unread_count
print(count_unread_emails)

qs = account.inbox.all()
items = qs.filter(is_read=False)  # Returns items where subject is exactly 'foo'. Case-sensitive


# for item in items:
#     # print(item.subject, item.body, item.datetime_created, item.datetime_received, item.item_id, item.message_id)
#     json.dumps(item)
#     print(json)
lst = list()

for i in items:
    a = dict()
    a["subject"] = i.subject
    a["datetime_created"] = str(i.datetime_created)
    a['body'] = str(i.body)
    a['datetime_received'] = str(i.datetime_received)
    lst.append(a)
print(json.dumps(lst))
