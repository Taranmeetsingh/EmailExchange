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

# for tm in qs.filter(is_read=False, datetime_received__gt=last30days):
#     tm.move(to_folder=account.Last30days)
# printemails = account.folders(Last30days.all())
# print(printemails)

searchbar = input("Please enter something to search")
# items = account.folders.get(last30days).all()  # Get everything
# items = Last30days.all().iterator()  # Get everything, but don't cache
items = qs.filter(subject__icontains=searchbar)  # Returns items where subject is exactly 'foo'. Case-sensitive
# sub_filter = qs.filter(subject__contains=searchbar)
# body_filter = qs.filter(body__contains=searchbar)
# items = {}
# if qs.filter(body__contains=searchbar) is not None:
#     items = items.append(body_filter)
# if qs.filter(subject__contains=searchbar) is not None:
#     items = items.append(sub_filter)
# if qs.filter(subject__icontains=searchbar) is not None:
#     items = items.append(qs.filter(subject__icontains=searchbar))
# if qs.filter(body__icontains=searchbar) is not None:
#     items = items.append(qs.filter(body__icontains=searchbar))
# if qs.filter(subject__startswith=searchbar) is not None:
#     items = items.append(qs.filter(subject__startswith=searchbar))
# if qs.filter(body__startswith=searchbar) is not None:
#     items = items.append(qs.filter(body__startswith=searchbar))
# if qs.filter(sender__icontains=searchbar) is not None:
#     items = items.append(qs.filter(sender__icontains=searchbar))

# qs.filter(start__range=(dt1, dt2))  # Returns items within range
# items = items.filter(subject__icontains=searchbar)  # Return items where subject is either 'foo' or 'bar'
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

for item in items:
    print(item.subject, item.body, item.datetime_created, item.datetime_received, item.sender)
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
