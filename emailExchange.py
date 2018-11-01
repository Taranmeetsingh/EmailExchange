# import export as export
import json
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version, FolderCollection, Folder, folders

credentials = Credentials('taranmeet.singh@healthec.com', 'welcome@789')
account = Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True)

# for item in account.inbox.all().order_by('-datetime_received')[:10]:
#     print(item.subject, item.sender, item.datetime_received)

count_emails = account.inbox.total_count
# print(count_emails)

qs = account.inbox.all()
searchbar = input("Please enter something to search")
q = (Q(subject__icontains=searchbar) | Q(body__icontains=searchbar) | Q(sender__icontains=searchbar))
filtered_items = qs.filter(q)  # Returns items where subject is exactly 'foo'. Case-sensitive
# items.save()

countemails = filtered_items.count()
print(countemails)

# qs.filter(start__range=(dt1, dt2))  # Returns items within range
# items = qs.filter(subject__in=searchbar)  # Return items where subject is either 'foo' or 'bar'

# print(items)
# print(qs.filter(subject=searchbar))
# for item in items:
    # print(item.subject, item.body, item.datetime_created, item.datetime_received, item.item_id, item.message_id)
lst = list()
pageNum = input("Enter Page number")

for i in filtered_items:
    a = dict()
    a["subject"] = i.subject
    a["datetime_created"] = str(i.datetime_created)
    a['body'] = str(i.body)
    a['datetime_received'] = str(i.datetime_received)
    lst.append(a)
print(json.dumps(lst))
# for each in items:
#     tot = export(each)


# folder = account.root
# for p in folder.people():
#     print(p)
# for p in folder.people().only('display_name').filter(display_name='naveen').order_by('display_name'):
#     print(p)