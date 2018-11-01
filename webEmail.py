import export as export
from exchangelib import Credentials, Account
from weboutlook import OutlookWebScraper

# credentials = Credentials('taranmeet.singh@healthec.com', 'welcome@789')
# account = Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True)
s = OutlookWebScraper('https://mail.healthec.com', 'taranmeet.singh@healthec.com', 'welcome@789')
s.login()

# for item in account.inbox.all().order_by('-datetime_received')[:10]:
#     print(item.subject, item.sender, item.datetime_received)

count_emails = account.inbox.total_count
print(count_emails)

s.inbox()

qs = account.inbox.all()
searchbar = input("Please enter something to search")
items = qs.filter(subject__icontains=searchbar)  # Returns items where subject is exactly 'foo'. Case-sensitive

# qs.filter(start__range=(dt1, dt2))  # Returns items within range
# items = qs.filter(subject__in=searchbar)  # Return items where subject is either 'foo' or 'bar'

print(items)
print(qs.filter(subject=searchbar))
for item in items:
    print(item.subject, item.body, item.datetime_created, item.datetime_received, item.item_id, item.message_id)

# for each in items:
#     tot = export(each)


# folder = account.root
# for p in folder.people():
#     print(p)
# for p in folder.people().only('display_name').filter(display_name='naveen').order_by('display_name'):
#     print(p)