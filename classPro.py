class Config:
    from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
        EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
        Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
        HTMLBody, Build, Version, FolderCollection, Folder, folders
    credentials = Credentials('taranmeet.singh@healthec.com', 'welcome@789')
    account = Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True)
    qs = account.inbox.all()
    tz = EWSTimeZone.localzone()
    count_emails = account.inbox.total_count
    qs = account.inbox.all()
    tz = EWSTimeZone.localzone()
    current_date = tz.localize(EWSDateTime.now())
    last30days = current_date - datetime.timedelta(days=30)
    last60days = current_date - datetime.timedelta(days=60)
    last90days = current_date - datetime.timedelta(days=90)


class EmailExchange:
    def __init__(self, name, age):
        self.name = name
        self.age = age
    searchbar = input("Please enter something to search")
    q = (Q(subject__icontains=searchbar) | Q(body__icontains=searchbar) | Q(sender__icontains=searchbar))
    filtered_items = qs.filter(q)
    for i in filtered_items:
        a = dict()
        a["subject"] = i.subject
        a["datetime_created"] = str(i.datetime_created)
        a['body'] = str(i.body)
        a['datetime_received'] = str(i.datetime_received)
        lst.append(a)
    print(json.dumps(lst))

  p1 = Person("John", 36)

    def