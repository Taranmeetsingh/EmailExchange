from flask_restful import Resource, reqparse
from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, ServiceAccount, \
    EWSDateTime, EWSTimeZone, Configuration, NTLM, GSSAPI, CalendarItem, Message, \
    Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
    HTMLBody, Build, Version, FolderCollection, Folder, folders
import json

credentials = Credentials('taranmeet.singh@healthec.com', 'welcome@789')
account = Account('taranmeet.singh@healthec.com', credentials=credentials, autodiscover=True)
qs = account.inbox.all()
tz = EWSTimeZone.localzone()
count_emails = account.inbox.total_count
qs = account.inbox.all()
tz = EWSTimeZone.localzone()
current_date = tz.localize(EWSDateTime.now())
import unreadEmails
from flask import Flask, request, jsonify
from json import dumps

searchbar = input("Please enter something to search")
q = (Q(subject__icontains=searchbar) | Q(body__icontains=searchbar) | Q(sender__icontains=searchbar))
filtered_items = qs.filter(q)

class User(Resource):

    def get(self):
        # parser = reqparse.RequestParser()
        # parser.add_argument('pagenum')
        # parser.add_argument('CountOfEmails')
        # args = parser.parse_args()
        # pagenum = args['pagenum']
        # countofemails = args['CountOfEmails']
        lst = list()
        for i in filtered_items:
            a = dict()
            a['subject'] = i.subject
            a['datetime_created'] = str(i.datetime_created)
            a['body'] = str(i.body)
            a['datetime_received'] = str(i.datetime_received)
            lst.append(a)
        jvar = json.dumps(lst)
        info = json.loads(jvar.encode("utf-8"))
        return info
        # items = {"name": "Rohit"}
        # return items

    # def set(self):
    #     parser = reqparse.RequestParser()
    #     parser.add_argument('pagenum')
    #     parser.add_argument('CountOfEmails')
    #     args = parser.parse_args()
    #     pagenum = args['pagenum']
    #     countofemails = args['CountOfEmails']
    #     items = {"name":"Rohit"}
    #     return items




