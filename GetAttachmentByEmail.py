# -*- coding: utf-8 -*-
import os
import win32com.client


folder = '//mnacpfs01/GPI-DS/Zach Olivier/DealerTrack/Emails'
file = os.listdir(folder)


#print(msg.SenderName)
#print(msg.SenderEmailAddress)
#print(msg.SentOn)
#print(msg.To)
#print(msg.CC)
#print(msg.BCC)
#print(msg.Subject)
#print(msg.Body)

for t in range(len(file)):
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    msg = outlook.OpenSharedItem(folder+'/'+file[t])
    
    count_attachments = msg.Attachments.Count
    name = []
    if count_attachments > 0:
        for item in range(count_attachments):
            name.append(msg.Attachments.Item(item + 1).Filename)
            
            att = msg.Attachments
            for i in att:
                i.SaveAsFile(folder+'/'+msg.Attachments.Item(item + 1).Filename)
                
    del outlook, msg



