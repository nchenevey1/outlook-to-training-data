"""
Author: Nicholas Chenevey
Date: 10/08/2024

This script processes email conversations from Microsoft Outlook, parses them into individual emails, 
and converts them into a dataset suitable for training machine learning models. The dataset includes 
prompts and completions based on the email content, sender, recipient, and other metadata.
"""


import win32com.client
from datetime import datetime
import itertools
import re

def parseConversation(conversation):
    """
    Parses an email conversation string into a dictionary of individual emails.
    Args:
        conversation (str): The email conversation string to be parsed.
    Returns:
        dict: A dictionary where each key is an email ID (starting from 0) and each value is a dictionary 
              containing the parsed details of the email with the following keys:
              - "From": The sender of the email.
              - "Sent": The date and time the email was sent.
              - "To": The recipient(s) of the email.
              - "Subject": The subject of the email.
              - "Body": The body content of the email.
    """
    email_sections = conversation.split("From: ")
    email_sections.pop(0)
    idEmail = 0
    conversationDict = {}

    for email in email_sections:
        mailDict = {}
        emailData = email.split("\r\n", 1)
        mailDict["From"] = emailData[0]

        emailData = emailData[1].split("To: ", 1)
        mailDict["Sent"] = emailData[0][6:-5]

        emailData = emailData[1].split("\r\n", 1)
        mailDict["To"] = emailData[0]

        emailData = emailData[1].split("\r\n", 1)
        mailDict["Subject"] = emailData[0][9:]

        mailDict["Body"] = emailData[1]

        conversationDict[idEmail] = mailDict
        idEmail += 1

    return conversationDict

def format_date(date_string):
    # Parse the date_string into a datetime object
    date_obj = datetime.fromisoformat(date_string)
    # Format the datetime object into the desired format
    formatted_date = date_obj.strftime("%A, %B %d, %Y %I:%M %p")
    return formatted_date

def GetConversations():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    numEmails = int(input("Enter the number of emails to process: "))

    if numEmails is None or numEmails < 1:
        print("Please enter a valid number")
        quit()
        
    sentInbox = outlook.GetDefaultFolder(5) # "5" refers to the index of a folder sent in this case
    olItems = sentInbox.Items

    conversationDict = {}
    olItem = olItems.GetFirst()
    for i in range(numEmails):
        if olItem is None:
            break
        oConv = olItem.GetConversation()
        if oConv is not None:
            if oConv.ConversationID not in conversationDict:
                childID = 1
                for childItem in oConv.GetChildren(olItem):
                    mailReply = childItem.Reply()
                    mailRecipients = mailReply.Recipients[0]
                    mailAddressEntry = mailRecipients.AddressEntry
                    if mailAddressEntry.GetExchangeUser() is not None:
                        senderEmail = mailAddressEntry.GetExchangeUser().PrimarySmtpAddress
                    else:
                        senderEmail = mailAddressEntry.Address
                    sender = "From: " + childItem.SenderName + " <" + senderEmail + ">\r\n"
                    time = "Sent: " + format_date(str(childItem.SentOn)) + "\r\n"
                    sentTo = "To: " + childItem.To + "\r\n"
                    subject = "Subject: " + childItem.Subject + "\r\n \r\n"
                    childBody = sender + time + sentTo + subject + childItem.Body
                    conversationDict.setdefault(oConv.ConversationID, {})[str(childID)] = parseConversation(childBody)
                    childID += 1
        olItem = olItems.GetNext()
    
    return conversationDict

def pairwise(iterableObject):
    """
    Generate pairs of consecutive elements from the given iterable.
    Returns:
        iterator: An iterator of tuples, where each tuple contains a 
        pair of consecutive elements from the input iterable.
    Example:
        list(pairwise([1, 2, 3, 4]))
        [(1, 2), (2, 3), (3, 4)]
    """
    a, b = itertools.tee(iterableObject)
    next(b, None)
    return zip(a, b)

def contentDict(role, content):
    return {"role": role, "content": content}

def is_phrase_in(phrase, text):
    """
    Check if a phrase is present in a given text as a whole word, case insensitive.
    Returns bool: True if the phrase is found in the text as a whole word, False otherwise.
    """
    return re.search(r'\b' + re.escape(phrase) + r'\b', text, re.IGNORECASE)

def ConvertToDataset(conversationDict, userName, userEmail):
    conversationDataDict = {"prompt": [], "completion": []}
    for conversationID in conversationDict:
        for childID in conversationDict[conversationID]:
            for emailIDA, emailIDB in pairwise(conversationDict[conversationID][childID]):
                emailA = conversationDict[conversationID][childID][emailIDA]
                emailB = conversationDict[conversationID][childID][emailIDB]
                if is_phrase_in(userName, emailA["From"]) or is_phrase_in(userEmail, emailA["From"]):

                    conversationDataPrompt = []
                    conversationDataCompletion = []

                    #roleData = "system"
                    #contentData = "You are an Outlook assistant writing emails for " + userName #" writing to " + emailA["To"] + " in response to an email sent on " + emailB["Sent"] + " with the subject " + emailB["Subject"]
                    #conversationData.append(contentDict(roleData, contentData))

                    roleData = "user"
                    contentData = "From: '" + emailB["From"] + "' To: '" + emailB["To"] + "' Sent Date: '" + emailB["Sent"] + "' With subject: '" + emailB["Subject"] + "' With content: '" + emailB["Body"] +"'"
                    conversationDataPrompt.append(contentDict(roleData, contentData))

                    roleData = "assistant"
                    contentData = emailA["Body"]
                    conversationDataCompletion.append(contentDict(roleData, contentData))

                    conversationDataDict["prompt"].append(conversationDataPrompt)
                    conversationDataDict["completion"].append(conversationDataCompletion)

    # return conversationDataList
    return conversationDataDict

# Run the functions

userName = "USER NAME"
userEmail = "USER EMAIL"
dataList = ConvertToDataset(GetConversations(), userName, userEmail)

with open("output.txt", "w") as f:
    f.write(str(dataList))



