"""
Author: Nicholas Chenevey
Date: 11/10/2024
This script processes sent email data and generates training data for the model.
"""

import win32com.client
from datetime import datetime
import itertools
import re

def extract_emails(email_body, conversationDict, convID):
    # Define patterns to match the start of quoted replies
    patterns = [
        (r"On\s.+,\s.+\swrote:", "Pattern 1"),              # Pattern for: On [Date], [Name] wrote:
        (r"On\s.+,\s\d{4}.*<.*?>\s*wrote:", "Pattern 2"),   # Handles variations with HTML-like tags in the email
        (r"_{5,}\s*From:", "Pattern 3"),                    # Pattern for lines that start with 5 or more underscores followed by "From:"
    ]

    # Combine the patterns into a single regex with OR condition
    combined_pattern = re.compile("|".join(p[0] for p in patterns), re.MULTILINE | re.IGNORECASE)
    # Find the position of the first match, if any
    match = combined_pattern.search(email_body)
    
    if match:
        # Determine which pattern was matched
        for pattern, name in patterns:
            if re.match(pattern, email_body[match.start():], re.MULTILINE | re.IGNORECASE):
                # print(f"Matched: {name}") # Debug print
                match_format = name[-1]
                break
        if match_format == "3":
            emailBody = email_body[:match.start()].strip()
            emailData = email_body[match.start():].strip()
            conversationDict[convID]["0"]["Body"] = emailBody
            emailDict = parseConversationFrom(emailData)
            conversationDict[convID].update(emailDict)
        elif match_format in ["1", "2"]:
            emailBody = email_body[:match.start()].strip()
            emailData = email_body[match.start():].strip()
            conversationDict[convID]["0"]["Body"] = emailBody
            emailDict = parseConversationOnWrote(emailData)
            conversationDict[convID].update(emailDict)
        return conversationDict
    # If no match is found, return the full email body
    conversationDict[convID]["0"]["Body"] = email_body.strip()
    return conversationDict

def extract_field_on_wrote(data, field_name, delimiter="\r\n"):
    field_data, remaining_data = data.split(delimiter, 1)
    field_len = len(field_name)
    if field_name == "On ":
        field_len = 0
    return field_data[field_len:].strip(), remaining_data

def parseConversationOnWrote(conversation):
    email_sections = conversation.split("wrote:\r\n\r\n")
    conversation_dict = {}
    id_email = 0
    for emailInfo, emailBody in pairwise(email_sections):
        id_email += 1
        if emailInfo.strip() == "":
            id_email -= 1
            continue
        mail_dict = {}
        email_data, remaining_data = emailInfo.split(",", 1)
        mail_dict["Sent"] = email_data[3:]

        email_data = remaining_data.strip()
        mail_dict["From"] = email_data

        mail_dict["Body"] = emailBody.strip()

        conversation_dict[str(id_email)] = mail_dict
    return conversation_dict

def extract_field_from(data, field_name, delimiter="\r\n"):
    field_data, remaining_data = data.split(delimiter, 1)
    field_len = len(field_name)
    if field_name == "From: ":
        field_len = 0
    return field_data[field_len:].strip(), remaining_data

def parseConversationFrom(conversation):
    email_sections = conversation.split("From: ")
    conversation_dict = {}
    id_email = 0
    for email in email_sections:
        id_email += 1
        if email.strip() == "" or set(email.strip()) == {"_"}:
            id_email -= 1
            continue
        mail_dict = {}
        email_data, remaining_data = extract_field_from(email, "From: ")
        mail_dict["From"] = email_data

        email_data, remaining_data = extract_field_from(remaining_data, "Sent: ")
        mail_dict["Sent"] = email_data

        email_data, remaining_data = extract_field_from(remaining_data, "To: ")
        mail_dict["To"] = email_data

        email_data, remaining_data = extract_field_from(remaining_data, "Subject: ")
        mail_dict["Subject"] = email_data

        mail_dict["Body"] = remaining_data.strip()

        conversation_dict[str(id_email)] = mail_dict

    return conversation_dict

def format_date(date_string):
    # Parse the date_string into a datetime object
    date_obj = datetime.fromisoformat(date_string)
    # Format the datetime object into the desired format
    formatted_date = date_obj.strftime("%A, %B %d, %Y %I:%M %p")
    return formatted_date

def GetConversationsFromSentEmails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    numEmails = input("Enter the number of emails to process: ")

    try:
        numEmails = int(numEmails)
        if numEmails < 1:
            print("Please enter a valid number")
            quit()
    except ValueError:
        print("Please enter a valid number")
        quit()
        
    sentInbox = outlook.GetDefaultFolder(5) # "5" refers to the index of the 'sent' folder
    olItems = sentInbox.Items
    olItems.Sort("[ReceivedTime]", True)

    conversationDict = {}
    olItem = olItems.GetFirst()
    for i in range(numEmails):
        if olItem is None:
            break
        oConv = olItem.GetConversation()
        if oConv is not None:
            ConvID = oConv.ConversationID
            if ConvID not in conversationDict:
                # Get the sender's email address
                mailReply = olItem.Reply()
                mailRecipients = mailReply.Recipients[0]
                mailAddressEntry = mailRecipients.AddressEntry
                if mailAddressEntry.GetExchangeUser() is not None:
                    senderEmail = mailAddressEntry.GetExchangeUser().PrimarySmtpAddress
                else:
                    senderEmail = mailAddressEntry.Address
                sender = olItem.SenderName + " <" + senderEmail + ">"
                # Extract the time, recipients, and subject of the email
                time = format_date(str(olItem.SentOn))
                sentTo = olItem.To
                subject = olItem.Subject
                childID = 0
                conversationDict.setdefault(ConvID, {})[str(childID)] = {"From": sender, "Sent": time, "To": sentTo, "Subject": subject}
                conversationDict = extract_emails(olItem.Body, conversationDict, ConvID)
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
        for emailIDA, emailIDB in pairwise(conversationDict[conversationID]):
            emailA = conversationDict[conversationID][emailIDA]
            emailB = conversationDict[conversationID][emailIDB]
            if is_phrase_in(userName, emailA["From"]) or is_phrase_in(userEmail, emailA["From"]):
                conversationDataPrompt = []
                conversationDataCompletion = []

                roleData = "user"
                contentData = "From: '" + emailB.get("From", "N/A") + "' To: '" + emailB.get("To", "N/A") + "' Sent Date: '" + emailB.get("Sent", "N/A") + "' With subject: '" + emailB.get("Subject", "N/A") + "' With content: '" + emailB.get("Body", "N/A") + "'"
                conversationDataPrompt.append(contentDict(roleData, contentData))

                roleData = "assistant"
                contentData = emailA["Body"]
                conversationDataCompletion.append(contentDict(roleData, contentData))

                conversationDataDict["prompt"].append(conversationDataPrompt)
                conversationDataDict["completion"].append(conversationDataCompletion)
    # return conversationDataList
    return conversationDataDict

# Run the functions
userName = "USER"
userEmail = "USER EMAIL"
dataList = ConvertToDataset(GetConversationsFromSentEmails(), userName, userEmail)

with open("output.txt", "w") as f:
    f.write(str(dataList))





