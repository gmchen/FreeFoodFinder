#!/usr/bin/env python
########################################################################################################
#
# QueryEmailForFood.py
# Author: Gregory M Chen
# 
# Description: Log into an IMAP4 email server, fetch unread email, and search the text for words
# which may be related to an event involving free food. Output a report of candidate emails.
#
#
########################################################################################################

import imaplib
import email
import re
import datetime
import quopri

regexes = ['FOOD', 'PIZZA', 'WILL BE PROVIDED', 'WILL ALSO BE PROVIDED', 'SNACK', 'REFRESHMENT', 'DRINKS', 'COOKIES', 'CUPCAKE', 'COFFEE', 'BBQ', 'POTLUCK', '(?<!FEEL )FREE(?!(DOM| SPEECH| THE CHILDREN| FOR MEMBERS))']

## A text file with two lines: username and password
f = open('creds.txt')
lines = f.readlines()
f.close()

username = lines[0].rstrip()
password = lines[1].rstrip()

print "Connecting to Gmail server..."
mail = imaplib.IMAP4_SSL('imap.gmail.com')
print "Connected. Logging in..."
mail.login(username, password)
print "Logged in. Fetching unread emails..."
mail.select('inbox')
result, data = mail.uid('search', None, 'UnSeen')
#result, data = mail.uid('search', None, 'ALL')
ids = data[0].split()

email_texts = [];
email_senders = []

for i in ids:
	result, data = mail.uid('fetch', i, '(RFC822)')
	raw_email = data[0][1]
	email_message = email.message_from_string(raw_email)
	email_senders.append(email_message['From'])
	text = "SUBJECT: " + email_message['subject'] + "\n";
	text = text + "BODY TEXT: "
        # The following snippet is from http://stackoverflow.com/a/1463144
	for part in email_message.walk():
		# each part is a either non-multipart, or another multipart message
		# that contains further parts... Message is organized like a tree
		if part.get_content_type() == 'text/plain':
			text = text + str(part.get_payload()) + "\n"
	text = text + "\n"
	email_texts.append(text)

# clean up the text a bit
for i in range(len(email_texts)):
	email_texts[i] = quopri.decodestring(email_texts[i])
	email_texts[i] = email_texts[i].replace("*", "")

to_keep = []
upper_texts = []
for i in range(len(email_texts)):
	to_keep.append(False)
	upper_texts.append(email_texts[i].upper())
print "Searching for regexes..."
# Keep only emails that match at least one regex
for i in range(len(upper_texts)):
	for reg in regexes:
		match = re.search(reg, upper_texts[i])
		if(match):
			start = match.start()
			length = len(match.group())
			email_texts[i] = email_texts[i][:start] + '\\cf2 ' + email_texts[i][start:start+length] + '\\cf1 ' + email_texts[i][start+length:]
			upper_texts[i] = email_texts[i].upper() # position is shifted now
			to_keep[i] = True
print "Writing to file..."
# Figure out how many were added
num_added = 0
for i in range(len(to_keep)):
	if(to_keep[i] == True):
		num_added = num_added + 1
filename = datetime.datetime.now().strftime("Emails/Emails_%Y-%m-%d_%H%M_") + str(num_added) + ".rtf"
myfile = open(filename, 'w')
myfile.write("{\\rtf1\\ansi\\deff0\n{\\colortbl;\\red0\\green0\\blue0;\\red255\\green0\\blue0;}\\cf1 ")
myfile.write("-------------------------Food Report-------------------------\line\line\line ")
for i in range(len(to_keep)):
	if(to_keep[i]):
		email_text = email_texts[i].replace('\r\n', '\n') # \line is for rtf
		email_text = re.sub("(\n|\n )+", "\n", email_text)
		email_text = email_text.replace('\n', '\\line ') # \line is for rtf
		myfile.write("--------------------" + email_senders[i] + "--------------------\line " + email_text + "\line\line\line ")

myfile.write("\n}")

myfile.close()

if(num_added == 1):
	print "Completed successfully. 1 email out of " + str(len(email_texts)) + " was found to be promising."
else:
	print "Completed successfully. " + str(num_added) + " emails out of " + str(len(email_texts)) + " were found to be promising."
