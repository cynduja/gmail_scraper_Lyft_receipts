import imaplib
import email
import re
import csv
import sys
import usaddress
import collections

#Reading username and pwd from an external file
file_object = open("usernamenpwd.txt", "r")
username = file_object.readline()
password = file_object.readline()

#Connecting to gmail
imap_server = imaplib.IMAP4_SSL("imap.gmail.com", 993)
#print imap_server
imap_server.login(username, password)

#Select a folder called lyft receipts where all our receipts are stored
# status, response = imap_server.select('uber receipts')
status, response = imap_server.select('lyft receipts')
#imap_server.select returns a status and the total number of messages in the folder
#print response

#Get all the email_ids of all the messages in that folder
status, email_ids = imap_server.search(None, "ALL")
#print email_ids[0]

#Create an array of email_ids
new_email_ids = email_ids[0].split()
# print new_email_ids

#We want to decode it to unicode string
def get_decoded_email_body(message_body):
	 msg = message_body
	 text = ""
	 if msg.is_multipart():
	 	html = None
	 	for part in msg.walk():
			charset = part.get_content_charset()
	 		if part.get_content_type() == 'text/plain':
	 			text = unicode(part.get_payload(decode=True), str(charset), "ignore").encode('utf8', 'replace')
 			if part.get_content_type() == 'text/html':
 				html = unicode(part.get_payload(decode=True), str(charset), "ignore").encode('utf8', 'replace')
 		if text is not None:
 			return text.strip()
 		else:
 			return html.strip()
 	 else:
 	 	text = unicode(msg.get_payload(decode=True), msg.get_content_charset(), 'ignore').encode('utf8', 'replace')
 	 	return text.strip()


def get_trip_data(data):
	trip_data = re.search(r'Ride (\d+\.*\d+) \w\w \W *([0-9]*) * [a-z]*', data)
	if trip_data:
		miles_travelled = trip_data.group(1)
		time_spent = trip_data.group(2)
		dollars_spent = re.search('\d: * \$(\d+).*', data)
		if dollars_spent:
			return miles_travelled,time_spent, dollars_spent.group(1)
		else:
			dollars=dollars_spent.group(1)
			dollars="NA"
			return miles_travelled, time_spent, dollars 
	else: 
		miles_travelled="NA"
		time_spent="NA"
		dollars_spent = re.search('\d: * \$(\d+).*', data)
		if dollars_spent:
			return miles_travelled,time_spent, dollars_spent.group(1)
		else:
			dollars=dollars_spent.group(1)
			dollars="NA"
			return miles_travelled, time_spent, dollars 	

def get_pickup(data):
	f = open("msg1.html", "w")
	f.write(data)
	f.close()
	infile = open("msg1.html","rb")
	stripped_data = [i.strip() for i in infile.readlines()]
	for i,d in enumerate(stripped_data):
		if 'Pickup:' in d:
			if'Dropoff:' not in stripped_data[i+1]:
				pickup = d + " " + stripped_data[i+1]
				return pickup.replace("Pickup:",'').lstrip()
			else:
				pickup = d
				return pickup.replace("Pickup:",'').lstrip()

def get_dropoff(data):
	f = open("msg1.html", "w")
	f.write(data)
	f.close()
	infile = open("msg1.html","rb")
	stripped_data = [i.strip()for i in infile.readlines()]
	for i, d in enumerate(stripped_data):
		if 'Dropoff:' in d:
			if'Ride'not in stripped_data[i+1] and'Lyft'not in stripped_data[i+1]:
				dropoff = d + "" + stripped_data[i+1]
				return dropoff.replace("Dropoff:", '').lstrip()
			else:
				dropoff = d
				return dropoff.replace("Dropoff:", '').lstrip()

def get_datetime(data):
	datetime = re.search(r'([a-zA-Z]+ +\d+ )\w\w (\d+:\d+ [APM]{2})', data)
	return datetime.group(1).split()[0], datetime.group(1).split()[1], datetime.group(2)

def parse_address(address):
	parsed_address, addresstype = usaddress.tag(address)
	missing_keys = whats_missing(parsed_address)
	if len(missing_keys)>0:
		for m in missing_keys:
			parsed_address.update({m:'NA'})
		return parsed_address['PlaceName'], parsed_address['StateName'], parsed_address['ZipCode'],parsed_address['CountryName']
	else:
		return parsed_address['PlaceName'], parsed_address['StateName'], parsed_address['ZipCode'],parsed_address['CountryName']



def whats_missing(addressdict):
	missing=[]
	keys = addressdict.keys()
	if 'PlaceName' and 'StateName' and 'ZipCode'and 'CountryName' in keys:
		return missing
	else:
		if 'PlaceName' not in keys:
			missing.append('PlaceName')
		if 'StateName' not in keys:
			missing.append('StateName')
		if 'CountryName' not in keys:
			missing.append('CountryName')
		if 'ZipCode' not in keys:
			missing.append('ZipCode')
		return missing


#Open a file with titles even before we get the data
f=open('lyft_gmail_final.csv','wt')
writer=csv.writer(f)
writer.writerow(("miles_travelled","time","dollars_spent","pickup_address", "pickup_city", "pickup_state", "pickup_zip", "pickup_country","dropoff_address", "dropoff_city", "dropoff_state", "dropoff_zip", "dropoff_country", "trip_month", "trip_date", "trip_time"))

for num in new_email_ids:
	rv, data = imap_server.fetch(num, '(RFC822)')
	msg = email.message_from_string(data[0][1])
	data = get_decoded_email_body(msg)
	miles_travelled,time,dollars_spent=get_trip_data(data)
	pickup_address=get_pickup(data)
	pickup_city, pickup_state, pickup_zip, pickup_country = parse_address(pickup_address)
	dropoff_address=get_dropoff(data)
	dropoff_city, dropoff_state, dropoff_zip, dropoff_country = parse_address(dropoff_address)
	trip_month, trip_date, trip_time = get_datetime(data)
	writer.writerow((miles_travelled,time,dollars_spent, pickup_address, pickup_city, pickup_state, pickup_zip, pickup_country,dropoff_address, dropoff_city, dropoff_state, dropoff_zip, dropoff_country, trip_month, trip_date, trip_time))
