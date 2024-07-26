import win32com.client 
from time import sleep
import datetime
# import win32ui
 
print("Please ensure you have outlook running prior to opening the tool, if you dont have it running then reopen the tool with outlook running!")

# validation = win32ui.FindWindow(None, "Outlook")	#Check if outlook is running

# if validation:
	# print("Yeah")


dispatcher = win32com.client.Dispatch('Outlook.Application')

mailsize = 0x0	#size of bytes upon initialisation

newmail = dispatcher.CreateItem(mailsize)

newmail.Subject = input("What is the subject of the mail")	#subject of the mail

newmail.To = input("Who will be recieving the mail")	#Who is recieving the mail

CC = input("Would you like to CC anyone (GRC already CC'd)")
if "y" in CC:
	newmail.CC = input("Who would you like to CC?")

BCC = input("Would you like anyone to be included in the BCC?")
if "y" in BCC:
	newmail.BCC = input("Who would you like to BCC?")

newmail.body = ""	#to be appended to depending on the final function

attach = ""	#if files are to be attached, experiemental and probably should be taken out

today = datetime.datetime.strptime(start_date, "%d/%m/%y")

fifteen_days_later = today + datetime.timedelta(days=15)

#Use pyinstaller to get the things into one file
def request():
	vendor = input("Enter the vendors name")

	#Vendors name
	contact = input("\nPlease enter the name of the vendor\n")

	#Internal Entity name; Can be set to a constant or predefined if you want
	internal = input("\nPlease enter the name of the internal entity you are requesting this for\n")

	assessorreassess = input("Is this a reassessment?")
	#due date, can be automated later to just be two weeks from where the date is now
	# due_date = input("\nPlease give me a due date for when the vendor should respond by (2 weeks ahead of )\n")

	#format string to input the variables into
	email_request_communication = f"Hi {contact}, \n Information Security GRC review all vendors for {internal} as part of the Vendor Security Risk Management Program. Please complete this by {fifteen_days_later}\n1. Could you please provide us with an overview of the services, software, or consultancy that you will be providing to our company. \n2. Please Provide the following documentation (If avaialable): \n \n\tThird Part vendor assurance reports relevant to security, such as a SOC 2 report, ISO27001 certificate, or SIC/SIG lite. \n\toLatest Security Review or summary of Penetration Test results \n\toInformation Security Policy \n\toIncident Response Policy \nIf none of the above are available to you or you are unable to share, please complete a Vendor seucity Risk Questionnaire from out Third-Party Risk Management solution - CyberGRX. \nIf you would require some help for The CyberGRX platform please email customercare@cyberGRX.com"

	#print out the formulated string
	newmail.body += email_request_communication
	if "y" in assessorreassess:
		newmail.Subject += "Reassessment for {vendor}"
	else:
		newmail.Subject += "Assessment for {vendor}"

def request_fulfilled():
	
	#Contact name
	contact_name = input("Person that you respond to if the vendor has completed the email\n")

	#vendor name
	vendor_name = input("Please enter the vendors name")

	#Internal Entity name; Can be set to a constant or predefined if you want
	internal = input("Please enter the name of the internal entity you are responding for\n")

	#request nuumber
	request_number = input("Please enter the Request number\n")

	#Tier ranking
	tier_ranking = input("Please enter the tier ranking that the client has ended up with\n")

	#short desctiption of why the tier is assigned
	tier_explanation = input("Please enter a brief description as to why the tier was assigned and what criteria facilitated that\n")

	#reviewers name, can be hard coded
	reviewers = input("Please enter the name of the reviewer:\n")

	date_of_review = input("Please enter the date the review was completed:\n")

	#client wesite
	vendors_site = input("Please enter the vendors website if applicable:\n")

	#clients number
	vendors_number = input("Please enter the vendors number if applicable:\n")

	#vendors speciaility, like a trading system or cloud computing company etc
	vendors_speciality = input("Please provide the vendors speciality")

	#vendors email
	vendors_email = input("Please provide the email that the vendor may are using")

	#what the vendor provides, mabye easier to take this out due to formatting issues
	vendor_provisions = input("Please enter what the vendor provides in terms of services")

	data_in_scope = input("Please enter what data is in scope")

	files_provided = input("Please enter what files have been provided, in a link format")

	#should be OK with the workaround being used but will still have to be manually numbered as there is no dictionary functionality
	document_list = input("The list of documents that have been sent, seperate with double commas for a new line, regular commas between words will be parsed as regular commas")

	#get it into a list
	document_list = document_list.split(",,")


	#What to send provided they have sent over the required information
	email_complete_communication = f"\n\n\n\nREQ-{request_number} - {vendor_name} - {tier_ranking} - Vendor Assessment - {internal}\nHi {contact_name}, \nThe review is complete, and we have no further question at this point. The vendor has been approved for use by {internal}. \nThe vendor has been assigned a {tier_ranking}. The next review will be in 2 years. \nPlease let us know if you have any other questions or concerns on the matter. \nNotes on the review actions performed:\n\tProvenanace REQ details: \n\tREQ-{request_number} - {vendor_name} - {tier_ranking} - Vendor Assessment - {internal}\n\t {tier_ranking} - {tier_explanation} \n\n\tVendor details and services provided:\n\t{vendors_site}\n\t{vendors_number}\n\t{vendors_speciality} \n\t{vendors_email} \n\t{vendor_name} provides the following: \n\t{vendor_provisions} \n\tData in scope: \n\t{data_in_scope}\n\tFiles provided\n\t{files_provided}. \nDocuments reviewed by {reviewers} on {date_of_review}\n\t"	#Document list goes on at the end but needs manual numbering

	print(email_complete_communication)
	newmail.body += email_complete_communication
	#cannot be done with list comprehension unfortunately so I have to use conventional for loop bu here is the implementation
	for i in document_list:
		print(f"\t{i.strip()}")
		newmail.body += f"\t{i.strip()}"

	newmail.Subject += f"REQ-{request_number} - {vendor_name} - {tier_ranking} - Vendor Assessment - {internal}"

	print("\n\nDont forget to add on the document list at the end, check the formatting for vender provisions (Just before data in scope) & delete this sentence before sending")

def off_boarding():
	vendor_name = input("Please enter the name of the vendor")
	company = input("Please enter the name of the company that you are requesting the Information for.")
	your_name = input("Please enter your name")
	Asking_string = f"Dear {vendor_name}, \nPlease can you inform us of the following: \n\tWhat data you hold on {company}\n\tWhat data has been retained (PII, Trade, Client etc)\n\tWhether in scope data has been destroyed\n\tWhether it has been destroyed in a secure manner\n\tIf there is a need for data to be held, how long would this be expected to be hosted by you or us\n\tIf at {vendor_name}, how and where is it stored and who will have access to the data. \n\tHow long will it be retained for.\nKind Regards {your_name}"
	print(Asking_string)
	newmail.Subject += "Off boarding" #might be append that i need to use
	newmail.body += Asking_string

#Introducing the topic to the vendor
decision = input("\nPlease select an option: \n[1] Making a request to a vendor via email \n[2] Making a response to confirm NFA is required from a client\n[3] Off boarding questions\n")

if decision == "1":
	request()
elif decision == "2":
	request_fulfilled()
elif decision == "3":
	off_boarding()
else:
	print("Please re run the program and pick an appropriate option either 1, 2 or 3")


newmail.Display()