import win32com.client as win32

# Create the Outlook application object
outlook = win32.Dispatch('Outlook.Application')

# Get the MAPI namespace
namespace = outlook.GetNamespace("MAPI")

# Get the Inbox folder
inbox = namespace.GetDefaultFolder(6)  # 6 represents the index of the Inbox folder

# Get the email items in the Inbox
email_items = inbox.Items

# Define the search criteria
search_criteria = "@SQL=\"urn:schemas:httpmail:textdescription\" LIKE '%chromatin%'"

# Perform the search
filtered_emails = email_items.Restrict(search_criteria)

root_folder = namespace.GetDefaultFolder(6)  # 6 represents the index of the Inbox folder

# Create a new folder
new_folder_name = "Hafsha_Work "
new_folder = root_folder.Folders.Add(new_folder_name)

for email in filtered_emails:
    # Set the flag on the email
    email.FlagRequest = "Urgent"  # Replace with your desired flag request

    # Save the changes made to the email
    email.Save()

# Iterate over the filtered emails
for email in filtered_emails:
    #email.Display()
    #subject = email.Subject
    sender = email.SenderEmailAddress
   #received_time = email.ReceivedTime
    body = email.Body


    # Print the email details

    #print("Subject:", subject)
    print("Sender:", sender)
    #print("Received Time:", received_time)
    print("Body:", body)
    print("----------------------------------------------")


    for message in filtered_emails:
        message.Move(new_folder)
