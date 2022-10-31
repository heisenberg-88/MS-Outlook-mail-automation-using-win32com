import win32com.client
import win32com.client as win32

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNameSpace('MAPI')

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')


while True:
	print("waiting for new mails..")
	root_folder = mapi.Folders('email ID').Folders('Inbox').Folders('Inbox2')  # .Folder for reaching root/sub folders
	
	messages = root_folder.Items
	messages = messages.Restrict("[Unread]=True")
	
	for message in messages :
	
		msg = message.body
		#print(type(msg))
		inx = msg.find('Report ID')
		rep_id_flag = False
		if inx != -1:
		    inx += 12
		    if msg[inx] == '1' and msg[inx + 1] == '2' and ................ :
			rep_id_flag = True


		inx1 = msg.find('Requested By')
		req_by_flag = False
		if inx1 != -1:
		    inx1 += 15
		    if msg[inx1] == 'S' and msg[inx1 + 1] == 'c' and ............... :
			req_by_flag = True


		#print(rep_id_flag)
		#print(req_by_flag)
		#print("________\n")

		if rep_id_flag == True and req_by_flag == True:
		    flag = False
		    index1 =msg.find('Status')
		    index1 += 8
		    if msg[index1] == 'S':
			flag = True


		    if flag == True :
			final_caseid = ""
			final cobEND = ""

			list_of_words = msg.split()
			next_word = list_of_words[list_of_words.index('Case') + 3]
			final_caseid += next_word

			index2 = msgfind('SampleData : ')
			index2 += 15

			while( msg[index2] != 'R'):
			    final_cobEND += msg[index2]
			    index += 1

			#removing whitespaces
			final_caseid.replace(" ", "")
			final_cobEND.replace(" ", "")

			#print(final_caseid)
			#print(final_cobEND)

			mailItem = olApp.CreateItem(0)
			mailItem.subject = 'hey there'
			mailItem.BodyFormat = 1
			mailItem.Body = "hey parth"
			mailItem.To = 'email@email.com'

			mailItem.Display()
			mailItem.Save()
			mailItem.Send()
			print("sent" + "ID : " +final_caseid+"\n")

		message.unread = False
                
                
                
       
       
       
	
