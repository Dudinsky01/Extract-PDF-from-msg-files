import os
import win32com.client

print("\n\n    /$$$$$$            /$$   /$$  /$$$$$$    /$$")                                                                                                                                                       
print("   /$$$_  $$          | $$  | $$ /$$__  $$ /$$$$")                                                                                                                                                       
print("  | $$$$\\ $$ /$$   /$$| $$  | $$|__/  \\ $$|_  $$")                                                                                                                                                      
print("  | $$ $$ $$|  $$ /$$/| $$$$$$$$  /$$$$$$/  | $$")                                                                                                                                                       
print("  | $$\\ $$$$ \\  $$$$/ |_____  $$ /$$____/   | $$")                                                                                                                                                       
print("  | $$ \\ $$$  >$$  $$       | $$| $$        | $$")                                                                                                                                                       
print("  |  $$$$$$/ /$$/\\  $$      | $$| $$$$$$$$ /$$$$$$")                                                                                                                                                     
print("   \\______/ |__/  \\__/      |__/|________/|______/")


Dest_Folder = r'path\\to\\your\\destination\\folder\\'                     				#Destination Folder
Src_Folder = 'path\\to\\your\\destination\\folder\\'				  			#Source Folder

src_count = 0												#Set counter to 0
dst_count = 0												

print("\nWORKING..\n")

for file in os.listdir(Src_Folder):									#Loop into source folder to list every files
	if file.endswith(".msg"):									#Filter only on .msg files
		src_count += 1										#Feed the source file counter
		filename = os.path.splitext(file)[0]							#Remove the file extension for renaming the pdf file like the .msg file
		outlook = win32com.client.Dispatch("Outlook.application").GetNamespace("MAPI")		#Use Outlook to open .msg file
		filepath = Src_Folder + '\\' + file
		msg = outlook.OpenSharedItem(filepath)
		att = msg.Attachments
		for i in att:										#Loop to get all attachments into .msg file
			if i.filename.upper().endswith(".PDF"):						#Filter only on .pdf files
				dst_count += 1								#Feed destination counter
				i.SaveAsFile(os.path.join(Dest_Folder, filename + '.pdf'))		#Save file with original name of the .msg file and add .pdf extension
			else:
				pass

if src_count == dst_count:										#Final Check - in my case, 1 email = 1 pdf file.			
	print("\nALL GOOD :)")
else:
	print("\nUhh.. you should check destination and source directory, something went wrong :( ")
