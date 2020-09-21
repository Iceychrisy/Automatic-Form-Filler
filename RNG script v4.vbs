'------------------------------------------------------------------------'
'Created by Chris Lipscombe on 20/09/2020'
'This script is a basic form fillout, it doesnt do verification'
'------------------------------------------------------------------------'

'The website'
websiteurl = "https://"

'Your details'
myfirstname = "James" 'Change if you want"
mylastname = "Thompson" 'Change if you want"
mycomment = "Great new Shop, Jean-Luc was a great help with picking out an outfit. He was a Delit to work with." 'Change if you want'
myemail = "@gmail.com" 'Don't change'

'How many times you want to fill out the form (change the first number not the +1)'
howmanyloops = 10 + 1 
dim x
x = 1

do while x < howmanyloops
 
	fullemail = myfirstname & "." & mylastname & x & myemail
	
	set webbrowser = createobject("internetexplorer.application")
	webbrowser.statusbar = false
	webbrowser.menubar = false
	webbrowser.toolbar = false
	webbrowser.visible = true

	webbrowser.navigate(websiteurl)

	wscript.sleep(2000)

	webbrowser.document.all.item("dwfrm_contactus_firstname").value = myfirstname
	webbrowser.document.all.item("dwfrm_contactus_lastname").value = mylastname
	webbrowser.document.all.item("dwfrm_contactus_email").value = fullemail

	wscript.sleep(500)

	webbrowser.document.all.item("dwfrm_contactus_comment").value = mycomment
	
	wscript.sleep(500)
	
	webbrowser.document.all.item("dwfrm_contactus_send").click
	
	wscript.sleep(500)
	
	webbrowser.quit

	wscript.sleep(500)
	
	x = x + 1
	
loop