Attention Box by Francis Arnold Balatico

Good news to newbies!

You don't have to use complex API's to make those wonderfukl customized controls.
With a little imagination, you could make it with minimal API use.

Just like my custom message box named Attention Box.

*Supports 3 styles:
1. OkCancel - contains the ok and cancel button
2. YesNo - contains the yes and no buttons
3. OkOnly - ontains the Ok button only

*Supports button highlighting

*Supports keyboard ENTER key to accept and ESCAPE key to cancel

*Enhanced GUI

*Movable Form


Hope you like it.


Instructions:
1.Add the following files to your project
	1.1 frmMsgBox
	1.2 MsgBox.res	
	1.3 AttBox.bas
2.Place the following code into the LOAD event procedure of the initial form of your project
	call LoadAttImages
3.Code in the attention box much like using the messagebox.
	ex: Attbox "This is a Sample",1
4.When unloading your application, dont forget to code in the following line
	call UnLoadAttImages
5.That's it. Simple isn't it. :)


Hope you like this project. Please do vote and comment on it.

Thanks and God Bless.