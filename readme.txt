This example builds on Srinis example found on www.vb-helper.com.
Srinis example loaded data to a treeview from a text file with tabs
denoting indentation.
My contribution is the save treeview to file function.
It only saves one level of childs, it was all I needed for my other project.
Srinis example loads more than one level of childs. I am shure some small
modifications to my example will enable it to save multiple levels of childs,
but thats your jobb ;)
If you make that modification, I will be pleased if you send it to me at
trond.sorensen@bi.no

Now this is a compleat (allmost) set of functions working both ways.
All credits should go to Srini.


Srini´s original work:

SRINIS WORLD
haisrini@email.com

	Purpose
Load a TreeView from a text file with tabs denoting indentation

	Method
Read the file and count the leading tabs. Give the new node the right
parent for its level of indentation.


Trond Sørensen´s contribution:

trond.sorensen@bi.no
 
	Purpose
Save a TreeView to a text file with tabs denoting indentation

	Method
Read the treeview and save it to file. Give each childname a leading tab. 

	Product
The example program shows Srini´s and my functions working together.

******************************************************************************
	Disclaimer
This example program is provided "as is" with no warranty of any kind. It is
intended for demonstration purposes only. In particular, it does no error
handling. You can use the example in any form.

                                -*-