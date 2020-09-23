To run the  InterCommVB Demonstration>>

Step 1> Open InterCommVB.vbp in VB and compile it

Step 2>Open demoClient.Frm by itself after selecting File >> Open (all files). Open it and manually add a reference to InterCommVB II. To do so, select Project >> References, Browse for InterCommVB.exe and click OK. Now compile it

Step 3>Perform the same steps with demoServer.frm

Now run the compiled versions of ClientDemo and ServerDemo....

THREADING ISSUES
This component has been tested and found to work perfectly with VB6 - However below are some possible pitfalls...

1>This component contains forms. Since the original version of VB 5 did not allow components containing user interface elements  such as forms and controls to mulithread, multithreading may not work with VB 5, though it must work as a single threaded component well enough

2>This component (along with multithreading support) must work perfectly with VB 5, if SP2 is installed (It has not been tested with VB 5(Sp2) though)

3>I have not got to use VB.NET as yet, so I cannot say whether it will work with VB.NET or not... One important problem I faced was that I had to use object pointers but the ObjPtr() function has been removed in VB.NET. Therefore, to increase the chances of compatibility, I have designed a new function, ObjectPtr() That does the same work....


F YOU FIND THIS CODE USEFUL,PLEASE VOTE !

Contacting Me-
You could contact me at srideepprasad@digitalme.com  if you have any questions/problems or suggestions...