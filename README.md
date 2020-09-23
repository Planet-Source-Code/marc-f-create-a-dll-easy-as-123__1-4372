<div align="center">

## Create a DLL easy as 123 \!


</div>

### Description

Step by step Insctructions on how to create an ActiveX DLL File
 
### More Info
 
You Need to know Basic Navigation in VB


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marc F\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marc-f.md)
**Level**          |Unknown
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marc-f-create-a-dll-easy-as-123__1-4372/archive/master.zip)





### Source Code

```
You can download this code in Windows Write Format. Its easier to read !
Contact me with any questions. Marc 3dtech@thelakes.net
<Begin Instructions>
Create an ActiveX DLL File
Follow these steps.
1. Open VB and select to create an AxtiveX DLL project
  (an empty Class Module will appear)
2. Click on the "Project" menu. Select "Project1 Properties".
3. In the Properties Window set the project name to : CntrlPnl
4. Close the window and rename the Class Module to : ControlPanel
5. Now lets enter some code into the Class Module. Enter the following...
	Option Explicit
	Public Sub HardWare()
	Dim B As Long
	B = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1")
	End Sub
The above code will create a function to access the Control Panels
"Add Hardware" dialog.
Ok... Now its time to save amd compile your DLL file.
6. From the "File" menu select to "Save Project As".
  Save the project and class module etc.
7. From the "File" menu select "Make CntrlPnl.DLL"
8. Set the destination for the output DLL file. Also set any options at this time.
  For now the default options will be ok.
Using Your New DLL File
Ok, now lets put this DLL to use !
1. Click "File" menu and select "New Project". Save any changes to your DLL
  project if prompted to.
2. Select "Standard EXE" project. VB Now create a new blank project and loads
  one default form named "Form1".
3. From the "Project" menu select "References". A new window will open and
  display all available object libraries.
4. Click the "Browse Button" and navigate to the location where you compiled
   your DLL file.
5. Click the file and click "Open".
6. Your DLL will now be added to the list of "References". It should also be checked.
7. Close the "References" window.
8. Draw a Command Button on the form.
9. Double click the form to access the "Code View".
10. Click the ComboBox on the left and from it, select (General)
11. Your cursor should now appear above the "Form Load" event.
12. Declare your DLL file with this code: Private CP As New ControlPanel
It should look like this...
	Private CP As New ControlPanel
	___________________________
	Private Sub Form_Load()
	End Sub
So lets review. You added a Reference to the DLL file and declared it in your project.
Notice in the line "Private CP As New ControlPanel" that ControlPanel is the name of
your Class Module. You want to call the Class Module name and NOT the project name.
Using the Function of the DLL file
Now lets use the function from the DLL
1. Double click on the Command button to open the code view.
2. Now enter the following code : CP.HardWare
The code should appear like this...
	Private Sub Command1_Click()
	CP.HardWare
	End Sub
Notice "CP". You used it in the General Declarations.
Here is the complete code for the form :
	Private CP As New ControlPanel
	___________________________
	Private Sub Command1_Click()
	P.HardWare
	End Sub
Advanced Use:
Here is the complete code for the Class module.
' Begin Module ---
Option Explicit
Public Sub Access()
Dim A As Long
A = Shell("rundll32.exe shell32.dll,Control_RunDLL access.cpl,,5")
End Sub
Public Sub HardWare()
Dim B As Long
B = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1")
End Sub
Public Sub AddPrinter()
Dim C As Long
C = Shell("rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter")
End Sub
Public Sub Uninstall()
Dim D As Long
D = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1")
End Sub
Public Sub WindowsSetUp()
Dim E As Long
E = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,2")
End Sub
Public Sub ShortCut()
Dim F As Long
F = Shell("rundll32.exe apwiz.cpl,NewLinkHere %1")
End Sub
Public Sub DateTime()
Dim G As Long
G = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,0")
End Sub
Public Sub DUN()
Dim H As Long
H = Shell("rundll32.exe rnaui.dll,RnaWizard")
End Sub
Public Sub Display()
Dim I As Long
I = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0")
End Sub
Public Sub Font()
Dim J As Long
J = Shell("rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL FontsFolder ")
End Sub
Public Sub FormatFloppy()
Dim K As Long
K = Shell("rundll32.exe shell32.dll,SHFormatDrive")
End Sub
Public Sub Modem()
Dim L As Long
L = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl,,add")
End Sub
Public Sub Sound()
Dim M As Long
M = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0")
End Sub
Public Sub NetWork()
Dim N As Long
N = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl")
End Sub
Public Sub System()
Dim O As Long
O = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0")
End Sub
Public Sub Restart()
Dim P As Long
P = Shell("rundll32.exe user.exe,restartwindows")
End Sub
Public Sub ShutDown()
Dim Q As Long
Q = Shell("rundll32.exe user.exe,exitwindows")
End Sub
Public Sub Control()
Dim rc As Long
rc = Shell("Control.exe", vbNormalFocus)
End Sub
' End Module ---
Make the same calls as above in the example
CP.Access
- or -
CP.HardWare
- or -
CP.AddPrinter
- or -
Etc... Etc...
Marc F.
3dtech@thelakes.net
```

