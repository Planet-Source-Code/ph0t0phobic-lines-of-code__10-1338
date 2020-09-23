<div align="center">

## Lines of code


</div>

### Description

Extremely simple program that recurses through directories and reads any .vb file (excludes AssemblyInfo.vb) and counts the number of lines (minus blank lines, commented lines, multiline commands) and returns a number. Fits my purpose... would have made it read a .sln file and read all code files, but didnt need that capability. Feel free to modify to fit your purposes. Only 91 lines!! Good sample of recursive programming.
 
### More Info
 
Directory

Total Lines


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ph0t0phobic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ph0t0phobic.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB\.NET
**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__10-29.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ph0t0phobic-lines-of-code__10-1338/archive/master.zip)

### API Declarations

nope


### Source Code

```
Option Explicit On
Option Strict On
Imports System.IO
Imports System.Threading
'there are a few lines you have to change because they were filtered as being profane
'add another 's' in 'Public Clas', the submission form thinks its profanity
Public Clas frmCount
	Inherits System.Windows.Forms.Form
	Dim _t1 As Thread
#Region " Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		'Add any initialization after the InitializeComponent() call
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	Friend WithEvents txtFolder As System.Windows.Forms.TextBox
	Friend WithEvents lblFolder As System.Windows.Forms.Label
	Friend WithEvents cmdRun As System.Windows.Forms.Button
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.txtFolder = New System.Windows.Forms.TextBox()
		Me.lblFolder = New System.Windows.Forms.Label()
		Me.cmdRun = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'txtFolder
		'
		Me.txtFolder.Location = New System.Drawing.Point(8, 24)
		Me.txtFolder.Name = "txtFolder"
		Me.txtFolder.Size = New System.Drawing.Size(208, 20)
		Me.txtFolder.TabIndex = 0
		Me.txtFolder.Text = ""
		'
		'lblFolder
		'
		Me.lblFolder.Location = New System.Drawing.Point(8, 8)
		Me.lblFolder.Name = "lblFolder"
		Me.lblFolder.Size = New System.Drawing.Size(176, 16)
		Me.lblFolder.TabIndex = 1
		Me.lblFolder.Text = "initial folder:"
		'
		'cmdRun
		'
		Me.cmdRun.Location = New System.Drawing.Point(80, 56)
		Me.cmdRun.Name = "cmdRun"
		Me.cmdRun.Size = New System.Drawing.Size(64, 24)
		Me.cmdRun.TabIndex = 2
		Me.cmdRun.Text = "run"
		'
		'frmCount
		'
		Me.AcceptButton = Me.cmdRun
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(224, 85)
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdRun, Me.lblFolder, Me.txtFolder})
		Me.Font = New System.Drawing.Font("Trebuchet MS", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Name = "frmCount"
		Me.Text = "LinesOfCode"
		Me.ResumeLayout(False)
	End Sub
#End Region
	Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click
		cmdRun.Enabled = False
		cmdRun.Text = "running..."
		_t1 = Nothing
		_t1 = New Thread(AddressOf RecurseFolder)
		_t1.Name = "directory scan"
		_t1.Start()
	End Sub
	Private Sub RecurseFolder()
		Dim result As Integer = ScanDirectory(New DirectoryInfo(txtFolder.Text))
		cmdRun.Text = "run"
		cmdRun.Enabled = True
		MessageBox.Show(String.Concat(result.ToString, " lines of code exist"), "result", MessageBoxButtons.OK, MessageBoxIcon.Information)
	End Sub
	Public Function ScanDirectory(ByVal Directory As DirectoryInfo) As Integer
		Dim RunningTotal As Integer = 0
		Dim SubDir As DirectoryInfo
		For Each SubDir In Directory.GetDirectories
			RunningTotal += ScanDirectory(SubDir)
		Next
		Dim File As FileInfo
		For Each File In Directory.GetFiles("*.vb")
			'add another 's' in 'AsemblyInfo.vb', the submission form thinks its profanity
			If File.Name = "AsemblyInfo.vb" Then GoTo skipFile
			'Dim Output As New StreamWriter(File.FullName & ".count")			 'added for debugging purposes
			Dim fileText As String = File.OpenText.ReadToEnd
			Dim Lines() As String = fileText.Split(ControlChars.Lf)
			'Dim fileTotal As Integer = 0			 'added for debugging purposes
			Dim thisLine As String
			For Each thisLine In Lines
				thisLine = thisLine.Trim()
				If Not (thisLine.StartsWith("'") Or thisLine.StartsWith("_ ") Or thisLine.EndsWith(" _") Or (thisLine = String.Empty)) Then
					RunningTotal += 1
					'fileTotal += 1					 'added for debugging purposes
					'Output.WriteLine(String.Concat(fileTotal, ControlChars.Tab, thisLine))					 'added for debugging purposes
					'Else				'added for debugging purposes
					'Output.WriteLine(String.Concat(ControlChars.Tab, thisLine))					 'added for debugging purposes
				End If
			Next
			'Output.Close()			 'added for debugging purposes
skipFile:
		Next
		Return RunningTotal
	End Function
'add another 's' in 'End Clas', the submission form thinks its profanity
End Clas
```

