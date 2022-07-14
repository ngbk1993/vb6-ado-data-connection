VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Main"
   ClientHeight    =   2064
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   2064
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1776
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4332
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example shows how to Connect to a Microsoft Access Database
' 1) Goto Project -> Reference -> Tick Microsoft DAO 3.6 Object Library
' Created on: 28-Nov-2016
' Example By: Ng, Boon Khai
' Email     : ngbk1993@yahoo.com

Dim db As DAO.Database
Dim RecordSet_1 As DAO.Recordset


Public Sub LoadDatabase() 'Function to load Database

    Set db = OpenDatabase(App.Path & "\testing.mdb")
    Set RecordSet_1 = db.OpenRecordset("tblUser")

    RecordSet_1.MoveFirst 'Move the record set to the first element

    Do While Not RecordSet_1.EOF
        List1.AddItem RecordSet_1.Fields(0).Value
        RecordSet_1.MoveNext
    Loop

    RecordSet_1.Close
    db.Close

End Sub

Public Function GetLastNumber() As Integer
    
    Set db = OpenDatabase(App.Path & "\testing.mdb")
    Set RecordSet_1 = db.OpenRecordset("tblCount")

    RecordSet_1.MoveLast
    GetLastNumber = RecordSet_1.Fields(1).Value
    
    MsgBox RecordSet_1.Fields(1).SourceField
    
    RecordSet_1.Close
    db.Close

End Function

Public Sub SetLastNumber(Number As Integer)

    Set db = OpenDatabase(App.Path & "\testing.mdb")
    Set RecordSet_1 = db.OpenRecordset("tblCount")

    With db
        With RecordSet_1
            .MoveLast
            .Edit
            !Count = Number
            .Update
            .Close
        End With
    End With

End Sub

Public Function GetHighestNumber() As Integer

    Dim Temp As Integer
    Set db = OpenDatabase(App.Path & "\testing.mdb")
    Set RecordSet_1 = db.OpenRecordset("tblUser")

    RecordSet_1.MoveFirst
    
    Do While Not RecordSet_1.EOF
        If Temp < RecordSet_1.Fields(1).Value Then Temp = RecordSet_1.Fields(1).Value
        RecordSet_1.MoveNext
    Loop
    
    RecordSet_1.Close
    db.Close
    GetHighestNumber = Temp
    
End Function

Private Sub Form_Load()

    LoadDatabase
    Me.Caption = GetHighestNumber
    SetLastNumber (10)
    
End Sub
