Attribute VB_Name = "Functions"
' A universal error message just to avoid typing out long lengths of code all the time
Public Function Error()
    Call MsgBox("You did domething wrong didnt you????", 16, "Error")
End Function

Public Sub LoadAddressBook()
    ' Clear the list box
    Call frmSend.lstAddress.Clear
    ' Set a string for the location of the file
    Dim address As String
   ' Add the location of the file to the string
    address = "C:\AddressBook.txt"
    ' A varable for the loop to read the data from the file into
    Dim n As Variant
    
    Dim infile As Integer
    infile = FreeFile
    ' Open the file to be inputed into the list box
    Open address For Input As #infile
    ' A loop until the end of the file to read the data in
    Do While Not EOF(infile)
        Line Input #infile, n
        'add the data read in to lstAddress
        frmSend.lstAddress.AddItem n
    Loop
    ' close the file
    Close infile
End Sub
Public Sub LoadAddressBook2()
    ' Clear the list box
    Call frmAddress.lstAddress.Clear
    ' Set a string for the location of the file
    Dim address As String
   ' Add the location of the file to the string
    address = "C:\AddressBook.txt"
    ' A varable for the loop to read the data from the file into
    Dim n As Variant
    
    Dim infile As Integer
    infile = FreeFile
    ' Open the file to be inputed into the list box
    Open address For Input As #infile
    ' A loop until the end of the file to read the data in
    Do While Not EOF(infile)
        Line Input #infile, n
        'add the data read in to lstAddress
        frmAddress.lstAddress.AddItem n
    Loop
    ' close the file
    Close infile
End Sub
