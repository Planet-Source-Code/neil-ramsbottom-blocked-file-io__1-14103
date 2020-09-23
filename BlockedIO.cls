VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlockedIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Base 1   'VERY IMPORTANT

Public Event CopyProgress(lngPercentDone As Long)
Public Event CopyError(strDescription As String)
Public Event CopyComplete()
Public Event CopyCancelled()

Private bCancelAction As Boolean

Public Sub Cancel()

    bCancelAction = True
    
End Sub

Public Function CopyFile(strSource As String, strDestination As String, Optional bOverwrite As Boolean = False, Optional intBufferSize As Integer = 2048, Optional bSuppressErrors As Boolean = True) As Boolean
    
    'Author:        Neil Ramsbottom
    'Date:          31/12/2000
    'Purpose:       Blocked IO transfer from one file to another
    'License:       NON-COMMERCIAL (Contact me for commercial licence)
    'WWW:           http://www.nramsbottom.co.uk
    'Email:         nramsbottom@hotmail.com
    '
    'Known Bugs:    This code is a little unstable when UNC filenames are invloved. Sometimes it works, others it doesnt. I am looking at this.
    
    '
    '
    'DISCLAIMER:    I do not accept liability for loss of data of damage incurred through the use
    '               of this software. This code is not guaranteed in any way.
    '
    '               The above bit is so noone can claim that my code lost them money ~ 'Where There's Blame, There's a Claim!' ~ I think not ;-)
    '
    'Original Source (the code I was inspired by):
    '
    '/* textdump.cpp (30/9/1999)
    'blocked I/O demo*/
    '
    '#include <stdio.h>
    '#define FILENAME "arith.cpp"
    '#define BUFFER_LENGTH 1024
    '
    'int main()
    '{
    '   //Create a large buffer to store the data
    '   char block[BUFFER_LENGTH + 1];
    '   //a normal integer, and a special file handle
    '   int bytesread;
    '   FILE* handle=NULL;
    '
    '   //try to open the file and check for success
    '   handle=fopen(FILENAME,"r");
    '   if (!handle) return -1;
    '
    '   do{
    '       //read block of data from the file
    '       bytesread=fread(block,1,BUFFER_LENGTH,handle);
    '
    '       //add a terminating null so it becomes a  legal string
    '       block[bytesread]='\0';
    '
    '       //print the block
    '       printf(block);
    '
    '   //loop is we get as many bytes as we asked for
    '   } while (bytesread==BUFFER_LENGTH);
    '
    '   //close the file and quit
    '   fclose(handle);
    '   return 0;
    '}
    '
    'This code was taken from 'Teach Yourself C++ by Richard Riley'
    'ISBN: 9780340775288    Publisher: Hodder & Stoughton
    'Hope they dont ask for royaltys!
    
    Dim strError As String  'Holds an error string - duh!
    
    Dim BUFFER() As Byte    'holds <x> number of bytes until we want to write them
    
    Dim intSourceFile As Integer
    Dim intDestinationFile As Integer
    
    If strSource = "" Then
        strError = "No Source File Specified"
    ElseIf strDestination = "" Then
        strError = "No Destination File Specified"
    ElseIf Dir(strSource) = "" Then
        strError = "Source File Does Not Exist"
    End If
    
    If strError <> "" Then
        GoTo copy_error
    End If
    
    On Error Resume Next
    'On Error GoTo copy_error
    
    If Dir(strDestination) <> "" Then
        If bOverwrite Then  'kill the file and create a new one
            'Is file read only? if so, make writable and kill it
            Err.Clear
            If (GetAttr(strDestination) And vbReadOnly) = vbReadOnly Then
                SetAttr strDestination, vbNormal
            End If
            Kill strDestination
                If Err.Number <> 0 Then 'There was an error with the kill
                    
                    'If Not bSuppressErrors Then
                        'MsgBox Err.Description, vbExclamation, "Error - " & Err.Number
                        RaiseEvent CopyError(Err.Description)
                    'End If
                    
                    Exit Function
                End If
        Else
            
            RaiseEvent CopyError("Destination file exists")
        
            Exit Function
        End If
    End If
    
    'Get a file handle for the source file
    intSourceFile = FreeFile
    'Open the source file for reading, locked so nobody can write to it while we use it
    Open strSource For Binary Access Read Lock Write As #intSourceFile

    
    'Get a file handle for the destination file
    intDestinationFile = FreeFile
    'Open the destination file for writing locked so only this process can write to it
    Open strDestination For Binary Access Write Lock Write As #intDestinationFile
    
    'If the size of the file is less than the buffer size (i.e. 11 bytes)
    'resize the buffer to the size of the file and transfer it all at once
    If LOF(intSourceFile) < intBufferSize Then
        ReDim BUFFER(LOF(intSourceFile))
    Else
        'Otherwise resize the buffer to the specified size and continue
        ReDim BUFFER(intBufferSize)
    End If

    'The AND bit in this loop was a bit of a botch fix because EOF was never being reached (dunno why)
    While Not EOF(intSourceFile) And Loc(intSourceFile) <> LOF(intSourceFile)
    
        'if we are not at the end of the file then
        If Loc(intSourceFile) <> LOF(intSourceFile) Then
            'If the remainder of the file is less than buffer size then resize
            'the buffer to remainder size
            If LOF(intSourceFile) - Loc(intSourceFile) < intBufferSize Then
                ReDim BUFFER(LOF(intSourceFile) - Loc(intSourceFile))
            End If
        Else
            'Otherwise erase the buffer (because we are at the end of the file)
            'This prevents bounds errors when (redim BUFFER(0) fails because of the option base)
            Erase BUFFER
        End If
    
        'Get a block from the source file
        Get #intSourceFile, , BUFFER
        'Write the block to the destination file
        Put #intDestinationFile, , BUFFER
    
        'Let the CopyProgress event know how far in we are (percent value out of 100)
        RaiseEvent CopyProgress(Loc(intSourceFile) / LOF(intSourceFile) * 100)
        
        'Yield to other processes to allow user to click a cancel button, etc
        DoEvents
        
        If bCancelAction Then
            'Shutdown the copy operation
            Close intSourceFile
            Close intDestinationFile
            
            Kill strDestination     'Erase the file we were copying
            
            'Reset the cancel flag
            bCancelAction = False
            'Let parent object know we have cancelled
            RaiseEvent CopyCancelled
            Exit Function
            
        End If
        
    Wend
    
    Close intSourceFile
    Close intDestinationFile
    
    'Inform the parent that we have finished the operation
    RaiseEvent CopyComplete
    
    CopyFile = True

Exit Function

copy_error:
    
    RaiseEvent CopyError(strError)
    Exit Function
    
End Function
