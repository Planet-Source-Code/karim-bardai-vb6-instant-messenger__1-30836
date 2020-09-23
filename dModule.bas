Attribute VB_Name = "dModule"
Public Function FileExists(sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Public Function PlaySound(sound As String)
    Client.MMControl1.FileName = sound
    ' Open the MCI device.
    Client.MMControl1.Wait = True
    Client.MMControl1.Command = "Open"
    ' Play the sound without waiting.
    Client.MMControl1.Notify = True
    Client.MMControl1.Wait = False
    Client.MMControl1.Command = "Play"
End Function
