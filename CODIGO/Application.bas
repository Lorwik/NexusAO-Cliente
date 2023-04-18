Attribute VB_Name = "Application"
'**************************************************************
' Application.bas - General API methods regarding the Application in general.
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

''
' Retrieves the active window's hWnd for this app.
'
' @return Retrieves the active window's hWnd for this app. If this app is not in the foreground it returns 0.

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private sNotepadTaskId As String

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    IsAppActive = (GetActiveWindow <> 0)

End Function

Public Sub LogError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************
    Dim File As Integer
    File = FreeFile

    'Hacemos un Left para poder solo obtener la letra del HD
    'Por que por culpa del UAC no guarda los logs en la carpeta del juego...
    Dim ErroresPath As String
    ErroresPath = Left$(App.path, 2) & "\Nexus AO\Errores\"

    If Dir(ErroresPath, vbDirectory) = "" Then
        MkDir ErroresPath
    End If

    'Matamos Notepad para evitar abrir decenas de block de notas.
    Shell ("taskkill /PID " & sNotepadTaskId)
        
    Open ErroresPath & "\Errores.log" For Append As #File
    
        Print #File, "Error: " & Numero
        Print #File, "Descripcion: " & Descripcion
        
        If LenB(Linea) <> 0 Then
            Print #File, "Linea: " & Linea
        End If
        
        Print #File, "Componente: " & Componente
        Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        Print #File, vbNullString
        
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine

    sNotepadTaskId = Shell("Notepad " & ErroresPath & "\Errores.log")

    Call AddtoRichTextBox(frmMain.RecTxt, "Errores fueron encontrados y se pusieron en el archivo Errores.log. Por favor envia el contenido de este archivo a los desarrolladores con el boton reporte de errores asi ayudas a mejorar el juego.", _
                            252, 257, 220, False, False, True)

End Sub

