VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsSurfaceManStatic.cls - Inherits from clsSurfaceManager. Is designed to load
' surfaces at startup, and never deallocating them.
' This grants high performance can use a lot of RAM.
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Nexus AO mod Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

'Inherit from the surface manager
Implements clsSurfaceManager

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 541

Private Type SURFACE_ENTRY_STATIC

    fileIndex As Long
    Surface As Direct3DTexture8

End Type

Private Type HashBucket

    sufraceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_STATIC

End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashBucket

Private DirectD3D                        As D3DX8

Private Const DIB_PAL_COLORS             As Long = 1

Private Const DIB_RGB_COLORS             As Long = 0

Private Declare Function SetDIBitsToDevice _
                Lib "GDI32.dll" (ByVal hdc As Long, _
                                 ByVal XDest As Long, _
                                 ByVal YDest As Long, _
                                 ByVal dwWidth As Long, _
                                 ByVal dwHeight As Long, _
                                 ByVal xSrc As Long, _
                                 ByVal ySrc As Long, _
                                 ByVal uStartScan As Long, _
                                 ByVal CScanLine As Long, _
                                 lpBits As Any, _
                                 lpbmi As BITMAPINFO, _
                                 ByVal fuColorUse As Long) As Long
    
Private Sub Class_Initialize()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    '
    '**************************************************************
End Sub

Private Sub Class_Terminate()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Clean up
    '**************************************************************
    Dim i As Long

    Dim J As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        For J = 1 To surfaceList(i).sufraceCount
            Set surfaceList(i).SurfaceEntry(J).Surface = Nothing
        Next J
        
        'Destroy the array
        Erase surfaceList(i).SurfaceEntry
    Next i

End Sub

Private Sub clsSurfaceManager_Initialize(ByRef DD As D3DX8, _
                                         ByVal graphicPath As String, _
                                         Optional ByVal maxMemoryUsageInMb As Long = -1&)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 29/07/2012 - ^[GS]^
    '**************************************************************
    Set DirectD3D = DD

    Call LoadSurfaces(graphicPath)

End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As Direct3DTexture8

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Retrieves the requested texture
    '**************************************************************
    On Error GoTo ErrHandler:

    Dim i As Long
    
    ' Search for it and return it
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)

        For i = 1 To .sufraceCount

            If .SurfaceEntry(i).fileIndex = fileIndex Then
                Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
                Exit Property

            End If

        Next i

    End With
    
    'We have a missing resource!
    MsgBox "No se encuentra el recurso " & CStr(fileIndex) & ". Reinstale el juego, " & "y si el problema persiste contactese con los adminsitradores", vbOKOnly Or vbCritical Or vbExclamation, "Error"
    
    End
    Exit Property

ErrHandler:
    MsgBox "Un error inesperado ocurrió a intentar cargar el recurso " & CStr(fileIndex) & ". & vbcrlf" & "Error: " & CStr(Err.number), vbOKOnly Or vbCritical Or vbExclamation, "Error"
    
    End

End Property

Private Sub LoadSurfaces(ByVal ResourcePath As String)

    '**************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 12/02/2007
    'Loads all surfaces in random order and then sorts them
    '**************************************************************
    On Error Resume Next

    Dim surfaceIndex As Long

    Dim bmpInfo      As BITMAPINFO

    Dim data()       As Byte

    Dim i            As Long
    
    While GetNext_Bitmap(ResourcePath, i, bmpInfo, data(), surfaceIndex)

        Call LoadSurface(bmpInfo, data(), surfaceIndex)
    Wend

End Sub

Private Sub LoadSurface(ByRef bmpInfo As BITMAPINFO, _
                        ByRef data() As Byte, _
                        ByVal fileIndex As Long)

    '**************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 11/30/2007
    'Loads the surface named fileIndex + ".bmp" and inserts it to the
    'surface list in the listIndex position
    '**************************************************************
    On Error GoTo ErrHandler

    Dim newSurface   As SURFACE_ENTRY_STATIC

    Dim surface_desc As D3DSURFACE_DESC

    Dim texture_info As D3DXIMAGE_INFO

    With newSurface
        .fileIndex = fileIndex
        
        Set newSurface.Surface = DirectD3D.CreateTextureFromFileInMemoryEx(DirectDevice, data(0), UBound(data) + 1, D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
        
        newSurface.Surface.GetLevelDesc 0, surface_desc
        
    End With
    
    'Store the surface in the given index (it MUST be empty or data will be lost)
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .sufraceCount = .sufraceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .sufraceCount) As SURFACE_ENTRY_STATIC
        
        .SurfaceEntry(.sufraceCount) = newSurface

    End With

    Exit Sub

ErrHandler:

End Sub

