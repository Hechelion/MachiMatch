Attribute VB_Name = "ModVariables"
Option Explicit

Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Public Enum enumError
    sin_error = 0
    sin_SNAP = 1
    SNAP_Repetida = 2
    SNAP_Repetida_Revisada = 3
    Error_copia = 4
End Enum

Type archivo
    full_name As String
    lcase_name As String
    name_SinExtension As String
    Extension As String
    new_name(9) As Long
    similitud(9) As Double
    used As Integer
    index_newName As Long
    NumReferencias As Long
    have100 As Boolean 'Indica si la referencia encontrada tiene 100% de similitud
    
    movida As Integer
    conError As Integer
    NumError As enumError
End Type

Public lstROM() As archivo 'Lista de rom
Public lstSNAP() As archivo 'Lista de SNAP
Public lstIndex() As Long 'Lista de indices de ROM mostradas en el cuadro de edición
Public lstCharSearch() As String
Public lstCharReplace() As String

Dim Foco As Integer
Dim FactorScroll As Integer
Public sinROM As Boolean
Public movRoll As Boolean 'Auxiliar para indicar inhibir la carga cuando se mueve el scroll
Public iIdioma As Integer 'almacena el idioma en que está el programa
Public DatosFiltrados As Long

Public Function Get_locale()
    Get_locale = GetThreadLocale()
End Function

Public Function Get_SnapName(nIndex As Long) As String
If lstROM(nIndex).new_name(lstROM(nIndex).index_newName) > 0 Then
    Get_SnapName = lstSNAP(lstROM(nIndex).new_name(lstROM(nIndex).index_newName) - 1).full_name
Else
    Get_SnapName = ""
End If
End Function

Public Function Get_Por(nIndex As Long) As String
If lstROM(nIndex).new_name(lstROM(nIndex).index_newName) > 0 Then
    Get_Por = lstROM(nIndex).similitud(lstROM(nIndex).index_newName)
Else
    Get_Por = 0#
End If
End Function

Public Sub QSROM(nLstIndex() As Long, ByVal First As Long, ByVal Last As Long)
    Dim Low As Long, High As Long
    Dim Aux As Long
    Dim MidValue As String
    
    Low = First
    High = Last
    MidValue = lstROM(nLstIndex((First + Last) / 2)).name_SinExtension
    
    Do
        While lstROM(nLstIndex(Low)).name_SinExtension < MidValue
            Low = Low + 1
        Wend
        
        While lstROM(nLstIndex(High)).name_SinExtension > MidValue
            High = High - 1
        Wend
        
        If Low <= High Then
            'Swap C(Low), C(High)
            Aux = nLstIndex(Low)
            nLstIndex(Low) = nLstIndex(High)
            nLstIndex(High) = Aux
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then QSROM nLstIndex, First, High
    If Low < Last Then QSROM nLstIndex, Low, Last
    
End Sub

Public Sub QSSNAP(nLstIndex() As Long, ByVal First As Long, ByVal Last As Long)
    Dim Low As Long, High As Long
    Dim Aux As Long
    Dim MidValue As String
    
    Low = First
    High = Last
    MidValue = Get_SnapName(nLstIndex((First + Last) / 2))
    
    Do
        While Get_SnapName(nLstIndex(Low)) < MidValue
            Low = Low + 1
        Wend
        
        While Get_SnapName(nLstIndex(High)) > MidValue
            High = High - 1
        Wend
        
        If Low <= High Then
            'Swap C(Low), C(High)
            Aux = nLstIndex(Low)
            nLstIndex(Low) = nLstIndex(High)
            nLstIndex(High) = Aux
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then QSSNAP nLstIndex, First, High
    If Low < Last Then QSSNAP nLstIndex, Low, Last
    
End Sub

Public Sub QSPor(nLstIndex() As Long, ByVal First As Long, ByVal Last As Long)
    Dim Low As Long, High As Long
    Dim Aux As Long
    Dim MidValue As Double
    
    Low = First
    High = Last
    MidValue = Get_Por(nLstIndex((First + Last) / 2))
    
    Do
        While Get_Por(nLstIndex(Low)) < MidValue
            Low = Low + 1
        Wend
        
        While Get_Por(nLstIndex(High)) > MidValue
            High = High - 1
        Wend
        
        If Low <= High Then
            'Swap C(Low), C(High)
            Aux = nLstIndex(Low)
            nLstIndex(Low) = nLstIndex(High)
            nLstIndex(High) = Aux
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then QSPor nLstIndex, First, High
    If Low < Last Then QSPor nLstIndex, Low, Last
    
End Sub



