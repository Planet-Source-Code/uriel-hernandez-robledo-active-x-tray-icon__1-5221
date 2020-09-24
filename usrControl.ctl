VERSION 5.00
Begin VB.UserControl SystemTray 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "usrControl.ctx":0000
   ScaleHeight     =   990
   ScaleWidth      =   1155
   ToolboxBitmap   =   "usrControl.ctx":0010
   Begin VB.Image Image1 
      Height          =   810
      Left            =   120
      Picture         =   "usrControl.ctx":0322
      Stretch         =   -1  'True
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "SystemTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Asignar la imagen,y la propiedad ToolTip
'Ejecutar el metodo AñadirIcono. Se puede actualizar la imagen y el
'ToolTip o bien quitar del area del system tray el icono.
'Uriel Hernandez Robledo
'Todos los  metodos regresan un valor de tipo boolean.

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Public Event DblClick(Button As Integer)
'Public Event MouseUp(Button As Integer)
'Public Event MouseDown(Button As Integer)

'Public Event InTrayIconLeftButtonUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Public Event InTrayIconLeftButtonDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event InTrayIconLeftButtonDblClick(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute InTrayIconLeftButtonDblClick.VB_Description = "Se ejecuta cuando se pulsa dos veces rápidamente el botón izquierdo del mouse."

'Public Event InTrayIconMiddleButtonUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event InTrayIconMiddleButtonDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event InTrayIconMiddleButtonDblClick(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute InTrayIconMiddleButtonDblClick.VB_Description = "Se ejecuta cuando se pulsa dos veces rápidamente el botón interno del mouse."

Public Event InTrayIconRightButtonUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute InTrayIconRightButtonUp.VB_Description = "Se ejecuta cuando se libera el botón derecho del mouse."
'Public Event InTrayIconRightButtonDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Public Event InTrayIconRightButtonDblClick(Button As Integer, Shift As Integer, x As Single, Y As Single)

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Dim TheData As NOTIFYICONDATA


Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
    
'Private Const WM_LBUTTONDOWN = &H201
'Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203

'Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
'Private Const WM_RBUTTONDBLCLK = &H206

Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
'Private Const WM_MBUTTONUP = &H208

Private VarImagen As Picture, ImgDefault As Picture
Private msToolTip As String

Private Resultado As Boolean
'Default Property Values:
'Const m_def_ToolTipText = ""
'Property Variables:
'Dim m_ToolTipText As String


Public Function RestaurarIcono() As Boolean
    
    TheData.hIcon = VarImagen
    TheData.szTip = msToolTip
    RestaurarIcono = Shell_NotifyIcon(NIM_MODIFY, TheData)

End Function
Private Sub UserControl_Initialize()

    If ImgDefault Is Nothing Then Set ImgDefault = UserControl.Picture
End Sub
Private Sub UserControl_InitProperties()

    Set VarImagen = UserControl.Picture
'   m_ToolTipText = m_def_ToolTipText
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Static bBusy As Boolean
    Dim lAction As Long
  
   If bBusy = False Then
        bBusy = True

   If UserControl.ScaleMode = vbPixels Then
      lAction = x
    Else
        lAction = x / Screen.TwipsPerPixelX
    End If
        
        Select Case lAction
            ' *********************************************
            ' Eventos para el boton izquierdo del raton
            ' *********************************************
            'Case WM_LBUTTONDOWN
            '    RaiseEvent InTrayIconLeftButtonDown(Button, Shift, x, Y)
            'Case WM_LBUTTONUP
            '    RaiseEvent InTrayIconLeftButtonUp(Button, Shift, x, Y)
            Case WM_LBUTTONDBLCLK
                RaiseEvent InTrayIconLeftButtonDblClick(Button, Shift, x, Y)
               
            ' *********************************************
            ' Eventos para el boton interior del raton
            ' *********************************************
            Case WM_MBUTTONDBLCLK
                RaiseEvent InTrayIconMiddleButtonDblClick(Button, Shift, x, Y)
            Case WM_MBUTTONDOWN
                RaiseEvent InTrayIconMiddleButtonDown(Button, Shift, x, Y)
            'Case WM_MBUTTONUP
            '    RaiseEvent InTrayIconMiddleButtonUp(Button, Shift, x, Y)
            
            ' *********************************************
            ' Eventos para el boton derecho del raton
            ' *********************************************
            'Case WM_RBUTTONDOWN
            '   RaiseEvent InTrayIconRightButtonDown(Button, Shift, x, Y)
            Case WM_RBUTTONUP
                RaiseEvent InTrayIconRightButtonUp(Button, Shift, x, Y)
            'Case WM_RBUTTONDBLCLK
             '   RaiseEvent InTrayIconRightButtonDblClick(Button, Shift, x, Y)
         End Select

End If

bBusy = False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set VarImagen = PropBag.ReadProperty("Imagen", ImgDefault)
    msToolTip = PropBag.ReadProperty("ToolTip", vbNullString) & vbNullChar
    Set UserControl.Picture = VarImagen
'   m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    If VarImagen.Handle = ImgDefault.Handle Then
         PropBag.WriteProperty "Imagen", Nothing, Nothing
    Else
         PropBag.WriteProperty "Imagen", VarImagen, Nothing
    End If
    
    If msToolTip <> vbNullString Then
         PropBag.WriteProperty "ToolTip", msToolTip
    End If

'   Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
End Sub
Private Sub UserControl_Resize()
    UserControl.Height = UserControl.ScaleY(37, vbPixels, vbTwips)
    UserControl.Width = UserControl.ScaleX(47, vbPixels, vbTwips)
    Set UserControl.Picture = VarImagen
    Image1.Move 0, 0, ScaleWidth, ScaleHeight
    'Size 700, 600
End Sub
Private Sub UserControl_Terminate()
    Shell_NotifyIcon NIM_DELETE, TheData
End Sub
Public Function EliminarIcono() As Boolean
Attribute EliminarIcono.VB_Description = "Destruye el icono del System Tray."
    If Resultado Then
        EliminarIcono = Shell_NotifyIcon(NIM_DELETE, TheData)
        Resultado = False
    End If
End Function
Public Function ColocarIcono() As Boolean
Attribute ColocarIcono.VB_Description = "Añade un icono al System Tray."

    TheData.hIcon = VarImagen
    TheData.hWnd = UserControl.hWnd
    
    If msToolTip = vbNullString Then
        TheData.uFlags = NIF_ICON Or NIF_MESSAGE
    Else
        TheData.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        TheData.szTip = msToolTip
    End If

    TheData.uCallBackMessage = WM_MOUSEMOVE
    TheData.uId = VarImagen
    TheData.cbSize = Len(TheData)
    Resultado = Shell_NotifyIcon(NIM_ADD, TheData)
    ColocarIcono = Resultado
End Function
Public Function Actualizar(Optional Imagen As Picture, Optional New_ToolTip As String) As Boolean
    
    If Resultado Then
        If Not Imagen Is Nothing Then
            TheData.hIcon = Imagen.Handle
        End If
        
        If New_ToolTip <> vbNullString Then
            TheData.szTip = New_ToolTip & vbNullChar
        End If

        Actualizar = Shell_NotifyIcon(NIM_MODIFY, TheData)
    End If

End Function
Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Configura el texto que debe aparecer cuando se mueve el puntero del mouse sobre el icono en el System Tray."
    ToolTip = msToolTip
End Property
Public Property Let ToolTip(Text As String)
    msToolTip = Text & vbNullChar
    PropertyChanged ("ToolTip")
End Property

Public Property Get Imagen() As Picture
Attribute Imagen.VB_MemberFlags = "400"
    Set Imagen = VarImagen
End Property

Public Property Let Imagen(ByVal Imagen As Picture)
   If Imagen Is Nothing Then
      Set VarImagen = ImgDefault
   Else
      Set VarImagen = Imagen
   End If

   Set UserControl.Picture = VarImagen
   PropertyChanged "Imagen"
End Property
Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
Attribute ShowAboutBox.VB_MemberFlags = "40"
   dlgAbout.Show vbModal
   Unload dlgAbout
   Set dlgAbout = Nothing
End Sub
'Public Property Get ToolTipText() As String
'   ToolTipText = m_ToolTipText
'End Property
'
'Public Property Let ToolTipText(ByVal New_ToolTipText As String)
'   m_ToolTipText = New_ToolTipText
'   PropertyChanged "ToolTipText"
'End Property
'
