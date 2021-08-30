VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSendEmail 
   Caption         =   "ENVIAR CORREO ELECTRÓNICO"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmSendEmail.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSendEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As New DocumentEntity
    Dim Attachments As New Collection
    Dim HtmlBody As String
    
    If Not ValidFields Then Exit Sub
    
    If lblAttachedPdf <> Empty Then
        Attachments.Add PathJoin(Prop.Sfs.REPOPath, lblAttachedPdf)
    End If
    
    If lblAttachedZip <> Empty Then
        Attachments.Add PathJoin(Prop.Sfs.ENVIOPath, lblAttachedZip)
    End If
    
    HtmlBody = Join(Split(txtBody, vbNewLine), "<br>")
    
    If Attachments.Count > 0 Then
        If SendEmail(Trim(txtReceiverEmail), Trim(txtSubject), HtmlBody, Attachments) Then
            MsgBox "El correo electrónico fue enviado correctamente.", vbInformation, "Correo enviado"
            
            Set Document = DocumentRepo.GetItem(txtDocumentId)
            
            If Not Document Is Nothing Then
                Document.EmailSent = True
                
                DocumentRepo.Update Document
            End If
        Else
            MsgBox "Hubo un error al enviar el correo electrónico.", vbCritical, "ERROR"
        End If
    Else
        MsgBox "No se encontró archivos adjunto. El correo no ha sido enviado.", vbCritical, "ERROR"
    End If
    
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    If Trim(txtReceiverEmail) = Empty Then
        MsgBox "Debe ingresar el correo electrónico del destinatario.", vbExclamation, "Subsane la observación"
        txtReceiverEmail.SetFocus
        Exit Function
    End If
    If InStr(Trim(txtReceiverEmail), "@") = 0 Or InStr(Trim(txtReceiverEmail), ".") = 0 Then
        MsgBox "El correo electrónico no es válido.", vbExclamation, "Subsane la observación"
        txtReceiverEmail.SetFocus
        Exit Function
    End If
    If Trim(txtSubject) = Empty Then
        MsgBox "Debe ingresar un asunto.", vbExclamation, "Subsane la observación"
        txtSubject.SetFocus
        Exit Function
    End If
    If Trim(lblAttachedPdf) = Empty And Trim(lblAttachedZip) = Empty Then
        MsgBox "Sin archivos adjuntos, no tiene sentido enviar el correo electrónico.", vbExclamation, "No hay archivos adjuntos"
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub UserForm_Initialize()
    If Prop.App.Env = EnvProduction Then
        txtDocumentId.Visible = False
    End If
End Sub
