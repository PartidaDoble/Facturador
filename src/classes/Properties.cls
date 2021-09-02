VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pApp As AppType
Private pEmail As EmailType
Private pRate As RateType
Private pCompany As CompanyType
Private pSfs As SfsType

Public Property Get Company() As CompanyType
    Dim Row As Long
    
    With sheetSetting
        ' Datos de la empresa
        Row = 1
        Row = Row + 1
        pCompany.Ruc = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pCompany.Name = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pCompany.LocalCodeEmission = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pCompany.NroCtaDetraction = Trim(.Cells(Row, 2))
        
        ' Datos adicionales de la empresa
        Row = 23
        Row = Row + 1
        pCompany.TradeName = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pCompany.TaxResidence = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pCompany.EmissionPointAddress = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pCompany.ContactInformation = Trim(.Cells(Row, 2))
    End With
    
    Company = pCompany
End Property

Public Property Get Sfs() As SfsType
    Dim Row As Long
    
    With sheetSetting
        ' Facturador SUNAT
        Row = 7
        Row = Row + 1
        pSfs.Path = Trim(.Cells(Row, 2))

        pSfs.Port = "9000"
    End With
    
    pSfs.DATAPath = PathJoin(pSfs.Path, "sunat_archivos\sfs\DATA")
    pSfs.ENVIOPath = PathJoin(pSfs.Path, "sunat_archivos\sfs\ENVIO")
    pSfs.RPTAPath = PathJoin(pSfs.Path, "sunat_archivos\sfs\RPTA")
    pSfs.REPOPath = PathJoin(pSfs.Path, "sunat_archivos\sfs\REPO")
    
    Sfs = pSfs
End Property

Public Property Get App() As AppType
    Dim Row As Long
    
    With sheetSetting
        ' Facturador excel
        Row = 10
        Row = Row + 1
        pApp.DocDirName = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pApp.AutoProdCode = .Cells(Row, 2) = "Automático"
        
        Row = Row + 1
        pApp.LogLevel = IIf(.Cells(Row, 2) = "Depuración", 2, 0)
        
        
        ' Facturador excel (Premium)
        Row = 18
        Row = Row + 1
        pApp.PrintMode = IIf(.Cells(Row, 2) = "A4", PrintA4, IIf(.Cells(Row, 2) = "Ticket", PrintTicket, PrintSfs))
        
        Row = Row + 1
        pApp.A4CenterCompanyData = .Cells(Row, 2) = "Sí"
        
        Row = Row + 1
        pApp.TicketItemInTwoLines = .Cells(Row, 2) = "Sí"
    End With
    
    pApp.Env = EnvLocal
    pApp.Premium = True
    App = pApp
End Property

Public Property Get Rate() As RateType
    pRate.Igv = sheetSetting.Cells(14, 2) / 100
    Rate = pRate
End Property

Public Property Get Email() As EmailType
    Dim Row As Long
    
    With sheetSetting
        ' Correo electrónico
        Row = 29
        Row = Row + 1
        pEmail.Provider = IIf(.Cells(Row, 2) = "Gmail", GmailProv, OutlookProv)
        
        Row = Row + 1
        pEmail.SendWhenEmit = .Cells(Row, 2) = "Sí"
        
        Row = Row + 1
        pEmail.Address = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pEmail.Password = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pEmail.Message = Trim(.Cells(Row, 2))
        
        ' Firma del correo electrónico
        Row = 36
        Row = Row + 1
        pEmail.SignatureEmployeeName = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pEmail.SignatureDepartment = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pEmail.SignaturePhoneNumber = Trim(.Cells(Row, 2))
        
        Row = Row + 1
        pEmail.SignatureCompanyName = Trim(.Cells(Row, 2))
    End With
    
    Email = pEmail
End Property
