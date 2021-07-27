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
Private pRate As RateType
Private pCompany As CompanyType
Private pSfs As SfsType

Public Property Get App() As AppType
    pApp.Env = "local"
    pApp.Debug = True
    pApp.Internet = True
    pApp.AutoProdCode = True
    App = pApp
End Property

Public Property Get Rate() As RateType
    pRate.Igv = 0.18
    Rate = pRate
End Property

Public Property Get Company() As CompanyType
    pCompany.Ruc = "20123456789"
    pCompany.Name = "EMPRESA SAC"
    pCompany.LocalCodeEmission = "0000"
    Company = pCompany
End Property

Public Property Get Sfs() As SfsType
    Dim fs As New FileSystemObject
    pSfs.Port = "9000"
    pSfs.Path = "D:\sfs\SFS_v1.3.4.4"
    pSfs.DATAPath = fs.BuildPath(pSfs.Path, "sunat_archivos\sfs\DATA")
    pSfs.ENVIOPath = fs.BuildPath(pSfs.Path, "sunat_archivos\sfs\ENVIO")
    pSfs.RPTAPath = fs.BuildPath(pSfs.Path, "sunat_archivos\sfs\RPTA")
    pSfs.REPOPath = fs.BuildPath(pSfs.Path, "sunat_archivos\sfs\REPO")
    Sfs = pSfs
End Property
