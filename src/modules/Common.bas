Attribute VB_Name = "Common"
Option Explicit

Public Type CompanyType
    Ruc As String
    Name As String
    LocalCodeEmission As String
    NroCtaDetraction As String
    
    TradeName As String
    TaxResidence As String
    EmissionPointAddress As String
    ContactInformation As String
End Type

Public Type AppType
    Env As EnvironmentEnum
    Internet As Boolean
    AutoProdCode As Boolean
    DocDirName As String
    LogLevel As Integer
    Premium As Boolean
    PrintMode As PrintModeEnum
    TicketItemInTwoLines As Boolean
    A4CenterCompanyData As Boolean
End Type

Public Type EmailType
    Provider As EmailProviderEnum
    SendWhenEmit As Boolean
    Address As String
    Password As String
    Message As String
    SignatureEmployeeName As String
    SignatureDepartment As String
    SignaturePhoneNumber As String
    SignatureCompanyName As String
End Type

Public Type RateType
    Igv As Double
End Type

Public Type SfsType
    Port As String
    Path As String
    DATAPath As String
    ENVIOPath As String
    RPTAPath As String
    REPOPath As String
End Type

Public Enum EmailProviderEnum
    GmailProv
    OutlookProv
End Enum

Public Enum PrintModeEnum
    PrintA4
    PrintTicket
    PrintSfs
End Enum

Public Enum EnvironmentEnum
    EnvLocal
    EnvProduction
End Enum

Public Enum SituationEnum
    CdpPorGenerarXml
    CdpXmlGenerado
    CdpEnviadoAceptado
    CdpEnviadoAceptadoConObs
    CdpRechazado
    CdpConErrores
    CdpPorValidarXml
    CdpEnviadoPorProcesar
    CdpEnviadoProcesando
    CdpRechazado10
    CdpEnviadoAceptado11
    CdpEnviadoAceptadoConObs12
End Enum

Public Function Prop() As Properties
    Dim PropertiesInstance As New Properties
    Set Prop = PropertiesInstance
End Function

Public Function DB() As Database
    Dim DatabaseInstance As New Database
    DatabaseInstance.ConnectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & PathJoin(Prop.Sfs.Path, "bd", "BDFacturador.db") & ";"
    DatabaseInstance.DebugMode = False
    Set DB = DatabaseInstance
End Function

Public Function Test() As VBAUnit
    Dim UnitTest As New VBAUnit
    Set Test = UnitTest
End Function
