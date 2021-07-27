Attribute VB_Name = "Common"
Option Explicit

Public Type AppType
    Env As String
    Debug As Boolean
    Internet As Boolean
    AutoProdCode As Boolean
End Type

Public Type RateType
    Igv As Double
End Type

Public Type CompanyType
    Ruc As String
    Name As String
    LocalCodeEmission As String
End Type

Public Type SfsType
    Port As String
    Path As String
    DATAPath As String
    ENVIOPath As String
    RPTAPath As String
    REPOPath As String
End Type

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
    DatabaseInstance.ConnectionString = "DRIVER=SQLite3 ODBC Driver;Database=D:\sfs\SFS_v1.3.4.4\bd\BDFacturador.db;"
    DatabaseInstance.DebugMode = False
    'DatabaseInstance.AutomaticCreationAndUpdateTimestamp = false
    Set DB = DatabaseInstance
End Function
