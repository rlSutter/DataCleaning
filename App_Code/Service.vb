Option Explicit On
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Collections.Generic
Imports System.Collections
Imports System.Configuration
Imports System.Net.Mail
Imports System.Math
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
'Imports DataFlux.dfClient
Imports System
Imports OAuth
Imports Geocode
Imports log4net
Imports System.Web.Script.Serialization
Imports MelissaData
Imports System.Runtime.InteropServices
Imports System.Threading.Thread

<WebService(Namespace:="http://datafluxapp1.hq.local/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
    End Enum

    ' Classes Used for Geocodeio to deserialize results of geocoding
    Public Class GeoAddressComponentsInput
        Public Property Number As String
        Public Property Street As String
        Public Property Suffix As String
        Public Property FormattedStreet As String
        Public Property City As String
        Public Property State As String
        Public Property Zip As String
    End Class

    Public Class GeoInput
        Public Property AddressComponents As GeoAddressComponentsInput
        Public Property FormattedAddress As String
    End Class

    Public Class AddressComponents2
        Public Property number As String
        Public Property street As String
        Public Property suffix As String
        Public Property formatted_street As String
        Public Property city As String
        Public Property county As String
        Public Property state As String
        Public Property zip As String
    End Class

    Public Class GeoLocation
        Public Property Lat As Double
        Public Property Lng As Double
    End Class

    Public Class GeoResult
        Public Property AddressComponents As AddressComponents2
        Public Property FormattedAddress As String
        Public Property Location As GeoLocation
        Public Property Accuracy As Double
        Public Property AccuracyType As String
        Public Property Source As String
    End Class

    Public Class GeocodioObj
        Public Property input As GeoInput
        Public Property results As GeoResult()
    End Class
    Public Class Dupl_Addr_Check
        Public Shared g_dt As Data.DataTable = New Data.DataTable("Addr_Input")
        'Public Shared g_dt As Data.DataTable
        '{

        '}
    End Class


    <WebMethod(Description:="Standardizes the supplied contact record and create it if needed in the database")> _
    Public Function StandardizeContact(ByVal sXML As String) As XmlDocument
        ' This function takes a contact supplied, either within the supplied XML or in a 
        ' database record, and standardizes it and generates a match code.  Based on 
        ' parameters, it either returns the data generated or updates the database 

        ' The input parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <Contacts>
        '       <Contact>
        '           <Debug>             - A flag to indicate the service is to run in Debug mode or not
        '                                   "Y"  - Yes for debug mode on.. logging on
        '                                   "N"  - No for debug mode off.. logging off
        '                                   "T"  - Test mode on.. logging off
        '           <Database>          - "C" create S_CONTACT record(s), "U" update record, blank do nothing
        '           <ConId>             - The Id of an existing contact, if applicable
        '           <FirstName>         - First name of contact
        '           <MidName>           - Middle name of contact
        '           <LastName>          - Last name of contact
        '           <Gender>            - Gender of contact
        '           <FullName>          - Full name of contact
        '       </Contact>
        '   </Contacts>

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim mypath, debug, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SCDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim FST_NAME, MID_NAME, LAST_NAME, GENDER, MATCH_CODE, FULL_NAME, FULL_NAME_MD_PARSE, CON_ID As String
        Dim tFST_NAME, tMID_NAME, tLAST_NAME, tFULL_NAME As String
        Dim sFST_NAME, sMID_NAME, sLAST_NAME As String
        Dim temp, database As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        FST_NAME = ""
        MID_NAME = "L"
        LAST_NAME = ""
        GENDER = ""
        MATCH_CODE = ""
        FULL_NAME = ""
        FULL_NAME_MD_PARSE = ""
        tFST_NAME = ""
        tMID_NAME = ""
        tLAST_NAME = ""
        tFULL_NAME = ""
        sFST_NAME = ""
        sMID_NAME = ""
        sLAST_NAME = ""
        CON_ID = ""
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0
        debug = "N"

        ' ============================================
        ' Check parameters
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        HttpUtility.UrlDecode(sXML)
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//Contacts/Contact")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
                debug = Trim(UCase(debug))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("StandardizeContact_debug").ToUpper()
            If temp = "Y" And debug <> "T" Then debug = temp
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\standardize_contact_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\StandardizeContact.log"
            Try
                log4net.GlobalContext.Properties("SCLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' Commented out; Dataflux license expired; 8/1/2017;
        '' ============================================
        '' Connect to Dataflux
        'Dim dataflux = System.Configuration.ConfigurationManager.AppSettings("dataflux")
        'If dataflux = "" Then dataflux = "datafluxapp1.hq.local"
        'Dim config As New Hashtable
        'config.Add("server", dataflux)
        'config.Add("transport", "TCP")
        'If debug = "Y" Then config.Add("log_file", "C:\Logs\enter.log")

        'Dim dfsession As New DataFlux.dfClient.SessionObject(config)
        'If (dfsession Is Nothing) Then
        '    If debug = "Y" Then mydebuglog.Debug("Unable to open Dataflux")
        '    GoTo CloseOut2
        'Else
        '    If debug = "Y" Then mydebuglog.Debug("  Opening dataflux on " & dataflux)
        'End If

        ' ===========================================
        'MelissaData components (Name Object)
        'MelissaData Initialization
        Dim nameObj As New mdName
        Dim nameObjParseResult As String = ""
        Dim dPath As String = System.Configuration.ConfigurationManager.AppSettings("MD_DataPath")
        Dim dLICENSE As String = System.Configuration.ConfigurationManager.AppSettings("MD_Key")
        Try
            'Set License
            nameObj.SetLicenseString(dLICENSE)
            If Convert.ToDateTime(nameObj.GetLicenseExpirationDate) < Now Then
                If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File")
                errmsg = errmsg & "MelissaData Data License Expired: " & nameObj.GetLicenseExpirationDate
                GoTo CloseOut2
            End If
            'Error Checking
            nameObj.SetPathToNameFiles(dPath)
            If (nameObj.InitializeDataFiles() <> mdName.ProgramStatus.NoError) Then
                If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File: " + nameObj.GetInitializeErrorString())
                errmsg = errmsg & "Unable to Initiate MelissaData Data File: " + nameObj.GetInitializeErrorString()
                GoTo CloseOut2
            Else
                If debug = "Y" Then mydebuglog.Debug("MelissaData Data File Initialized")
                'Set Name Object Options
                nameObj.SetPrimaryNameHint(mdName.NameHints.DefinitelyFull)
                'nameObj.SetPrimaryNameHint(mdName.NameHints.DefinitelyInverse)

            End If
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File: " + ex.ToString())
            errmsg = errmsg & "Unable to Initiate MelissaData Data File: " + ex.ToString()
        End Try

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            Try
                'debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        For i = 0 To oNodeList.Count - 1
            errmsg = ""
            If debug <> "T" Then
                FST_NAME = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("FirstName", oNodeList.Item(i)))))
                MID_NAME = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("MidName", oNodeList.Item(i)))))
                LAST_NAME = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("LastName", oNodeList.Item(i)))))
                GENDER = Trim(GetNodeValue("Gender", oNodeList.Item(i)))
                FULL_NAME = Trim(HttpUtility.UrlDecode(GetNodeValue("FullName", oNodeList.Item(i))))
                'If FULL_NAME = "" Then  ' Modified to Teat FULL NAME tag as lower priority; 5/29/2016; Ren Hou;
                If LAST_NAME <> "" And FST_NAME <> "" Then  'Parse Last Name
                    FULL_NAME = FST_NAME & " " & IIf(MID_NAME = "", "", MID_NAME & " ") & LAST_NAME
                    'If last name is passed in, user inverse order with last name in front.
                    'FULL_NAME_MD_PARSE = LAST_NAME & ", " & FST_NAME & IIf(MID_NAME = "", "", " " & MID_NAME)
                    FULL_NAME_MD_PARSE = LAST_NAME
                    nameObj.SetPrimaryNameHint(mdName.NameHints.MixedLastName)
                    If debug = "Y" Then
                        mydebuglog.Debug("  Parsing Last Name... ")
                        mydebuglog.Debug("  FirstName: " & FST_NAME)
                        mydebuglog.Debug("  MidName: " & MID_NAME)
                        mydebuglog.Debug("  LastName: " & LAST_NAME)
                        mydebuglog.Debug("  Gender: " & GENDER)
                        mydebuglog.Debug("  FullName: " & FULL_NAME)
                        mydebuglog.Debug("  NAME_MD_PARSE: " & FULL_NAME_MD_PARSE)
                        mydebuglog.Debug("  ConId: " & CON_ID & vbCrLf & "------")
                    End If
                    Try
                        nameObj.SetFullName(FULL_NAME_MD_PARSE)
                        nameObj.Parse()
                        temp = If(Trim(nameObj.GetLastName()) <> "", nameObj.GetLastName(), "")
                        ' Check Parse results
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "MD Last Name Parse Result: " & nameObj.GetResults())
                        If nameObj.GetResults().Contains("NS02") Then
                            mydebuglog.Debug("MD Last Name Parse Error - Input Name: " & FULL_NAME_MD_PARSE & "; Parsed Name: " & temp & "; Parse Result Codes: " & nameObjParseResult)
                        End If
                        nameObjParseResult = String.Join(" ", nameObjParseResult, nameObj.GetResults())
                        tLAST_NAME = temp
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Standardized LAST NAME: " & temp)
                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Unable to standardize Last Name: " & ex.Message)
                    End Try
                    'Parse First Name
                    FULL_NAME_MD_PARSE = FST_NAME & IIf(MID_NAME = "", "", " " & MID_NAME)
                    nameObj.SetPrimaryNameHint(mdName.NameHints.MixedFirstName)
                    If debug = "Y" Then
                        mydebuglog.Debug("  Parsing First and Middle Name... ")
                        mydebuglog.Debug("  FirstName: " & FST_NAME)
                        mydebuglog.Debug("  MidName: " & MID_NAME)
                        mydebuglog.Debug("  NAME_MD_PARSE: " & FULL_NAME_MD_PARSE)
                    End If
                    Try
                        nameObj.SetFullName(FULL_NAME_MD_PARSE)
                        nameObj.Parse()
                        temp = Trim(nameObj.GetFirstName()) & _
                                If(Trim(nameObj.GetMiddleName()) <> "", " " & nameObj.GetMiddleName(), "")
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "MD First and Middle Name Parse Result: " & nameObj.GetResults())
                        If nameObj.GetResults().Contains("NS02") Then
                            myeventlog.Error("MD First and Middle Name Parse Error - Input Name: " & FULL_NAME_MD_PARSE & "; Parsed Name: " & temp & "; Parse Result Codes: " & nameObjParseResult)
                        End If
                        nameObjParseResult = String.Join(" ", nameObjParseResult, nameObj.GetResults())
                        tFST_NAME = Trim(nameObj.GetFirstName())
                        tMID_NAME = Trim(nameObj.GetMiddleName())
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Standardized First and Middle NAME: " & temp)
                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Unable to standardize First and Middle Name: " & ex.Message)
                    End Try
                Else 'Use <FULL_NAME>
                    FULL_NAME_MD_PARSE = FULL_NAME
                    If debug = "Y" Then
                        mydebuglog.Debug("  Parsing Full Name... ")
                        mydebuglog.Debug("  FirstName: " & FST_NAME)
                        mydebuglog.Debug("  MidName: " & MID_NAME)
                        mydebuglog.Debug("  LastName: " & LAST_NAME)
                        mydebuglog.Debug("  Gender: " & GENDER)
                        mydebuglog.Debug("  FullName: " & FULL_NAME)
                        mydebuglog.Debug("  FULL_NAME_MD_PARSE: " & FULL_NAME_MD_PARSE)
                        mydebuglog.Debug("  ConId: " & CON_ID & vbCrLf & "------")
                    End If
                    ' Standardize records using MelissaData
                    Try
                        nameObj.SetFullName(FULL_NAME_MD_PARSE)
                        nameObj.Parse()
                        temp = Trim(nameObj.GetFirstName()) & _
                                            If(Trim(nameObj.GetMiddleName()) <> "", " " & nameObj.GetMiddleName(), "") & _
                                            If(Trim(nameObj.GetLastName()) <> "", " " & nameObj.GetLastName(), "")
                        ' Check Parse results
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "MD Parse Result: " & nameObj.GetResults())
                        If nameObj.GetResults().Contains("NS02") Then
                            myeventlog.Error("MD FullName Parse Error - Input Name: " & FULL_NAME & "; Parsed Name: " & temp & "; Parse Result Codes: " & nameObjParseResult)
                        End If
                        nameObjParseResult = String.Join(" ", nameObjParseResult, nameObj.GetResults())
                        tFST_NAME = Trim(nameObj.GetFirstName()) '& If(Trim(nameObj.GetFirstName2()) <> "", " " & nameObj.GetFirstName2(), "")
                        tMID_NAME = Trim(nameObj.GetMiddleName()) '& If(Trim(nameObj.GetMiddleName2()) <> "", " " & nameObj.GetMiddleName2(), "")
                        tLAST_NAME = Trim(nameObj.GetLastName()) '& If(Trim(nameObj.GetLastName2()) <> "", " " & nameObj.GetLastName2(), "")
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Standardized: " & temp)
                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
                    End Try
                End If
                CON_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("ConId", oNodeList.Item(i))))
                CON_ID = KeySpace(CON_ID)
                database = GetNodeValue("Database", oNodeList.Item(i))
            End If

            ' Determine Gender
            If GENDER = "" Then
                Try
                    'GENDER = Trim(nameObj.GetGender()) & If(Trim(nameObj.GetGender2()) <> "", " " & nameObj.GetGender2(), "")
                    GENDER = If(Trim(nameObj.GetGender()) <> "", Trim(nameObj.GetGender()), nameObj.GetGender2())
                    GENDER = IIf(GENDER = "N", "X", GENDER).ToString()
                Catch ex As Exception
                    If debug = "Y" Then mydebuglog.Debug("  Unable to determine gender: " & ex.Message)
                End Try
            End If

            ' Parse out updated data
            Try
                ' Swap name back if cleaned incorrectly
                If LCase(Trim(tFST_NAME)) = LCase(Trim(MID_NAME)) And tFST_NAME <> "" Then
                    If debug = "Y" Then mydebuglog.Debug("  >>>Swapped<<<")
                    sFST_NAME = tFST_NAME
                    sMID_NAME = tMID_NAME
                    sLAST_NAME = tLAST_NAME
                    tFST_NAME = sLAST_NAME
                    tMID_NAME = sFST_NAME
                    tLAST_NAME = sMID_NAME
                    FULL_NAME = tFST_NAME & " "
                    If tMID_NAME <> "" Then FULL_NAME = FULL_NAME & tMID_NAME & " "
                    FULL_NAME = FULL_NAME & tLAST_NAME
                End If
                If debug = "Y" Then
                    mydebuglog.Debug("  tFirstName: " & tFST_NAME)
                    mydebuglog.Debug("  tMidName: " & tMID_NAME)
                    mydebuglog.Debug("  tLastName: " & tLAST_NAME)
                End If
                If tFST_NAME <> "" Then FST_NAME = tFST_NAME
                If tMID_NAME <> "" Then MID_NAME = tMID_NAME
                If tLAST_NAME <> "" And LCase(tLAST_NAME) <> LCase(FST_NAME) Then LAST_NAME = tLAST_NAME
            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Unable to parse name: " & ex.Message)
            End Try
            'If MATCH_CODE = "" Then
            '    Try
            '        MATCH_CODE = nameObj.GenMatchcodeParsed("NAME", 90, ParseName)
            '    Catch ex As Exception
            '        If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
            '    End Try
            'End If
            'FST_NAME = nameObj.ChangeCase("NAME", CaseType.CASE_PROPER, FST_NAME)
            'MID_NAME = nameObj.ChangeCase("NAME", CaseType.CASE_PROPER, MID_NAME)
            'LAST_NAME = nameObj.ChangeCase("NAME", CaseType.CASE_PROPER, LAST_NAME)
            FST_NAME = Left(FST_NAME, 50)
            MID_NAME = Left(MID_NAME, 50)
            LAST_NAME = Left(LAST_NAME, 50)
            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "------" & vbCrLf & "Contact out & " & vbCrLf & "  FirstName: " & FST_NAME)
                mydebuglog.Debug("  MidName: " & MID_NAME)
                mydebuglog.Debug("  LastName: " & LAST_NAME)
                mydebuglog.Debug("  Gender: " & GENDER)
                mydebuglog.Debug("------")
            End If

            ''*** Skip matchcode 
            'MATCH_CODE = ""
            '********* Create Match Code using MelissaData MatchUp Object in SQL Server; Ren Hou; 1/25/2018  ************
            SqlS = "Select siebeldb.dbo.fnGenMdMatchKey_Contact_HCI(NULL,@1, @2, @3, '', '', '', '', '', '', '')"
            cmd.CommandText = SqlS
            cmd.Parameters.Add("@1", Data.SqlDbType.NVarChar, 4000).Value = FST_NAME
            cmd.Parameters.Add("@2", Data.SqlDbType.NVarChar, 4000).Value = MID_NAME
            cmd.Parameters.Add("@3", Data.SqlDbType.NVarChar, 4000).Value = LAST_NAME
            Try
                MATCH_CODE = cmd.ExecuteScalar()
            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
                myeventlog.Error("  Unable to generate match code: " & ex.Message)
            End Try
            If MATCH_CODE.Contains("Error") Then
                MATCH_CODE = ""
                If debug = "Y" Then mydebuglog.Debug("   .. Error Generating MATCH_CODE from MelissaData ")
                myeventlog.Error("   Error Generating MATCH_CODE from MelissaData: ")
            Else
                If debug = "Y" Then mydebuglog.Debug("   .. Generated MATCH_CODE: " & MATCH_CODE & vbCrLf)
            End If
            ' ***********************************************************************************


            ' ============================================
            ' Database operations

            ' Create record
            If database = "C" Then
                If CON_ID = "" Then
                    ' Generate random contact id
                    CON_ID = LoggingService.GenerateRecordId("S_CONTACT", "N", debug)

                    ' Create contact record with new id
                    SqlS = "INSERT siebeldb.dbo.S_CONTACT (ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,CONFLICT_ID,BU_ID," & _
                    "FST_NAME,LAST_NAME,MID_NAME,SEX_MF,X_MATCH_CD,X_MATCH_DT,LOGIN) " & _
                    "SELECT TOP 1 '" & CON_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0,0,'0-R9NH'," & _
                    "'" & SqlString(FST_NAME) & "','" & SqlString(LAST_NAME) & "','" & SqlString(MID_NAME) & _
                    "','" & GENDER & "','" & MATCH_CODE & "',GETDATE(),'" & SqlString(FULL_NAME) & "') " & _
                    "FROM siebeldb.dbo.S_CONTACT WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & CON_ID & "')"
                    temp = ExecQuery("Create", "Contact record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp

                    ' Verify the record was written
                    SqlS = "SELECT COUNT(*) FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & CON_ID & "'"
                    If debug = "Y" Then mydebuglog.Debug("  Verifying uniqueness: " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    returnv = CheckDBNull(dr(0), enumObjectType.IntType)
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading S_CONTACT: " & ex.ToString & vbCrLf
                                End Try
                            End While
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error reading contact record. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                    If returnv > 0 Then
                        ' Create extension record
                        SqlS = "INSERT INTO siebeldb.dbo.S_CONTACT_X " & _
                        "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM,MODIFICATION_NUM,CONFLICT_ID,PAR_ROW_ID) " & _
                        "VALUES " & _
                        "SELECT TOP 1 '" & CON_ID & "',GETDATE(), '0-1', GETDATE(), '0-1', 0, 0, 0, '" & CON_ID & "' " & _
                        "FROM siebeldb.dbo.S_CONTACT_X WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT_X WHERE ROW_ID='" & CON_ID & "')"
                        temp = ExecQuery("Create", "Contact extension record", cmd, SqlS, mydebuglog, debug)
                        errmsg = errmsg & temp

                        ' Create contact position record
                        SqlS = "INSERT INTO siebeldb.dbo.S_POSTN_CON " & _
                        "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,CONFLICT_ID,CON_FST_NAME, CON_ID, CON_LAST_NAME, POSTN_ID, ROW_STATUS, ASGN_DNRM_FLG, ASGN_MANL_FLG, ASGN_SYS_FLG, STATUS) " & _
                        "VALUES " & _
                        "SELECT TOP 1 '" & CON_ID & "',GETDATE(), '0-1', GETDATE(), '0-1', 0, 0, '" & SqlString(FST_NAME) & "', '" & _
                        CON_ID & "', '" & SqlString(LAST_NAME) & "', '0-5220', 'Y', 'N', 'Y', 'N', 'Active' " & _
                        "FROM siebeldb.dbo.S_POSTN_CON WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_POSTN_CON WHERE ROW_ID='" & CON_ID & "')"
                        temp = ExecQuery("Create", "Contact position record", cmd, SqlS, mydebuglog, debug)
                        errmsg = errmsg & temp
                    End If
                Else
                    errmsg = errmsg & vbCrLf & "Contact Id error on creating record. "
                    results = "Failure"
                End If
            End If

            '-----
            ' Update record
            If database = "U" Then
                If CON_ID <> "" Then
                    SqlS = "UPDATE siebeldb.dbo.S_CONTACT " & _
                    "SET FST_NAME='" & SqlString(FST_NAME) & "', LAST_NAME='" & SqlString(LAST_NAME) & "', " & _
                    "MID_NAME='" & SqlString(MID_NAME) & "', SEX_MF='" & GENDER & "', " & _
                    "X_MATCH_CD='" & MATCH_CODE & "', X_MATCH_DT=GETDATE() " & _
                    "WHERE ROW_ID='" & CON_ID & "'"
                    temp = ExecQuery("Update", "Contact record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                Else
                    errmsg = errmsg & vbCrLf & "Contact Id error on updating record. "
                    results = "Failure"
                End If
            End If
        Next

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Close()
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Return the standardized information as an XML document:
        '   <Contact>
        '       <FirstName>   
        '       <MidName>
        '       <LastName>
        '       <FullName>
        '       <Gender>
        '       <MatchCode>
        '       <ConId>
        '   </Contact>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("Contact")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            ' Modified; Dataflux license expired; 8/1/2017;
            'If debug <> "T" And MATCH_CODE <> "" Then
            If debug <> "T" Then
                '    'If debug <> "T" Then
                AddXMLChild(odoc, resultsRoot, "FirstName", IIf(FST_NAME = "", " ", HttpUtility.UrlEncode(FST_NAME)))
                AddXMLChild(odoc, resultsRoot, "MidName", IIf(MID_NAME = "", " ", HttpUtility.UrlEncode(MID_NAME)))
                AddXMLChild(odoc, resultsRoot, "LastName", IIf(LAST_NAME = "", " ", HttpUtility.UrlEncode(LAST_NAME)))
                AddXMLChild(odoc, resultsRoot, "FullName", IIf(FULL_NAME = "", " ", HttpUtility.UrlEncode(FULL_NAME)))
                AddXMLChild(odoc, resultsRoot, "Gender", IIf(GENDER = "", " ", GENDER))
                AddXMLChild(odoc, resultsRoot, "MatchCode", IIf(MATCH_CODE = "", " ", HttpUtility.UrlEncode(MATCH_CODE)))
                AddXMLChild(odoc, resultsRoot, "ConId", IIf(CON_ID = "", " ", HttpUtility.UrlEncode(CON_ID)))
                'Ren Hou; 2017/09/22; Added MD Parse Result
                ''**** Comment out due to  License Issue *************
                AddXMLChild(odoc, resultsRoot, "MDParseResult", IIf(nameObjParseResult = "", " ", HttpUtility.UrlEncode(nameObjParseResult)))
                'AddXMLChild(odoc, resultsRoot, "MDParseResult", " ")
                '****************************************
            Else
                If MATCH_CODE <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        ' Commented out; Dataflux license expired; 8/1/2017;
        '' Close Dataflux
        'Try
        '    dfsession.Close()
        '    dfsession = Nothing
        '    config = Nothing
        'Catch ex As Exception
        '    errmsg = errmsg & "Problem closing Dataflux. " & ex.ToString
        'End Try

        ' Close MelissaData
        Try
            If Not nameObj Is Nothing Then nameObj.Dispose()
        Catch ex As Exception
            errmsg = errmsg & "Problem closing MelissaData. " & ex.ToString
        End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("StandardizeContact : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("StandardizeContact : Results: " & results & " for " & FST_NAME & " " & MID_NAME & " " & LAST_NAME & " generated matchcode " & MATCH_CODE)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for " & FST_NAME & " " & MID_NAME & " " & LAST_NAME & " generated matchcode " & MATCH_CODE & " at " & Now.ToString)
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "100"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function

    <WebMethod(Description:="Standardizes the supplied organization record and creates it if needed in the database")> _
    Public Function StandardizeOrganization(ByVal sXML As String) As XmlDocument
        ' This function takes an organization supplied, either within the supplied XML or in a 
        ' database record, and standardizes it and generates a match code.  Based on 
        ' parameters, it either returns the data generated or updates the database

        ' The input parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <Organizations>
        '       <Organization>
        '           <Debug>             - A flag to indicate the service is to run in Debug mode or not
        '                                   "Y"  - Yes for debug mode on.. logging on
        '                                   "N"  - No for debug mode off.. logging off
        '                                   "T"  - Test mode on.. logging off
        '           <Database>          - "C" create S_CONTACT record(s), "U" update record, blank do nothing
        '           <OrgId>             - The Id of an existing organization, if applicable
        '           <Name>         	    - Name of organization
        '           <Loc>           	- Location of organization
        '           <FullName>          - Full name of organization
        '       </Organization>
        '   </Organizations>

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim mypath, debug, errmsg, logging As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SODebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim NAME, LOC, MATCH_CODE, FULL_NAME, ORG_ID As String
        Dim temp, database As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        NAME = "Health Communications, Inc."
        LOC = "L"
        MATCH_CODE = ""
        FULL_NAME = ""
        ORG_ID = ""
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0

        ' ============================================
        ' Check parameters
        debug = "N"
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        HttpUtility.UrlDecode(sXML)
        sXML = Regex.Replace(sXML, "&", "&amp;")
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//Organizations/Organization")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        debug = UCase(debug)

        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\standardize_organization_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\StandardizeOrganization.log"
            Try
                log4net.GlobalContext.Properties("SOLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' Commented out; Dataflux license expired; 8/1/2017;
        '' ============================================
        '' Connect to Dataflux
        'Dim dataflux = System.Configuration.ConfigurationManager.AppSettings("dataflux")
        'If dataflux = "" Then dataflux = "datafluxapp1.hq.local"
        'Dim config As New Hashtable
        'config.Add("server", dataflux)
        'config.Add("transport", "TCP")
        'If debug = "Y" Then config.Add("log_file", "C:\Logs\enter.log")

        'Dim dfsession As New DataFlux.dfClient.SessionObject(config)
        'If (dfsession Is Nothing) Then
        '    If debug = "Y" Then mydebuglog.Debug("Unable to open Dataflux")
        '    GoTo CloseOut2
        'Else
        '    If debug = "Y" Then mydebuglog.Debug("  Opening dataflux on " & dataflux)
        'End If

        ' ===========================================
        'MelissaData components (Name Object)
        'MelissaData Initialization
        Dim nameObj As New mdName
        Dim nameObjParseResult As String = ""
        Dim dPath As String = System.Configuration.ConfigurationManager.AppSettings("MD_DataPath")
        Dim dLICENSE As String = System.Configuration.ConfigurationManager.AppSettings("MD_Key")
        Try
            'Set License
            nameObj.SetLicenseString(dLICENSE)
            If Convert.ToDateTime(nameObj.GetLicenseExpirationDate) < Now Then
                If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File")
                errmsg = errmsg & "MelissaData Data License Expired: " & nameObj.GetLicenseExpirationDate
                GoTo CloseOut2
            End If
            'Error Checking
            nameObj.SetPathToNameFiles(dPath)
            If (nameObj.InitializeDataFiles() <> mdName.ProgramStatus.NoError) Then
                If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File: " + nameObj.GetInitializeErrorString())
                errmsg = errmsg & "Unable to Initiate MelissaData Data File: " + nameObj.GetInitializeErrorString()
                GoTo CloseOut2
            Else
                If debug = "Y" Then mydebuglog.Debug("MelissaData Data File Initialized")
                'InitErrorString = addrObj.GetInitializeErrorString
                'DatabaseDate = addrObj.GetDatabaseDate
                'ExpDate = addrObj.GetExpirationDate
                'BuildNum = addrObj.GetBuildNumber
            End If
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File: " + ex.ToString())
            errmsg = errmsg & "Unable to Initiate MelissaData Data File: " + ex.ToString()
        End Try

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try
        If debug = "Y" Then
            Try
                mydebuglog.Debug(vbCrLf & "Session-")
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        For i = 0 To oNodeList.Count - 1
            errmsg = ""
            If debug <> "T" Then
                NAME = Trim(HttpUtility.UrlDecode(GetNodeValue("Name", oNodeList.Item(i))))
                LOC = Trim(HttpUtility.UrlDecode(GetNodeValue("Loc", oNodeList.Item(i))))
                ORG_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("OrgId", oNodeList.Item(i))))
                ORG_ID = KeySpace(ORG_ID)
                database = GetNodeValue("Database", oNodeList.Item(i))
            End If
            If debug = "Y" Then
                mydebuglog.Debug("  Name: " & NAME)
                mydebuglog.Debug("  Loc: " & LOC)
                mydebuglog.Debug("  OrgId: " & ORG_ID)
            End If

            ' Commented out; Dataflux license expired; 8/1/2017;
            '' Standardize records   
            'Try
            '    temp = dfsession.Standardize("ORG", NAME)
            '    NAME = temp
            '    temp = dfsession.Standardize("ORG", LOC)
            '    LOC = temp
            '    If debug = "Y" Then mydebuglog.Debug("  Standardized NAME: " & NAME)
            'Catch ex As Exception
            '    If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
            'End Try

            ' Generate match code       
            'Try
            'FULL_NAME = NAME & " " & LOC
            'MATCH_CODE = dfsession.GenMatchcode("ORG", 90, FULL_NAME)
            'Catch ex As Exception
            '    If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
            'End Try

            '' Propercase organization
            'Try
            '    FULL_NAME = dfsession.ChangeCase("ORG", CaseType.CASE_PROPER, FULL_NAME)
            '    NAME = dfsession.ChangeCase("ORG", CaseType.CASE_PROPER, NAME)
            '    NAME = Left(NAME, 100)
            '    LOC = dfsession.ChangeCase("ORG", CaseType.CASE_PROPER, LOC)
            '    LOC = Left(LOC, 50)
            'Catch ex As Exception
            '    If debug = "Y" Then mydebuglog.Debug("  Unable to change case: " & ex.Message)
            'End Try
            'If UCase(Left(NAME, 3)) = "PF " Then NAME = "PF" & Mid(NAME, 3)
            'If debug = "Y" Then mydebuglog.Debug("  Changecase NAME: " & NAME)

            ' Standardize records using MelissaData
            Try
                temp = nameObj.StandardizeCompany(NAME)
                NAME = temp
                'nameObj.SetFullName(LOC)
                'nameObj.Parse()
                'temp = If(Trim(nameObj.GetFirstName()) <> "", " " & nameObj.GetFirstName(), "") & _
                '        If(Trim(nameObj.GetMiddleName()) <> "", " " & nameObj.GetMiddleName(), "") & _
                '        If(Trim(nameObj.GetLastName()) <> "", " " & nameObj.GetLastName(), "")
                temp = nameObj.StandardizeCompany(LOC)
                LOC = temp
                ' Check Parse results
                nameObjParseResult = nameObj.GetResults()
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "MD Parse Result: " & nameObjParseResult)
                If nameObjParseResult.Contains("NS02") Then
                    myeventlog.Error("MD StandardizeCompany Error - Input Org Name: " & FULL_NAME & "; Parsed Org Name: " & temp & "; Parse Result Codes: " & nameObjParseResult)
                End If
                FULL_NAME = NAME & " " & LOC
                If debug = "Y" Then mydebuglog.Debug("  Standardized NAME: " & NAME)
            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
            End Try

            '' Generate match code       
            'Try
            '    FULL_NAME = NAME & " " & LOC
            '    MATCH_CODE = nameObj.GenMatchcode("ORG", 90, FULL_NAME)
            'Catch ex As Exception
            '    If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
            'End Try

            ' Propercase organization
            Try
                'NAME = nameObj.ChangeCase("ORG", CaseType.CASE_PROPER, NAME) 
                NAME = Left(NAME, 100)
                'LOC = nameObj.ChangeCase("ORG", CaseType.CASE_PROPER, LOC) 
                LOC = Left(LOC, 50)
                'FULL_NAME = nameObj.ChangeCase("ORG", CaseType.CASE_PROPER, FULL_NAME)
                FULL_NAME = NAME & " " & LOC
            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Unable to change case: " & ex.Message)
            End Try
            If UCase(Left(NAME, 3)) = "PF " Then NAME = "PF" & Mid(NAME, 3)
            If debug = "Y" Then mydebuglog.Debug("  Changecase NAME: " & NAME)

            ''*** Skip matchcode 
            'MATCH_CODE = ""
            '********* Create Match Code using MelissaData MatchUp Object in SQL Server; Ren Hou; 1/25/2018  ************
            SqlS = "Select siebeldb.dbo.fnGenMdMatchKey_Company_HCI(NULL,@1, @2, '', '', '', '', '', '', '', '')"
            cmd.CommandText = SqlS
            cmd.Parameters.Add("@1", Data.SqlDbType.NVarChar, 4000).Value = NAME
            cmd.Parameters.Add("@2", Data.SqlDbType.NVarChar, 4000).Value = LOC
            Try
                MATCH_CODE = cmd.ExecuteScalar()
            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
                myeventlog.Error("  Unable to generate match code: " & ex.Message)
            End Try
            If MATCH_CODE.Contains("Error") Then
                MATCH_CODE = ""
                If debug = "Y" Then mydebuglog.Debug("   .. Error Generating MATCH_CODE from MelissaData ")
                myeventlog.Error("   Error Generating MATCH_CODE from MelissaData: ")
            Else
                If debug = "Y" Then mydebuglog.Debug("   .. Generated MATCH_CODE: " & MATCH_CODE & vbCrLf)
            End If
            ' ***********************************************************************************


            ' ============================================
            ' Database operations
            ' Create record
            If database = "C" Then
                If ORG_ID = "" Then
                    ' Generate random organization id
                    ORG_ID = LoggingService.GenerateRecordId("S_ORG_EXT", "N", debug)

                    ' Create org record with new id
                    SqlS = "INSERT INTO siebeldb.dbo.S_ORG_EXT " & _
                    "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM,MODIFICATION_NUM,CONFLICT_ID,BU_ID," & _
                    "DISA_CLEANSE_FLG,NAME,LOC, PROSPECT_FLG,PRTNR_FLG,ENTERPRISE_FLAG,LANG_ID,BASE_CURCY_CD," & _
                    "CREATOR_LOGIN,CUST_STAT_CD,DESC_TEXT,DISA_ALL_MAILS_FLG,FRGHT_TERMS_CD,MAIN_FAX_PH_NUM,MAIN_PH_NUM," & _
                    "PR_POSTN_ID,X_DATAFLEX_FLG,X_ACCOUNT_NUM, PR_ADDR_ID, PR_BL_ADDR_ID, PR_SHIP_ADDR_ID, PR_BL_PER_ID, " & _
                    "PR_SHIP_PER_ID, DEDUP_TOKEN, X_MATCH_DT) VALUES " & _
                    "('" & ORG_ID & "', getdate(), '0-1', getdate(), '0-1', 0, 0, 0, '0-R9NH', " & _
                    "'N', '" & SqlString(NAME) & "','" & SqlString(LOC) & "', 'Y', 'N', 'Y', 'ENU', 'USD', " & _
                    "'SADMIN', 'Prospect', 'StandardizeOrganization service', 'N', 'FOB', '', '', " & _
                    "'0-5220', 'N', " & Right(ORG_ID, Len(ORG_ID) - 2) & ", '', '', '', '', '','" & MATCH_CODE & "', GETDATE())"
                    Try
                        temp = ExecQuery("Create", "Org record", cmd, SqlS, mydebuglog, debug)
                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Insert account error: " & ex.ToString)
                        temp = "Insert Org record error"
                    End Try
                    errmsg = errmsg & temp

                    ' Verify the record was written
                    SqlS = "SELECT COUNT(*) FROM siebeldb.dbo.S_ORG_EXT WHERE ROW_ID='" & ORG_ID & "'"
                    If debug = "Y" Then mydebuglog.Debug("  Verifying uniqueness: " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    returnv = CheckDBNull(dr(0), enumObjectType.IntType)
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading S_ORG_EXT: " & ex.ToString & vbCrLf
                                End Try
                            End While
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error reading org record. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                    If debug = "Y" Then mydebuglog.Debug("  Records found: " & returnv.ToString)

                    ' Create account position record
                    If returnv > 0 Then
                        ' Create account position record
                        SqlS = "INSERT INTO siebeldb.dbo.S_ACCNT_POSTN " & _
                        "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM," & _
                        "CONFLICT_ID,ACCNT_NAME,OU_EXT_ID,POSITION_ID,ROW_STATUS)  VALUES " & _
                        "('" & ORG_ID & "', getdate(), '0-1', getdate(), '0-1', 0, " & _
                        "0, '" & NAME & "', '" & ORG_ID & "', '0-5220', 'N')"
                        temp = ExecQuery("Create", "Org position record", cmd, SqlS, mydebuglog, debug)
                        errmsg = errmsg & temp
                    End If
                Else
                    errmsg = errmsg & vbCrLf & "Org Id error on creating record. "
                    results = "Failure"
                End If
            End If

            '-----
            ' Update record
            If database = "U" Then
                If ORG_ID <> "" Then
                    SqlS = "UPDATE siebeldb.dbo.S_ORG_EXT " & _
                    "SET NAME='" & SqlString(NAME) & "', LOC='" & SqlString(LOC) & "', " & _
                    "DEDUP_TOKEN='" & MATCH_CODE & "', X_MATCH_DT=GETDATE() " & _
                    "WHERE ROW_ID='" & ORG_ID & "'"
                    temp = ExecQuery("Update", "Org record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                Else
                    errmsg = errmsg & vbCrLf & "Org Id error on updating record. "
                    results = "Failure"
                End If
            End If
        Next

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Close()
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Return the standardized information as an XML document:
        '   <Organization>
        '       <Name>   
        '       <Loc>
        '       <FullName>
        '       <MatchCode>
        '       <OrgId>
        '   </Organization>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("Organization")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            ' Modified; Dataflux license expired; 8/1/2017;
            'If debug <> "T" And MATCH_CODE <> "" Then
            If debug <> "T" Then
                AddXMLChild(odoc, resultsRoot, "Name", IIf(NAME = "", " ", HttpUtility.UrlEncode(NAME)))
                AddXMLChild(odoc, resultsRoot, "Loc", IIf(LOC = "", " ", HttpUtility.UrlEncode(LOC)))
                AddXMLChild(odoc, resultsRoot, "FullName", IIf(FULL_NAME = "", " ", HttpUtility.UrlEncode(FULL_NAME)))
                AddXMLChild(odoc, resultsRoot, "MatchCode", IIf(MATCH_CODE = "", " ", HttpUtility.UrlEncode(MATCH_CODE)))
                AddXMLChild(odoc, resultsRoot, "OrgId", IIf(ORG_ID = "", " ", HttpUtility.UrlEncode(ORG_ID)))
                'Ren Hou; 2017/09/22; Added MD Parse Result
                AddXMLChild(odoc, resultsRoot, "MDParseResult", IIf(nameObjParseResult = "", " ", HttpUtility.UrlEncode(nameObjParseResult)))
            Else
                If MATCH_CODE <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
                End If
                If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

        ' Commented out; Dataflux license expired; 8/1/2017;
        '' Close Dataflux
        'Try
        '    If Not dfsession Is Nothing Then dfsession.Close()
        '    dfsession = Nothing
        '    config = Nothing
        'Catch ex As Exception
        '    errmsg = errmsg & "Problem closing Dataflux. " & ex.ToString
        'End Try

        ' Close MelissaData
        Try
            If Not nameObj Is Nothing Then nameObj.Dispose()
        Catch ex As Exception
            errmsg = errmsg & "Problem closing MelissaData. " & ex.ToString
        End Try
        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("StandardizeOrganization : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("StandardizeOrganization : Results: " & results & " for '" & NAME & "' generated matchcode " & MATCH_CODE)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for '" & NAME & "' generated matchcode " & MATCH_CODE & " at " & Now.ToString)
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "100"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function

    <WebMethod(Description:="Standardizes and geoencodes a provided address when directed and creates it if needed in the database")> _
    Public Function StandardizeAddress_SAS(ByVal sXML As String) As XmlDocument
        ' This function takes an address supplied, either within the supplied XML or in a 
        ' database record, and standardizes it and generates a match code.  Based on 
        ' parameters, it either returns the data generated or updates the database

        ' The input parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <AddressList>
        '       <AddressRec>
        '           <Debug>             - A flag to indicate the service is to run in Debug mode or not
        '                                   "Y"  - Yes for debug mode on.. logging on
        '                                   "N"  - No for debug mode off.. logging off
        '                                   "T"  - Test mode on.. logging off
        '           <Database>          - "C" create S_ADDR_ORG record(s), "U" update record, "V"erification turned off
        '           <AddrId>            - The Id of an existing session, if applicable
        '           <OrgId>             - Organization Id for organizational addresses - optional
        '           <ConId>             - Contact Id for personal addresses - optional
        '           <Type>              - "O"rganization or "P"ersonal address
        '           <GeoCode>           - Geocode address flag ("Y" or "N") - optional
        '           <Address>           - Street address
        '           <City>              - City
        '           <State>             - State
        '           <County>            - County
        '           <Zipcode>           - Zipcode
        '           <Country>           - Country
        '       </AddressRec>
        '   </AddressList>

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim mypath, debug, errmsg, logging, oxml As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SADebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim ADDR, CITY, STATE, ZIPCODE, COUNTY, COUNTRY, COUNTRY2, GEOCODEADDR, ADDR_TYPE, ADDR_ID, MATCH_CODE As String
        Dim temp, database, lastline, match2, LAT, LON, ORG_ID, CON_ID, JURIS_ID, wp As String
        Dim deliverable, geocodio_key As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        ADDR_ID = ""
        ADDR_TYPE = "O"
        GEOCODEADDR = "N"
        ADDR = "1101 Wilson Suite 1700"
        CITY = "Arlington"
        STATE = "VA"
        COUNTY = ""
        ZIPCODE = ""
        COUNTRY = ""
        COUNTRY2 = ""
        MATCH_CODE = ""
        ORG_ID = ""
        CON_ID = ""
        JURIS_ID = ""
        match2 = ""
        LAT = ""
        LON = ""
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0
        wp = ""
        deliverable = ""

        ' ============================================
        ' Check parameters
        debug = "N"
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        oxml = sXML
        HttpUtility.UrlDecode(sXML)
        sXML = Regex.Replace(sXML, "&", "&amp;")
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//AddressList/AddressRec")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        debug = UCase(debug)

        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\standardize_address_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\StandardizeAddress.log"
            Try
                log4net.GlobalContext.Properties("SALogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & oxml)
            End If
        End If

        ' Commented out; Dataflux license expired; 8/1/2017;
        '' ============================================
        '' Connect to Dataflux
        'Dim dataflux = System.Configuration.ConfigurationManager.AppSettings("dataflux")
        'If dataflux = "" Then dataflux = "datafluxapp1.hq.local"
        'Dim config As New Hashtable
        'config.Add("server", dataflux)
        'config.Add("transport", "TCP")
        'If debug = "Y" Then config.Add("log_file", "C:\Logs\enter.log")

        'Dim dfsession As New DataFlux.dfClient.SessionObject(config)
        'If (dfsession Is Nothing) Then
        '    If debug = "Y" Then mydebuglog.Debug("Unable to open Dataflux")
        '    GoTo CloseOut2
        'Else
        '    If debug = "Y" Then mydebuglog.Debug("  Opening dataflux on " & dataflux)
        'End If

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try
        If debug = "Y" Then
            Try
                mydebuglog.Debug(vbCrLf & "Session-")
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Geocoding key 
        geocodio_key = System.Configuration.ConfigurationManager.AppSettings.Get("geocodio_key")
        If geocodio_key = "" Then
            geocodio_key = "4c56e20d4de8d8b5f2aa9a4851145221d9595ed"
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        For i = 0 To oNodeList.Count - 1
            errmsg = ""
            If debug <> "T" Then
                ADDR_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("AddrId", oNodeList.Item(i))))
                ADDR_ID = KeySpace(ADDR_ID)
                ORG_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("OrgId", oNodeList.Item(i))))
                ORG_ID = KeySpace(ORG_ID)
                CON_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("ConId", oNodeList.Item(i))))
                CON_ID = KeySpace(CON_ID)
                ADDR_TYPE = Left(GetNodeValue("Type", oNodeList.Item(i)), 1)
                If ADDR_TYPE = "" And ORG_ID <> "" Then ADDR_TYPE = "O"
                If ADDR_TYPE = "" And CON_ID <> "" Then ADDR_TYPE = "P"
                GEOCODEADDR = Left(GetNodeValue("GeoCode", oNodeList.Item(i)), 1)
                If GEOCODEADDR <> "Y" Then GEOCODEADDR = "N"
                ADDR = Trim(HttpUtility.UrlDecode(GetNodeValue("Address", oNodeList.Item(i))))
                ADDR = CleanString(ADDR)
                ADDR = Left(ADDR, 200)
                CITY = Trim(HttpUtility.UrlDecode(GetNodeValue("City", oNodeList.Item(i))))
                CITY = CleanString(CITY)
                CITY = Left(CITY, 50)
                STATE = Trim(HttpUtility.UrlDecode(GetNodeValue("State", oNodeList.Item(i))))
                STATE = RemoveSymbols(STATE)
                STATE = Left(STATE, 10)
                COUNTY = Trim(HttpUtility.UrlDecode(GetNodeValue("County", oNodeList.Item(i))))
                COUNTY = CleanString(COUNTY)
                COUNTY = Left(COUNTY, 50)
                ZIPCODE = Trim(HttpUtility.UrlDecode(GetNodeValue("Zipcode", oNodeList.Item(i))))
                ZIPCODE = RemoveSymbols(ZIPCODE)
                ZIPCODE = Left(ZIPCODE, 30)
                COUNTRY = Trim(HttpUtility.UrlDecode(GetNodeValue("Country", oNodeList.Item(i))))
                COUNTRY = RemoveSymbols(COUNTRY)
                COUNTRY = Left(COUNTRY, 30)
                database = Trim(Left(GetNodeValue("Database", oNodeList.Item(i)), 1))
            End If
            If debug = "Y" Then
                mydebuglog.Debug("INPUTS------" & vbCrLf & "  ADDR_ID: " & ADDR_ID)
                mydebuglog.Debug("  ORG_ID: " & ORG_ID)
                mydebuglog.Debug("  CON_ID: " & CON_ID)
                mydebuglog.Debug("  ADDR_TYPE: " & ADDR_TYPE)
                mydebuglog.Debug("  GEOCODEADDR: " & GEOCODEADDR)
                mydebuglog.Debug("  ADDR: " & ADDR)
                mydebuglog.Debug("  CITY: " & CITY)
                mydebuglog.Debug("  STATE: " & STATE)
                mydebuglog.Debug("  COUNTY: " & COUNTY)
                mydebuglog.Debug("  ZIPCODE: " & ZIPCODE)
                mydebuglog.Debug("  COUNTRY: " & COUNTRY)
                mydebuglog.Debug("  database: " & database & vbCrLf & "------------")
            End If

            ' -----
            ' Commented out; Dataflux license expired; 8/1/2017;
            '' Process Address Line
            'If database <> "V" Then
            '    Try
            '        temp = dfsession.Standardize("ADDR", ADDR)
            '        ADDR = temp
            '        If debug = "Y" Then mydebuglog.Debug("  Standard address: " & temp)
            '    Catch ex As Exception
            '        If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
            '    End Try
            'Else
            '    If debug = "Y" Then mydebuglog.Debug("  *Address standardization disabled*")
            'End If

            '' Generate match code
            'Try
            '    MATCH_CODE = dfsession.GenMatchcode("ADDR", 90, ADDR)
            '    If debug = "Y" Then mydebuglog.Debug("   .. Generated MATCH_CODE: " & MATCH_CODE & vbCrLf)
            'Catch ex As Exception
            '    If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
            'End Try

            '' -----
            '' Process City Line
            'lastline = CITY & ", " & STATE & " " & ZIPCODE
            'Try
            '    temp = dfsession.Standardize("LAST_LINE", lastline)
            '    lastline = temp
            '    If debug = "Y" Then mydebuglog.Debug("  Standard last line: " & lastline)
            'Catch ex As Exception
            '    If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
            'End Try

            '' Verify address if Verification not turned off
            'If database <> "V" Then
            '    Dim verified As Hashtable
            '    Try
            '        'verified = dfsession.TransformAddress(AddrFlags.ADDR_US_CASS, "", ADDR, "", lastline, "US")
            '        verified = dfsession.VerifyAddress(ADDR, "", lastline, "", "", "US")
            '        deliverable = verified.Item("deliverability")

            '        ' Parse out address fields from last line
            '        If Trim(verified.Item("address_line1")) <> "" Then
            '            ADDR = dfsession.ChangeCase("NAME", CaseType.CASE_PROPER, verified.Item("address_line1"))
            '            If Trim(verified.Item("address_line2")) <> "" Then
            '                ADDR = ADDR & " " & dfsession.ChangeCase("NAME", CaseType.CASE_PROPER, verified.Item("address_line2"))
            '            End If
            '        End If
            '        ZIPCODE = verified.Item("postal_code")
            '        STATE = verified.Item("state")
            '        CITY = dfsession.ChangeCase("NAME", CaseType.CASE_PROPER, verified.Item("city"))
            '        COUNTY = dfsession.ChangeCase("NAME", CaseType.CASE_PROPER, verified.Item("county"))
            '        COUNTRY = verified.Item("country_code")
            '        lastline = CITY & ", " & STATE & " " & ZIPCODE
            '        verified = Nothing
            '        If debug = "Y" Then
            '            mydebuglog.Debug("   .. Verified addr: " & ADDR)
            '            mydebuglog.Debug("   .. Verified postal_code: " & ZIPCODE)
            '            mydebuglog.Debug("   .. Verified deliverability: " & deliverable)
            '        End If
            '    Catch ex As Exception
            '        If debug = "Y" Then mydebuglog.Debug("  Unable to verify address: " & ex.Message)
            '    End Try
            'Else
            '    If debug = "Y" Then mydebuglog.Debug("  *Address verification disabled*")
            '    CITY = dfsession.ChangeCase("NAME", CaseType.CASE_PROPER, CITY)
            '    COUNTY = dfsession.ChangeCase("NAME", CaseType.CASE_PROPER, COUNTY)
            'End If

            ' -----
            ' Set country code
            If COUNTRY = "US" Then
                COUNTRY = "USA"
                COUNTRY2 = "US"
            End If

            ' Lookup 3-character country code
            If Len(COUNTRY) = 2 Then
                COUNTRY2 = COUNTRY
                SqlS = "SELECT VAL FROM siebeldb.dbo.S_LST_OF_VAL WHERE TYPE='COUNTRY_CODE' AND CODE='" & COUNTRY & "'"
                If debug = "Y" Then mydebuglog.Debug("  Get 3-char country code: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                If Len(temp) = 3 Then COUNTRY = temp
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading country code: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error locating country code. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If debug = "Y" Then mydebuglog.Debug("   .. 3-char country code: " & COUNTRY)
            End If

            ' Lookup 2-character code if Geoencoding
            If COUNTRY2 = "" And GEOCODEADDR = "Y" Then
                SqlS = "SELECT CODE FROM siebeldb.dbo.S_LST_OF_VAL WHERE TYPE='COUNTRY_CODE' AND VAL='" & COUNTRY & "'"
                If debug = "Y" Then mydebuglog.Debug("  Get 2-char country code: " & SqlS & vbCrLf)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                If Len(temp) = 2 Then COUNTRY2 = temp
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading country code: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error locating country code. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If debug = "Y" Then mydebuglog.Debug("   .. 2-char country code: " & COUNTRY2 & vbCrLf)
            End If

            ' -----
            ' Commented out; Dataflux license expired; 8/1/2017;
            '' Generate match code       
            'Try
            '    match2 = dfsession.GenMatchcode("LAST_LINE", 90, lastline)
            'Catch ex As Exception
            '    If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
            'End Try
            'MATCH_CODE = MATCH_CODE & match2
            'If debug = "Y" Then mydebuglog.Debug("   .. Generated match2: " & match2 & vbCrLf)

            ' -----
            ' Geoencode address if applicable
            If GEOCODEADDR = "Y" Then

                ' Check to see if this was already done - if applicable
                If ADDR_ID <> "" Then
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "SELECT X_LAT, X_LONG " & _
                            "FROM siebeldb.dbo.S_ADDR_PER " & _
                            "WHERE ROW_ID='" & ADDR_ID & "'"
                        Case Else
                            SqlS = "SELECT X_LAT, X_LONG " & _
                            "FROM siebeldb.dbo.S_ADDR_ORG " & _
                            "WHERE ROW_ID='" & ADDR_ID & "'"
                    End Select
                    If debug = "Y" Then mydebuglog.Debug("  Get existing Latitude: " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    temp = Str(CheckDBNull(dr(0), enumObjectType.DblType))
                                    If Len(temp) > 0 And Val(temp) <> 0 Then LAT = temp
                                    temp = Str(CheckDBNull(dr(1), enumObjectType.DblType))
                                    If Len(temp) > 0 And Val(temp) <> 0 Then LON = temp
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading latitude: " & ex.ToString & vbCrLf
                                End Try
                            End While
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error locating country code. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                    If debug = "Y" Then mydebuglog.Debug("   .. Existing latitude: " & LAT & vbCrLf)
                End If

                ' Using Yahoo
                'temp = ""
                'Dim addrstr As String
                'Dim oAuth = New OAuthBase
                'addrstr = HttpUtility.UrlEncode(ADDR) & ",+" & HttpUtility.UrlEncode(CITY) & ",+" & HttpUtility.UrlEncode(STATE) & ",+" & HttpUtility.UrlEncode(ZIPCODE)
                'Dim geoservice As New OAuth.Geoencode
                'temp = geoservice.Get(addrstr, LAT, LON)
                'If debug = "Y" Then mydebuglog.Debug("  Geocoding Results: " & LAT & ", " & LON)
                'geoservice = Nothing
                'oAuth = Nothing

                ' Using BING
                'Try
                'Dim Geocoder As New GeocodeRequest()

                ' Set the credentials using a valid Bing Maps Key
                'Dim GeoCredentials As New Credentials()
                'Dim bingkey As String
                'bingkey = "Amz0YaXhXmYNpChX0g7cWl5qYB6qOevOjebXt-EOMPuRtXRy-CRoAnIEVMJc3RPB"
                'Geocoder.Credentials = New Credentials() With {.ApplicationId = bingkey}

                ' Set address to query
                'Dim addrstr As String
                'addrstr = ADDR & ", " & CITY & ", " & STATE & " " & ZIPCODE
                'If debug = "Y" Then mydebuglog.Debug("  Geocoding Address: " & addrstr)
                'Geocoder.Query = addrstr

                ' Set the options to only return high confidence results
                'Dim filters() As ConfidenceFilter = {New ConfidenceFilter() With {.MinimumConfidence = Confidence.Medium}}

                'Dim geocodeOptions As New GeocodeOptions() _
                'With {.Filters = filters}

                'Geocoder.Options = geocodeOptions

                ' Make the geocode request
                'Dim GeocodeService As New GeocodeServiceClient("BasicHttpBinding_IGeocodeService")
                'Dim geocodeResponse = GeocodeService.Geocode(Geocoder)

                ' Use the results in your application.
                'results = geocodeResponse.Results(0).DisplayName
                'If results.Length > 0 Then
                '    LAT = geocodeResponse.Results(0).Locations(0).Latitude
                '    LON = geocodeResponse.Results(0).Locations(0).Longitude
                'End If
                'If debug = "Y" Then mydebuglog.Debug("  Geocoding Results: " & LAT & ", " & LON & vbCrLf)

                ' Remove objects created
                'Geocoder = Nothing
                'GeoCredentials = Nothing
                'GeocodeService = Nothing
                'geocodeResponse = Nothing

                'Catch ex As Exception
                'If debug = "Y" Then mydebuglog.Debug("  Unable to geocode: " & ex.Message)
                'End Try

                ' Using Geocodio
                If Trim(ADDR) <> "" And (LAT = "" Or LON = "") Then
                    Try
                        Dim JsonSerial As New JavaScriptSerializer
                        Dim http As New simplehttp()

                        ' Prepare URL
                        Dim SvcURL As String
                        Dim addrstr As String
                        SvcURL = System.Configuration.ConfigurationManager.AppSettings("GeocodeUrl")
                        addrstr = "street=" & Replace(Trim(ADDR), " ", "+") & "&city=" & Replace(Trim(CITY), " ", "+") & "&state=" & Replace(Trim(STATE), " ", "+") & "&postal_code=" & Replace(Trim(ZIPCODE), " ", "+")
                        addrstr = addrstr & "&api_key=" & geocodio_key
                        If debug = "Y" Then mydebuglog.Debug("  Geocode SvcURL: " & SvcURL & addrstr)

                        ' Generate results
                        Dim georesults As String
                        georesults = http.geturl(SvcURL & addrstr, System.Configuration.ConfigurationManager.AppSettings("Geocode_proxyIP"), 80, "", "")
                        If results.Length > 0 Then

                            ' Deserialize
                            Dim JsonObj As GeocodioObj = JsonSerial.Deserialize(Of GeocodioObj)(georesults)

                            ' Locate LAT/LON
                            If JsonObj.results.Length > 0 Then
                                LAT = JsonObj.results(0).Location.Lat
                                If debug = "Y" Then mydebuglog.Debug("   .. Geocode LAT: " & LAT.ToString)
                                LON = JsonObj.results(0).Location.Lng
                            End If

                            JsonObj = Nothing
                        End If

                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Unable to geocode: " & ex.Message)
                    End Try
                End If
            End If

            ' -----
            ' Get address jurisdiction
            If JURIS_ID = "" Then
                temp = ""
                wp = "<Address><debug>N</debug><regulated></regulated>"
                wp = wp & "<street></street>"
                wp = wp & "<city>" & CITY & "</city>"
                wp = wp & "<state>" & STATE & "</state>"
                wp = wp & "<county>" & COUNTY & "</county>"
                wp = wp & "<zipcode>" & ZIPCODE & "</zipcode>"
                wp = wp & "<country>" & COUNTRY & "</country></Address>"
                JURIS_ID = LoggingService.FindJurisdiction(wp)
                If debug = "Y" Then mydebuglog.Debug("  FindJurisdiction: " & wp)
                If debug = "Y" Then mydebuglog.Debug("   .. JURIS_ID match2: " & JURIS_ID & vbCrLf)
            End If

            ' ============================================
            ' Database operations
            ' Create record
            If database = "C" Then
                If ADDR_ID = "" Then
                    If ADDR_TYPE = "P" And CON_ID = "" Then GoTo UpdateAddr ' Skip if personal and no contact id
                    If ADDR_TYPE = "O" And ORG_ID = "" Then GoTo UpdateAddr ' Skip if organizational and no organization id
GenerateID:
                    ' Generate random address id
                    Select Case ADDR_TYPE
                        Case "P"
                            ADDR_ID = LoggingService.GenerateRecordId("S_ADDR_PER", "N", debug)
                        Case "O"
                            ADDR_ID = LoggingService.GenerateRecordId("S_ADDR_ORG", "N", debug)
                    End Select

                    ' Create address record with new id
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "INSERT INTO siebeldb.dbo.S_ADDR_PER " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM," & _
                            "MODIFICATION_NUM,CONFLICT_ID,DISA_CLEANSE_FLG,PER_ID,ADDR,CITY,COMMENTS," & _
                            "COUNTY,COUNTRY,STATE,ZIPCODE,X_MATCH_CD," & _
                            "X_MATCH_DT,X_LAT,X_LONG,X_JURIS_ID,X_CASS_CHECKED,X_CASS_CODE) " & _
                            "VALUES " & _
                            "('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0," & _
                            "0,0,'N','" & CON_ID & "','" & ADDR & "','" & CITY & "','From StandardizeAddress', '" & _
                            COUNTY & "','" & COUNTRY & "','" & STATE & "', '" & ZIPCODE & "','" & MATCH_CODE & _
                            "',GETDATE(),'" & LAT & "','" & LON & "','" & JURIS_ID & "',GETDATE(),'" & deliverable & ")"
                        Case "O"
                            SqlS = "INSERT INTO siebeldb.dbo.S_ADDR_ORG " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM," & _
                            "MODIFICATION_NUM,CONFLICT_ID,DISA_CLEANSE_FLG,OU_ID,ADDR,CITY,COMMENTS," & _
                            "COUNTY,COUNTRY,STATE,ZIPCODE,X_MATCH_CD," & _
                            "X_MATCH_DT,X_LAT,X_LONG,X_JURIS_ID,X_CASS_CHECKED,X_CASS_CODE) " & _
                            "VALUES " & _
                            "('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0," & _
                            "0,0,'N','" & ORG_ID & "','" & ADDR & "','" & CITY & "','From StandardizeAddress', '" & _
                            COUNTY & "','" & COUNTRY & "','" & STATE & "', '" & ZIPCODE & "','" & MATCH_CODE & _
                            "',GETDATE(),'" & LAT & "','" & LON & "','" & JURIS_ID & "',GETDATE(),'" & deliverable & ")"
                    End Select
                    temp = ExecQuery("Create", "Address record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                Else
                    errmsg = errmsg & vbCrLf & "Address Id error on creating record. "
                    results = "Failure"
                End If
            End If

            '-----
            ' Update record
UpdateAddr:
            If database = "U" Then
                If ADDR_ID <> "" Then
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "UPDATE siebeldb.dbo.S_ADDR_PER SET LAST_UPD=GETDATE()," & _
                            "ADDR='" & SqlString(ADDR) & "',CITY='" & SqlString(CITY) & "',STATE='" & SqlString(STATE) & "',COUNTRY='" & COUNTRY & "'," & _
                            "ZIPCODE='" & ZIPCODE & "',X_MATCH_CD='" & MATCH_CODE & "',X_MATCH_DT=GETDATE(),X_CASS_CHECKED=GETDATE(),X_CASS_CODE='" & deliverable & "'"
                            If COUNTY <> "" Then SqlS = SqlS & ",COUNTY='" & SqlString(COUNTY) & "'"
                            If LAT <> "" Then SqlS = SqlS & ",X_LAT='" & LAT & "'"
                            If LON <> "" Then SqlS = SqlS & ",X_LONG='" & LON & "'"
                            If JURIS_ID <> "" Then SqlS = SqlS & ",X_JURIS_ID='" & JURIS_ID & "'"
                            SqlS = SqlS & " WHERE ROW_ID='" & ADDR_ID & "'"
                        Case "O"
                            SqlS = "UPDATE siebeldb.dbo.S_ADDR_ORG SET LAST_UPD=GETDATE()," & _
                            "ADDR='" & SqlString(ADDR) & "',CITY='" & SqlString(CITY) & "',STATE='" & SqlString(STATE) & "',COUNTRY='" & COUNTRY & "'," & _
                            "ZIPCODE='" & ZIPCODE & "',X_MATCH_CD='" & MATCH_CODE & "',X_MATCH_DT=GETDATE(),X_CASS_CHECKED=GETDATE(),X_CASS_CODE='" & deliverable & "'"
                            If COUNTY <> "" Then SqlS = SqlS & ",COUNTY='" & SqlString(COUNTY) & "'"
                            If LAT <> "" Then SqlS = SqlS & ",X_LAT='" & LAT & "'"
                            If LON <> "" Then SqlS = SqlS & ",X_LONG='" & LON & "'"
                            If JURIS_ID <> "" Then SqlS = SqlS & ",X_JURIS_ID='" & JURIS_ID & "'"
                            SqlS = SqlS & " WHERE ROW_ID='" & ADDR_ID & "'"
                    End Select
                    temp = ExecQuery("Update", "Address record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                End If
            End If
        Next

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Close()
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Return the standardized information as an XML document:
        '   <AddressRec>
        '       <AddrId>   
        '       <JurisId>   
        '       <MatchCode>
        '       <Type>
        '       <Address>        
        '       <City>           
        '       <State>          
        '       <County>         
        '       <Zipcode>         
        '       <Country>         
        '       <Lat>         
        '       <Long>
        '       <Deliverable>
        '   </AddressRec>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("AddressRec")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            ' Modified; Dataflux license expired; 8/1/2017;
            'If debug <> "T" And MATCH_CODE <> "" Then
            If debug <> "T" Then
                AddXMLChild(odoc, resultsRoot, "AddrId", IIf(ADDR_ID = "", " ", HttpUtility.UrlEncode(ADDR_ID)))
                AddXMLChild(odoc, resultsRoot, "JurisId", IIf(JURIS_ID = "", " ", HttpUtility.UrlEncode(JURIS_ID)))
                AddXMLChild(odoc, resultsRoot, "MatchCode", IIf(MATCH_CODE = "", " ", HttpUtility.UrlEncode(MATCH_CODE)))
                AddXMLChild(odoc, resultsRoot, "Type", IIf(ADDR_TYPE = "", " ", ADDR_TYPE))
                AddXMLChild(odoc, resultsRoot, "Address", IIf(ADDR = "", " ", HttpUtility.UrlEncode(ADDR)))
                AddXMLChild(odoc, resultsRoot, "City", IIf(CITY = "", " ", HttpUtility.UrlEncode(CITY)))
                AddXMLChild(odoc, resultsRoot, "State", IIf(STATE = "", " ", HttpUtility.UrlEncode(STATE)))
                AddXMLChild(odoc, resultsRoot, "County", IIf(COUNTY = "", " ", HttpUtility.UrlEncode(COUNTY)))
                AddXMLChild(odoc, resultsRoot, "Zipcode", IIf(ZIPCODE = "", " ", HttpUtility.UrlEncode(ZIPCODE)))
                AddXMLChild(odoc, resultsRoot, "Country", IIf(COUNTRY = "", " ", HttpUtility.UrlEncode(COUNTRY)))
                AddXMLChild(odoc, resultsRoot, "Lat", IIf(LAT = "", " ", HttpUtility.UrlEncode(LAT)))
                AddXMLChild(odoc, resultsRoot, "Long", IIf(LON = "", " ", HttpUtility.UrlEncode(LON)))
                AddXMLChild(odoc, resultsRoot, "Deliverable", IIf(deliverable = "", " ", HttpUtility.UrlEncode(deliverable)))
            Else
                ' Commented out; Dataflux license expired; 8/1/2017;
                'If MATCH_CODE <> "" Then
                '    results = "Success"
                'Else
                '    results = "Failure"
                'End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        ' Commented out; Dataflux license expired; 8/1/2017;
        '' Close Dataflux
        'Try
        '    If Not dfsession Is Nothing Then dfsession.Close()
        '    dfsession = Nothing
        '    config = Nothing
        'Catch ex As Exception
        '    errmsg = errmsg & "Problem closing Dataflux. " & ex.ToString
        'End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("StandardizeAddress : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("StandardizeAddress : Results: " & results & " for '" & ADDR & "' generated matchcode " & MATCH_CODE)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for '" & ADDR & "' generated matchcode " & MATCH_CODE)
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "100"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function
    '<DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    <WebMethod(Description:="Standardizes and geoencodes a provided address usu=ing Melissa DB")> _
    Public Function StandardizeAddress(ByVal sXML As String) As XmlDocument
        ' This function takes an address supplied, either within the supplied XML or in a 
        ' database record, and standardizes it and generates a match code.  Based on 
        ' parameters, it either returns the data generated or updates the database

        ' The input parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <AddressList>
        '       <AddressRec>
        '           <Debug>             - A flag to indicate the service is to run in Debug mode or not
        '                                   "Y"  - Yes for debug mode on.. logging on
        '                                   "N"  - No for debug mode off.. logging off
        '                                   "T"  - Test mode on.. logging off
        '           <Database>          - "C" create S_ADDR_ORG record(s), "U" update record, "V"erification turned off
        '           <AddrId>            - The Id of an existing session, if applicable
        '           <OrgId>             - Organization Id for organizational addresses - optional
        '           <ConId>             - Contact Id for personal addresses - optional
        '           <Type>              - "O"rganization or "P"ersonal address
        '           <GeoCode>           - Geocode address flag ("Y" or "N") - optional
        '           <Address>           - Street address
        '           <City>              - City
        '           <State>             - State
        '           <County>            - County
        '           <Zipcode>           - Zipcode
        '           <Country>           - Country
        '       </AddressRec>
        '   </AddressList>

        ' web.config Parameters used:
        '   siebeldb        - connection string to siebeldb database

        ' Variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim mypath, debug, errmsg, logging, oxml As String

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SADebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim ADDR, CITY, STATE, ZIPCODE, COUNTY, COUNTRY, COUNTRY2, GEOCODEADDR, ADDR_TYPE, ADDR_ID, MATCH_CODE, ADDR_GEOCODE, MD_MAK, MD_MAK_BASE, MDResultCodes As String
        Dim temp, database, lastline, match2, LAT, LON, ORG_ID, CON_ID, JURIS_ID, wp As String
        Dim deliverable, geocodio_key As String
        Dim duplicatedAddrVeri As Boolean


        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        ADDR_ID = ""
        ADDR_TYPE = "O"
        GEOCODEADDR = "N"
        ADDR = "1101 Wilson Suite 1700"
        CITY = "Arlington"
        STATE = "VA"
        COUNTY = ""
        ZIPCODE = ""
        COUNTRY = ""
        COUNTRY2 = ""
        MATCH_CODE = ""
        ORG_ID = ""
        CON_ID = ""
        JURIS_ID = ""
        match2 = ""
        LAT = ""
        LON = ""
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0
        wp = ""
        deliverable = ""
        ADDR_GEOCODE = ""

        ' ============================================
        ' Check parameters
        debug = "N"
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        oxml = sXML
        HttpUtility.UrlDecode(sXML)
        sXML = Regex.Replace(sXML, "&", "&amp;")
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//AddressList/AddressRec")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        debug = UCase(debug)

        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\standardize_address_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\StandardizeAddress.log"
            Try
                log4net.GlobalContext.Properties("SALogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & oxml)
            End If
        End If

        ' Commented out; Dataflux license expired; 8/1/2017;
        '' ============================================
        '' Connect to Dataflux
        'Dim dataflux = System.Configuration.ConfigurationManager.AppSettings("dataflux")
        'If dataflux = "" Then dataflux = "datafluxapp1.hq.local"
        Dim config As New Hashtable
        'config.Add("server", dataflux)
        'config.Add("transport", "TCP")
        'If debug = "Y" Then config.Add("log_file", "C:\Logs\enter.log")

        'Dim dfsession As New DataFlux.dfClient.SessionObject(config)
        'If (dfsession Is Nothing) Then
        '    If debug = "Y" Then mydebuglog.Debug("Unable to open Dataflux")
        '    GoTo CloseOut2
        'Else
        '    If debug = "Y" Then mydebuglog.Debug("  Opening dataflux on " & dataflux)
        'End If

        '=============================================
        'MelissaData Initialization
        Dim addrObj As New mdAddr
        Dim dPath As String = System.Configuration.ConfigurationManager.AppSettings("MD_DataPath")
        Dim dLICENSE As String = System.Configuration.ConfigurationManager.AppSettings("MD_Key")
        'Dim matchupHybObj As New mdHybrid
        'Dim dMUPath As String = System.Configuration.ConfigurationManager.AppSettings("MD_MU_DataPath")
        'Dim dMULicense As String = System.Configuration.ConfigurationManager.AppSettings("MD_MU_Key")
        'Dim parseObj As New mdParse
        'Dim streetObj As New mdStreet
        'Dim zipObj As New mdZip
        'Dim parseFlag As Integer = 0

        If (addrObj Is Nothing) Then
            If debug = "Y" Then mydebuglog.Debug("Unable to open MelissaData")
            GoTo CloseOut2
        Else
            If debug = "Y" Then mydebuglog.Debug("  Opening MelissaData") ' on " & dataflux)
        End If
        '--- end of MelissaData Initialization

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config."
            results = "Failure"
            GoTo CloseOut2
        End Try
        If debug = "Y" Then
            Try
                mydebuglog.Debug(vbCrLf & "Session-")
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Geocoding key 
        geocodio_key = System.Configuration.ConfigurationManager.AppSettings.Get("geocodio_key")
        If geocodio_key = "" Then
            geocodio_key = "4c56e20d4de8d8b5f2aa9a4851145221d9595ed"
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        For i = 0 To oNodeList.Count - 1
            errmsg = ""
            If debug <> "T" Then
                ADDR_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("AddrId", oNodeList.Item(i))))
                ADDR_ID = KeySpace(ADDR_ID)
                ORG_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("OrgId", oNodeList.Item(i))))
                ORG_ID = KeySpace(ORG_ID)
                CON_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("ConId", oNodeList.Item(i))))
                CON_ID = KeySpace(CON_ID)
                ADDR_TYPE = Left(GetNodeValue("Type", oNodeList.Item(i)), 1)
                If ADDR_TYPE = "" And ORG_ID <> "" Then ADDR_TYPE = "O"
                If ADDR_TYPE = "" And CON_ID <> "" Then ADDR_TYPE = "P"
                GEOCODEADDR = Left(GetNodeValue("GeoCode", oNodeList.Item(i)), 1)
                If GEOCODEADDR <> "Y" Then GEOCODEADDR = "N"
                ADDR = Trim(HttpUtility.UrlDecode(GetNodeValue("Address", oNodeList.Item(i))))
                ADDR = CleanString(ADDR)
                ADDR = Left(ADDR, 200)
                CITY = Trim(HttpUtility.UrlDecode(GetNodeValue("City", oNodeList.Item(i))))
                CITY = CleanString(CITY)
                CITY = Left(CITY, 50)
                STATE = Trim(HttpUtility.UrlDecode(GetNodeValue("State", oNodeList.Item(i))))
                STATE = RemoveSymbols(STATE)
                STATE = Left(STATE, 10)
                COUNTY = Trim(HttpUtility.UrlDecode(GetNodeValue("County", oNodeList.Item(i))))
                COUNTY = CleanString(COUNTY)
                COUNTY = Left(COUNTY, 50)
                ZIPCODE = Trim(HttpUtility.UrlDecode(GetNodeValue("Zipcode", oNodeList.Item(i))))
                ZIPCODE = RemoveSymbols(ZIPCODE)
                ZIPCODE = Left(ZIPCODE, 30)
                COUNTRY = Trim(HttpUtility.UrlDecode(GetNodeValue("Country", oNodeList.Item(i))))
                COUNTRY = RemoveSymbols(COUNTRY)
                COUNTRY = Left(COUNTRY, 30)
                database = Trim(Left(GetNodeValue("Database", oNodeList.Item(i)), 1))
            End If
            If debug = "Y" Then
                mydebuglog.Debug("INPUTS------" & vbCrLf & "  ADDR_ID: " & ADDR_ID)
                mydebuglog.Debug("  ORG_ID: " & ORG_ID)
                mydebuglog.Debug("  CON_ID: " & CON_ID)
                mydebuglog.Debug("  ADDR_TYPE: " & ADDR_TYPE)
                mydebuglog.Debug("  GEOCODEADDR: " & GEOCODEADDR)
                mydebuglog.Debug("  ADDR: " & ADDR)
                mydebuglog.Debug("  CITY: " & CITY)
                mydebuglog.Debug("  STATE: " & STATE)
                mydebuglog.Debug("  COUNTY: " & COUNTY)
                mydebuglog.Debug("  ZIPCODE: " & ZIPCODE)
                mydebuglog.Debug("  COUNTRY: " & COUNTRY)
                mydebuglog.Debug("  database: " & database & vbCrLf & "------------")
            End If

            lastline = CITY & ", " & STATE & " " & ZIPCODE

            '***** STEP 1.  Validate Address using MD  *****
            '' ** Determine if this is duplicated Address Verification request **
            'Check siebeldb database to see if the same matchcode has been verified withing certain day(s).
            Dim day_ago As String = System.Configuration.ConfigurationManager.AppSettings.Get("MDVerifiedDaysAgo")
            If MATCH_CODE <> "" Then
                duplicatedAddrVeri = IsAddressVerifiedBefore(cmd, ADDR_TYPE, MATCH_CODE, day_ago)
                'duplicatedAddrVeri = False 'Disable duplicate check for troubleshooting; 10/26/16; Ren Hou;
                ' Verify address if Verification not turned off
            Else
                duplicatedAddrVeri = False
            End If
            If database <> "V" Then
                If Not duplicatedAddrVeri Then
                    'Dim verified As Hashtable
                    Try
                        'Set License
                        addrObj.SetLicenseString(dLICENSE)
                        If Convert.ToDateTime(addrObj.GetLicenseExpirationDate) < Now Then
                            If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File")
                            errmsg = errmsg & "MelissaData Data License Expired: " & addrObj.GetLicenseExpirationDate
                            GoTo CloseOut2
                        End If
                        'Error Checking
                        addrObj.SetPathToUSFiles(dPath)
                        addrObj.SetPathToDPVDataFiles(dPath)
                        addrObj.SetPathToLACSLinkDataFiles(dPath)
                        addrObj.SetPathToCanadaFiles(dPath)
                        addrObj.SetPathToAddrKeyDataFiles(dPath)

                        If (addrObj.InitializeDataFiles() <> 0) Then
                            If debug = "Y" Then mydebuglog.Debug("Unable to Initiate MelissaData Data File")
                            errmsg = errmsg & "Unable to Initiate MelissaData Data File"
                            GoTo CloseOut2
                        Else
                            If debug = "Y" Then mydebuglog.Debug("MelissaData Data File Initialized")
                            'InitErrorString = addrObj.GetInitializeErrorString
                            'DatabaseDate = addrObj.GetDatabaseDate
                            'ExpDate = addrObj.GetExpirationDate
                            'BuildNum = addrObj.GetBuildNumber
                        End If

                        'Set address for verification
                        addrObj.SetAddress(ADDR)
                        addrObj.SetStandardizationType(mdAddr.StandardizeMode.ShortFormat)
                        addrObj.SetLastLine(CITY & ", " & STATE & " " & ZIPCODE)
                        addrObj.SetCountryCode(COUNTRY)

                        Dim Result As Boolean
                        Result = addrObj.VerifyAddress()

                        If Result Then
                            If debug = "Y" Then mydebuglog.Debug("  *Address Verification Finished*")
                        Else
                            If debug = "Y" Then mydebuglog.Debug("  *Address Verification Failed*")
                        End If

                        ' Parse out address fields from last line
                        If Trim(addrObj.GetAddress()) <> "" Then
                            ADDR = addrObj.GetAddress() & _
                                    If(Trim(addrObj.GetAddress2()) <> "", " " & addrObj.GetAddress2(), "") & _
                                    If(Trim(addrObj.GetSuite()) <> "", " " & addrObj.GetSuite(), "")
                            ' Ren Hou; 1/20/17; Added to get Adrress for geodoing where the SUITE number is not relevant and sometime causes problem.
                            ADDR_GEOCODE = addrObj.GetAddress() & _
                                    If(Trim(addrObj.GetAddress2()) <> "", " " & addrObj.GetAddress2(), "")
                        End If
                        If Trim(addrObj.GetZip()) <> "" Then
                            ZIPCODE = addrObj.GetZip() & If(Trim(addrObj.GetPlus4()) <> "", "-" & addrObj.GetPlus4(), "")
                        End If
                        STATE = addrObj.GetState()
                        CITY = addrObj.GetCity()
                        COUNTY = addrObj.GetCountyName()
                        COUNTRY = addrObj.GetCountryCode()
                        lastline = CITY & ", " & STATE & " " & ZIPCODE

                        ' determing dliverability 
                        If (addrObj.GetResults.Contains("AS01") _
                            OrElse addrObj.GetResults.Contains("AS02") _
                            OrElse addrObj.GetResults.Contains("AS03") _
                            OrElse addrObj.GetResults.Contains("AS09") _
                            OrElse addrObj.GetResults.Contains("AS10") _
                            OrElse addrObj.GetResults.Contains("AS13") _
                            OrElse addrObj.GetResults.Contains("AS14") _
                            OrElse addrObj.GetResults.Contains("AS15") _
                            OrElse addrObj.GetResults.Contains("AS16") _
                            OrElse addrObj.GetResults.Contains("AS17") _
                            OrElse addrObj.GetResults.Contains("AS20") _
                            OrElse addrObj.GetResults.Contains("AS23") _
                            OrElse addrObj.GetResults.Contains("AS24") _
                            ) Then
                            deliverable = "VALID"
                        End If

                        If deliverable <> "" Then
                            deliverable = "VALID"
                            ' Get MD_MAK and MD_MAK_BASE
                            MD_MAK = addrObj.GetMelissaAddressKey()
                            MD_MAK_BASE = addrObj.GetMelissaAddressKeyBase()
                            MDResultCodes = addrObj.GetResults()
                        End If
                        If debug = "Y" Then
                            mydebuglog.Debug("   .. Verified addr: " & ADDR)
                            mydebuglog.Debug("   .. Verified postal_code: " & ZIPCODE)
                            mydebuglog.Debug("   .. Verified deliverability: " & deliverable)
                            mydebuglog.Debug(" MS Verification Result: " + GetMDResultDesc(addrObj.GetResults))
                        End If
                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Unable to verify address: " & ex.Message)
                    End Try
                Else
                    If debug = "Y" Then mydebuglog.Debug("  *Address verification skipped*: Address already verified within " + day_ago + " day(s) ago.")
                End If

            Else
                If debug = "Y" Then mydebuglog.Debug("  *Address verification disabled*")
            End If

            ''*** Skip matchcode 
            'MATCH_CODE = ""
            '********* Create Match Code using MelissaData MatchUp Object in SQL Server; Ren Hou; 10/25/2017  ************
            SqlS = "Select siebeldb.dbo.fnGenMdMatchKey_Address_HCI(NULL, @1, @2, @3, @4, '', '', '', '', '', '')"
            cmd.CommandText = SqlS
            cmd.Parameters.Add("@1", Data.SqlDbType.NVarChar, 4000).Value = ADDR
            cmd.Parameters.Add("@2", Data.SqlDbType.NVarChar, 4000).Value = CITY
            cmd.Parameters.Add("@3", Data.SqlDbType.NVarChar, 4000).Value = STATE
            cmd.Parameters.Add("@4", Data.SqlDbType.NVarChar, 4000).Value = ZIPCODE
            Try
                MATCH_CODE = cmd.ExecuteScalar()
            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Unable to generate match code: " & ex.Message)
                myeventlog.Error("  Unable to generate match code: " & ex.Message)
            End Try
            If MATCH_CODE.Contains("Error") Then
                MATCH_CODE = ""
                If debug = "Y" Then mydebuglog.Debug("   .. Error Generating MATCH_CODE from MelissaData ")
                myeventlog.Error("   Error Generating MATCH_CODE from MelissaData: ")
            Else
                If debug = "Y" Then mydebuglog.Debug("   .. Generated MATCH_CODE: " & MATCH_CODE & vbCrLf)
            End If
            ' ***********************************************************************************

            ' -----
            ' Set country code
            If COUNTRY = "US" Then
                COUNTRY = "USA"
                COUNTRY2 = "US"
            End If

            ' Lookup 3-character country code
            If Len(COUNTRY) = 2 Then
                COUNTRY2 = COUNTRY
                SqlS = "SELECT VAL FROM siebeldb.dbo.S_LST_OF_VAL WHERE TYPE='COUNTRY_CODE' AND CODE='" & COUNTRY & "'"
                If debug = "Y" Then mydebuglog.Debug("  Get 3-char country code: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                If Len(temp) = 3 Then COUNTRY = temp
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading country code: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error locating country code. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If debug = "Y" Then mydebuglog.Debug("   .. 3-char country code: " & COUNTRY)
            End If

            ' Lookup 2-character code if Geoencoding
            If COUNTRY2 = "" And GEOCODEADDR = "Y" Then
                SqlS = "SELECT CODE FROM siebeldb.dbo.S_LST_OF_VAL WHERE TYPE='COUNTRY_CODE' AND VAL='" & COUNTRY & "'"
                If debug = "Y" Then mydebuglog.Debug("  Get 2-char country code: " & SqlS & vbCrLf)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                If Len(temp) = 2 Then COUNTRY2 = temp
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading country code: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error locating country code. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If debug = "Y" Then mydebuglog.Debug("   .. 2-char country code: " & COUNTRY2 & vbCrLf)
            End If

            '***** STEP 5.  Geoencode address using 3rd party web service  *****
            ' ^^^^^^^^^^^ Commented out due to changes for Asynchronous geocoding changes ^^^^^^^^
            '' Geoencode address if applicable
            'If GEOCODEADDR = "Y" Then



            ' Using Geocodio
            'If (ADDR_GEOCODE = "") Then ADDR_GEOCODE = ADDR
            'If Trim(ADDR_GEOCODE) <> "" And (LAT = "" Or LON = "") Then
            '    Try
            '        Dim JsonSerial As New JavaScriptSerializer
            '        Dim http As New simplehttp()

            '        ' Prepare URL
            '        Dim SvcURL As String
            '        Dim addrstr As String
            '        SvcURL = System.Configuration.ConfigurationManager.AppSettings("GeocodeUrl")
            '        addrstr = "street=" & Replace(Trim(ADDR_GEOCODE), " ", "+") & "&city=" & Replace(Trim(CITY), " ", "+") & "&state=" & Replace(Trim(STATE), " ", "+") & "&postal_code=" & Replace(Trim(ZIPCODE), " ", "+")
            '        addrstr = addrstr & "&api_key=" & geocodio_key
            '        If debug = "Y" Then mydebuglog.Debug("  Geocode SvcURL: " & SvcURL & addrstr)

            '        ' Generate results
            '        Dim georesults As String
            '        georesults = http.geturl(SvcURL & addrstr, System.Configuration.ConfigurationManager.AppSettings("Geocode_proxyIP"), 80, "", "")
            '        'Dim urlContents As Byte() = Await GetURLContentsAsync(SvcURL)  

            '        If georesults.Length > 0 Then

            '            ' Deserialize
            '            Dim JsonObj As GeocodioObj = JsonSerial.Deserialize(Of GeocodioObj)(georesults)

            '            ' Locate LAT/LON
            '            If JsonObj.results.Length > 0 Then
            '                LAT = JsonObj.results(0).Location.Lat
            '                If debug = "Y" Then mydebuglog.Debug("   .. Geocode LAT: " & LAT.ToString)
            '                LON = JsonObj.results(0).Location.Lng
            '            End If

            '            JsonObj = Nothing
            '        End If

            '    Catch ex As Exception
            '        If debug = "Y" Then mydebuglog.Debug("  Unable to geocode: " & ex.Message)
            '    End Try
            'End If
            'End If
            ' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            ' -----
            ' Get address jurisdiction
            If JURIS_ID = "" Then
                temp = ""
                wp = "<Address><debug>N</debug><regulated></regulated>"
                wp = wp & "<street></street>"
                wp = wp & "<city>" & CITY & "</city>"
                wp = wp & "<state>" & STATE & "</state>"
                wp = wp & "<county>" & COUNTY & "</county>"
                wp = wp & "<zipcode>" & ZIPCODE & "</zipcode>"
                wp = wp & "<country>" & COUNTRY & "</country></Address>"
                JURIS_ID = LoggingService.FindJurisdiction(wp)
                If debug = "Y" Then mydebuglog.Debug("  FindJurisdiction: " & wp)
                If debug = "Y" Then mydebuglog.Debug("   .. JURIS_ID match2: " & JURIS_ID & vbCrLf)
            End If

            ' ============================================
            ' Database operations
            ' Create record
            If database = "C" Then
                If ADDR_ID = "" Then
                    If ADDR_TYPE = "P" And CON_ID = "" Then GoTo UpdateAddr ' Skip if personal and no contact id
                    If ADDR_TYPE = "O" And ORG_ID = "" Then GoTo UpdateAddr ' Skip if organizational and no organization id
GenerateID:
                    ' Generate random address id
                    Select Case ADDR_TYPE
                        Case "P"
                            ADDR_ID = LoggingService.GenerateRecordId("S_ADDR_PER", "N", debug)
                        Case "O"
                            ADDR_ID = LoggingService.GenerateRecordId("S_ADDR_ORG", "N", debug)
                    End Select

                    ' Create address record with new id
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "INSERT INTO siebeldb.dbo.S_ADDR_PER " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM," & _
                            "MODIFICATION_NUM,CONFLICT_ID,DISA_CLEANSE_FLG,PER_ID,ADDR,CITY,COMMENTS," & _
                            "COUNTY,COUNTRY,STATE,ZIPCODE,X_MATCH_CD," & _
                            "X_MATCH_DT,X_LAT,X_LONG,X_JURIS_ID,X_CASS_CHECKED,X_CASS_CODE) " & _
                            "VALUES " & _
                            "('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0," & _
                            "0,0,'N','" & CON_ID & "','" & ADDR & "','" & CITY & "','From StandardizeAddress', '" & _
                            COUNTY & "','" & COUNTRY & "','" & STATE & "', '" & ZIPCODE & "','" & MATCH_CODE & _
                            "',GETDATE(),'" & LAT & "','" & LON & "','" & JURIS_ID & "',GETDATE(),'" & deliverable & "')" & _
                            "; " & _
                            "INSERT INTO siebeldb.dbo.S_ADDR_PER_X " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY," & _
                            "MODIFICATION_NUM,CONFLICT_ID,PAR_ROW_ID,ATTRIB_03,ATTRIB_04, ATTRIB_34)" & _
                            " VALUES " & _
                            "('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1'," & _
                            "0,0,'" & ADDR_ID & "','" & MD_MAK & "','" & MD_MAK_BASE & "', '" & MDResultCodes & "') "

                        Case "O"
                            SqlS = "INSERT INTO siebeldb.dbo.S_ADDR_ORG " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM," & _
                            "MODIFICATION_NUM,CONFLICT_ID,DISA_CLEANSE_FLG,OU_ID,ADDR,CITY,COMMENTS," & _
                            "COUNTY,COUNTRY,STATE,ZIPCODE,X_MATCH_CD," & _
                            "X_MATCH_DT,X_LAT,X_LONG,X_JURIS_ID,X_CASS_CHECKED,X_CASS_CODE) " & _
                            " VALUES " & _
                            "('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0," & _
                            "0,0,'N','" & ORG_ID & "','" & ADDR & "','" & CITY & "','From StandardizeAddress', '" & _
                            COUNTY & "','" & COUNTRY & "','" & STATE & "', '" & ZIPCODE & "','" & MATCH_CODE & _
                            "',GETDATE(),'" & LAT & "','" & LON & "','" & JURIS_ID & "',GETDATE(),'" & deliverable & "')" & _
                            "; " & _
                            "INSERT INTO siebeldb.dbo.S_ADDR_ORG_X " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY," & _
                            "MODIFICATION_NUM,CONFLICT_ID,PAR_ROW_ID,ATTRIB_03,ATTRIB_04, ATTRIB_34)" & _
                            " VALUES " & _
                            "('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1'," & _
                            "0,0,'" & ADDR_ID & "','" & MD_MAK & "','" & MD_MAK_BASE & "', '" & MDResultCodes & "') "
                    End Select
                    temp = ExecQuery("Create", "Address record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                Else
                    errmsg = errmsg & vbCrLf & "Address Id error on creating record. "
                    results = "Failure"
                End If
            End If

            '-----
            ' Update record
UpdateAddr:
            If database = "U" Then
                If ADDR_ID <> "" Then
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "UPDATE siebeldb.dbo.S_ADDR_PER SET LAST_UPD=GETDATE()," & _
                            "ADDR='" & SqlString(ADDR) & "',CITY='" & SqlString(CITY) & "',STATE='" & SqlString(STATE) & "',COUNTRY='" & COUNTRY & "'," & _
                            "ZIPCODE='" & ZIPCODE & "',X_MATCH_CD='" & MATCH_CODE & "',X_MATCH_DT=GETDATE(),X_CASS_CHECKED=GETDATE(),X_CASS_CODE='" & deliverable & "'"
                            If COUNTY <> "" Then SqlS = SqlS & ",COUNTY='" & SqlString(COUNTY) & "'"
                            If LAT <> "" Then SqlS = SqlS & ",X_LAT='" & LAT & "'"
                            If LON <> "" Then SqlS = SqlS & ",X_LONG='" & LON & "'"
                            If JURIS_ID <> "" Then SqlS = SqlS & ",X_JURIS_ID='" & JURIS_ID & "'"
                            SqlS = SqlS & " WHERE ROW_ID='" & ADDR_ID & "'"
                            ' for MAK
                            SqlS = SqlS & _
                            "; MERGE siebeldb.dbo.S_ADDR_PER_X as T" & _
                            " USING (SELECT 'ADDR_ID' as ROW_ID, '" & MD_MAK & "' as MD_MAK, '" & MD_MAK_BASE & "' as MD_MAK_BASE, '" & MDResultCodes & "' as MD_CODES) as S " & _
                            " ON S.ROW_ID = T.PAR_ROW_ID" & _
                            " WHEN MATCHED THEN" & _
                            " UPDATE SET ATTRIB_03 = S.MD_MAK, ATTRIB_04 = S.MD_MAK_BASE" & _
                            " WHEN NOT MATCHED THEN " & _
                            " INSERT(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,CONFLICT_ID,PAR_ROW_ID,ATTRIB_03,ATTRIB_04, ATTRIB_34) VALUES " & _
                            " (S.ROW_ID,GETDATE(),'0-1',GETDATE(),'0-1',0,0,S.ROW_ID,S.MD_MAK,S.MD_MAK_BASE, S.MD_CODES) "

                        Case "O"
                            SqlS = "UPDATE siebeldb.dbo.S_ADDR_ORG SET LAST_UPD=GETDATE()," & _
                            "ADDR='" & SqlString(ADDR) & "',CITY='" & SqlString(CITY) & "',STATE='" & SqlString(STATE) & "',COUNTRY='" & COUNTRY & "'," & _
                            "ZIPCODE='" & ZIPCODE & "',X_MATCH_CD='" & MATCH_CODE & "',X_MATCH_DT=GETDATE(),X_CASS_CHECKED=GETDATE(),X_CASS_CODE='" & deliverable & "'"
                            If COUNTY <> "" Then SqlS = SqlS & ",COUNTY='" & SqlString(COUNTY) & "'"
                            If LAT <> "" Then SqlS = SqlS & ",X_LAT='" & LAT & "'"
                            If LON <> "" Then SqlS = SqlS & ",X_LONG='" & LON & "'"
                            If JURIS_ID <> "" Then SqlS = SqlS & ",X_JURIS_ID='" & JURIS_ID & "'"
                            SqlS = SqlS & " WHERE ROW_ID='" & ADDR_ID & "'"
                            ' for MAK
                            SqlS = SqlS & _
                            "; MERGE siebeldb.dbo.S_ADDR_ORG_X as T" & _
                            " USING (SELECT '" & ADDR_ID & "' as ROW_ID, '" & MD_MAK & "' as MD_MAK, '" & MD_MAK_BASE & "' as MD_MAK_BASE, '" & MDResultCodes & "' as MD_CODES ) as S " & _
                            " ON S.ROW_ID = T.PAR_ROW_ID" & _
                            " WHEN MATCHED THEN" & _
                            " UPDATE SET ATTRIB_03 = S.MD_MAK, ATTRIB_04 = S.MD_MAK_BASE" & _
                            " WHEN NOT MATCHED THEN " & _
                            " INSERT(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,CONFLICT_ID,PAR_ROW_ID,ATTRIB_03,ATTRIB_04, ATTRIB_34) VALUES " & _
                            " (S.ROW_ID,GETDATE(),'0-1',GETDATE(),'0-1',0,0,S.ROW_ID,S.MD_MAK,S.MD_MAK_BASE, S.MD_CODES); "
                    End Select
                    temp = ExecQuery("Update", "Address record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                End If
            End If

            '*** Process geocoding in a new thread (Asynch)
            ' Check to see if this was already done - if applicable
            If ADDR_ID <> "" Then
                Select Case ADDR_TYPE
                    Case "P"
                        SqlS = "SELECT X_LAT, X_LONG " & _
                        "FROM siebeldb.dbo.S_ADDR_PER " & _
                        "WHERE ROW_ID='" & ADDR_ID & "'"
                    Case Else
                        SqlS = "SELECT X_LAT, X_LONG " & _
                        "FROM siebeldb.dbo.S_ADDR_ORG " & _
                        "WHERE ROW_ID='" & ADDR_ID & "'"
                End Select
                If debug = "Y" Then mydebuglog.Debug("  Get existing Latitude: " & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                temp = Str(CheckDBNull(dr(0), enumObjectType.DblType))
                                If Len(temp) > 0 And Val(temp) <> 0 Then LAT = temp
                                temp = Str(CheckDBNull(dr(1), enumObjectType.DblType))
                                If Len(temp) > 0 And Val(temp) <> 0 Then LON = temp
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading latitude: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error locating country code. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If debug = "Y" Then mydebuglog.Debug("   .. Existing latitude: " & LAT & vbCrLf)
            End If

            ' Using Geocodio
            If (ADDR_GEOCODE = "") Then ADDR_GEOCODE = ADDR
            If Trim(ADDR_GEOCODE) <> "" And (LAT = "" Or LON = "") Then
                If GEOCODEADDR = "Y" Then
                    Dim jsonParam As String = ""
                    Dim addrstr As String
                    addrstr = "street=" & Replace(Trim(ADDR_GEOCODE), " ", "+") & "&city=" & Replace(Trim(CITY), " ", "+") & "&state=" & Replace(Trim(STATE), " ", "+") & "&postal_code=" & Replace(Trim(ZIPCODE), " ", "+")
                    addrstr = addrstr & "&api_key=" & geocodio_key
                    jsonParam = "{""addrstr"":""" & addrstr & _
                                """,""addr_type"":""" & ADDR_TYPE & _
                                """,""addrid"":""" & ADDR_ID & _
                                """,""conid"":""" & CON_ID & _
                                """,""orgid"":""" & ORG_ID & _
                                """,""database"":""" & database & """}"
                    Try
                        System.Threading.ThreadPool.QueueUserWorkItem(AddressOf ProcessHCIGeocoding, DirectCast(jsonParam, Object))
                    Catch ex As Exception
                        errmsg = "Failed to call Asynchronous geocoding function (ProcessHCIGeocoding); ErrrMsg: " & ex.Message
                        mydebuglog.Debug(errmsg & "DT: " & Now().ToString)
                        myeventlog.Error(errmsg & "DT: " & Now().ToString)
                    End Try
                End If
            End If
        Next

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Close()
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf
        End Try

CloseOut2:
        ' ============================================
        ' Return the standardized information as an XML document:
        '   <AddressRec>
        '       <AddrId>   
        '       <JurisId>   
        '       <MatchCode>
        '       <Type>
        '       <Address>        
        '       <City>           
        '       <State>          
        '       <County>         
        '       <Zipcode>         
        '       <Country>         
        '       <Lat>         
        '       <Long>
        '       <Deliverable>
        '   </AddressRec>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("AddressRec")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            'If debug <> "T" And MATCH_CODE <> "" Then
            If debug <> "T" Then
                AddXMLChild(odoc, resultsRoot, "AddrId", If(ADDR_ID = "", " ", HttpUtility.UrlEncode(ADDR_ID)))
                AddXMLChild(odoc, resultsRoot, "JurisId", If(JURIS_ID = "", " ", HttpUtility.UrlEncode(JURIS_ID)))
                AddXMLChild(odoc, resultsRoot, "MatchCode", If(MATCH_CODE = "", " ", HttpUtility.UrlEncode(MATCH_CODE)))
                AddXMLChild(odoc, resultsRoot, "Type", If(ADDR_TYPE = "", " ", ADDR_TYPE))
                AddXMLChild(odoc, resultsRoot, "Address", If(ADDR = "", " ", HttpUtility.UrlEncode(ADDR)))
                AddXMLChild(odoc, resultsRoot, "City", If(CITY = "", " ", HttpUtility.UrlEncode(CITY)))
                AddXMLChild(odoc, resultsRoot, "State", If(STATE = "", " ", HttpUtility.UrlEncode(STATE)))
                AddXMLChild(odoc, resultsRoot, "County", If(COUNTY = "", " ", HttpUtility.UrlEncode(COUNTY)))
                AddXMLChild(odoc, resultsRoot, "Zipcode", If(ZIPCODE = "", " ", HttpUtility.UrlEncode(ZIPCODE)))
                AddXMLChild(odoc, resultsRoot, "Country", If(COUNTRY = "", " ", HttpUtility.UrlEncode(COUNTRY)))
                AddXMLChild(odoc, resultsRoot, "Lat", If(LAT = "", " ", HttpUtility.UrlEncode(LAT)))
                AddXMLChild(odoc, resultsRoot, "Long", If(LON = "", " ", HttpUtility.UrlEncode(LON)))
                AddXMLChild(odoc, resultsRoot, "Deliverable", If(deliverable = "", " ", HttpUtility.UrlEncode(deliverable)))
                AddXMLChild(odoc, resultsRoot, "MelissaMAK", If(MD_MAK = "", " ", HttpUtility.UrlEncode(MD_MAK))) 'MelissaData MAK key
                AddXMLChild(odoc, resultsRoot, "MelissaMAKBase", If(MD_MAK_BASE = "", " ", HttpUtility.UrlEncode(MD_MAK_BASE))) 'MelissaData MAK Base key
                AddXMLChild(odoc, resultsRoot, "MDResultCodes", If(MDResultCodes = "", " ", HttpUtility.UrlEncode(MDResultCodes))) 'MD Result Codes
                'addrObj.GetResults
                'AddXMLChild(odoc, resultsRoot, "MDAddrVerified", If(duplicatedAddrVeri, "Duplicate", "Verified"))
            Else
                If MATCH_CODE <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

            'Close MelissaData
            Try
                If Not addrObj Is Nothing Then addrObj.Dispose()
                config = Nothing
            Catch ex As Exception
                errmsg = errmsg & "Problem closing MelissaData object. " & ex.ToString
            End Try

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

        If debug = "Y" Then mydebuglog.Debug("Result XML: " & odoc.OuterXml)

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("StandardizeAddress : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("StandardizeAddress : Results: " & results & " for '" & ADDR & "'; " & If(duplicatedAddrVeri, "", "Address Verified (MD)") & "; generated matchcode: " & MATCH_CODE)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for '" & ADDR & "' generated matchcode " & MATCH_CODE & If(duplicatedAddrVeri, "", "; Address Verified by MD"))
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "100"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function

    Public Sub ProcessHCIGeocoding(ByVal jsonParam As Object)
        '**********************************
        '1. Call Google geocoding Web Service.
        '2, Update resulted LAT, Lon into address table
        '**********************************
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SADebugLog")
        Dim errmsg As String, results As String
        Dim addrstr As String, addr_type As String, addrid As String, conid As String, orgid As String, database As String
        Dim JsonSerial As New JavaScriptSerializer
        Dim con As SqlConnection = New SqlConnection(), cmd As SqlCommand = New SqlCommand
        Dim sqlStr As String
        '*********  Main Try block  *******
        Try
            ' ============================================
            ' Open log file if applicable
            Try
                log4net.GlobalContext.Properties("SALogFileName") = "C:\Logs\StandardizeAddress.log"
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = " Error Opening Log."
                Throw New Exception(errmsg)
            End Try

            mydebuglog.Debug("----------------------------------")
            mydebuglog.Debug("Running ProcessHCIGeocoding: " & Now.ToString & vbCrLf)

            'Get parameters
            Try
                Dim dict As Dictionary(Of String, String) = JsonSerial.Deserialize(Of Dictionary(Of String, String))(jsonParam)
                addrstr = dict("addrstr")
                addr_type = dict("addr_type")
                addrid = dict("addrid")
                conid = dict("conid")
                orgid = dict("orgid")
                database = dict("database")
            Catch ex As Exception
                errmsg = " Unable to parse Json parameters: " & ex.Message
                Throw New Exception(errmsg)
            End Try

            mydebuglog.Debug("Parameters: ")
            mydebuglog.Debug("   ..addr_type: " & addr_type)
            mydebuglog.Debug("   ..addrid: " & addrid)
            mydebuglog.Debug("   ..conid: " & conid)
            mydebuglog.Debug("   ..orgid: " & orgid)
            mydebuglog.Debug("   ..database: " & database)

            'Check if addrid exits in database
            If addrid <> "" Then
                'Create connections
                con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
                con.Open()
                cmd.Connection = con

                ' Check address table for this addrid
                sqlStr = "select count(1) from siebeldb.dbo."
                ' Generate random address id
                Select Case addr_type
                    Case "P"
                        sqlStr = sqlStr & "S_ADDR_PER"
                    Case "O"
                        sqlStr = sqlStr & "S_ADDR_ORG"
                End Select
                sqlStr = sqlStr & " where ROW_ID = '" & addrid & "'"
                mydebuglog.Debug("  Check existing address ID. Query: " & sqlStr)
                cmd.CommandText = sqlStr
                Dim cnt As Integer = cmd.ExecuteScalar()
                Dim retryLimit As Integer = 3
                While (cnt = 0 And retryLimit > 0)
                    If cnt = 0 Then Sleep(200) 'Pause
                    cnt = cmd.ExecuteScalar()
                    retryLimit = retryLimit - 1
                End While
                If cnt = 0 Then
                    errmsg = " No Addr record exits: no LAT LON values are updated."
                    'mydebuglog.Debug(errmsg)
                    'myeventlog.Info(errmsg)
                    Throw New Exception(errmsg)
                End If
                ' Using Geocodio
                Dim ADDR_GEOCODE As String = "", LAT As String = "", LON As String = ""
                Try
                    Dim http As New simplehttp()
                    ' Prepare URL
                    Dim SvcURL As String
                    SvcURL = System.Configuration.ConfigurationManager.AppSettings("GeocodeUrl")
                    mydebuglog.Debug("  Geocode SvcURL: " & SvcURL & addrstr)

                    ' Generate results
                    Dim georesults As String
                    'georesults = http.geturl(SvcURL & addrstr, System.Configuration.ConfigurationManager.AppSettings("Geocode_proxyIP"), 80, "", "")
                    georesults = http.geturl(SvcURL & Replace(addrstr, "#", ""), System.Configuration.ConfigurationManager.AppSettings("Geocode_proxyIP"), 80, "", "")
                    mydebuglog.Debug("   .. Geocode results: " & georesults)
                    'Dim urlContents As Byte() = Await GetURLContentsAsync(SvcURL)  

                    If georesults.Length > 0 And Not georesults.Contains("returned an er") Then
                        ' Deserialize
                        Dim JsonObj As GeocodioObj = JsonSerial.Deserialize(Of GeocodioObj)(georesults)
                        ' Locate LAT/LON
                        If JsonObj.results.Length > 0 Then
                            LAT = JsonObj.results(0).Location.Lat
                            LON = JsonObj.results(0).Location.Lng
                            mydebuglog.Debug("   .. Geocode LAT: " & LAT.ToString & " LON: " & LON.ToString)
                        End If
                        JsonObj = Nothing
                    Else
                        errmsg = "  Call to https://api.geocod.io/v1/geocode returns error; ErrMsg: " & georesults
                        Throw New Exception(errmsg)
                    End If
                Catch ex As Exception
                    errmsg = "  Unable to geocode: " & ex.Message
                    Throw New Exception(errmsg)
                End Try

                'Update Addres tables
                Select Case addr_type
                    Case "P"
                        sqlStr = "UPDATE siebeldb.dbo.S_ADDR_PER SET LAST_UPD=GETDATE() "
                        If LAT <> "" Then sqlStr = sqlStr & ",X_LAT='" & LAT & "'"
                        If LON <> "" Then sqlStr = sqlStr & ",X_LONG='" & LON & "'"
                        sqlStr = sqlStr & " WHERE ROW_ID='" & addrid & "'"
                    Case "O"
                        sqlStr = "UPDATE siebeldb.dbo.S_ADDR_ORG SET LAST_UPD=GETDATE() "
                        If LAT <> "" Then sqlStr = sqlStr & ",X_LAT='" & LAT & "'"
                        If LON <> "" Then sqlStr = sqlStr & ",X_LONG='" & LON & "'"
                        sqlStr = sqlStr & " WHERE ROW_ID='" & addrid & "'"
                End Select
                cmd.CommandText = sqlStr
                Try
                    mydebuglog.Debug("    Updating Address Table: " & sqlStr)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errmsg = "Failed to update Addres table; ADDR_ID: " & addrid &
                                    "; ADDR_TYPE: " & addr_type & " ErrMsg: " & ex.Message
                    Throw New Exception(errmsg)
                End Try
                myeventlog.Info("Finished google geocoding and LAT, LON updates; ADDR_ID: " & addrid & _
                                "; ADDR_TYPE: " & addr_type)
                mydebuglog.Debug("--- End --- " & Now.ToString())
                mydebuglog.Debug("------------------------------------")
                con.Close()
                con.Dispose()
            Else
                errmsg = "StandardizeAddress-ProcessHCIGeocoding: Skiped geocoding: there is no address id. "
                mydebuglog.Debug(errmsg)
                myeventlog.Info(errmsg)
                mydebuglog.Debug("--- End --- " & Now.ToString())
                mydebuglog.Debug("------------------------------------")
                'Throw New Exception(errmsg)
            End If
        Catch ex As Exception
            mydebuglog.Debug(ex.Message & "Timestamp: ")
            myeventlog.Info("StandardizeAddress-ProcessHCIGeocoding: Failure; " & ex.Message & "Timestamp: " & Now.ToString())
            mydebuglog.Debug("--- End --- " & Now.ToString())
            mydebuglog.Debug("------------------------------------")
            con.Close()
            con.Dispose()
        End Try

    End Sub
    <WebMethod(Description:="Geoencode the supplied address")> _
    Public Function GeoencodeAddress(ByVal Addr As String, ByVal Debug As String) As XmlDocument

        ' Generic variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim mypath, errmsg, logging As String

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("GADebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service
        Dim BasicService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim S_LAT, S_LON As String
        Dim bingkey, geocodio_key As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        S_LON = "0"
        S_LAT = "0"

        ' ============================================
        ' Check parameters
        Debug = UCase(Debug)
        If Debug = "" Then Debug = "N"
        If Debug <> "N" And Debug <> "Y" And Debug <> "T" Then Debug = "N"
        If Debug = "T" Then
            Addr = "1400 Key Blvd, Ste 700, Arlington, VA 22209 USA"
        End If
        If Addr = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Get defaults
        bingkey = System.Configuration.ConfigurationManager.AppSettings.Get("bing_key")
        If bingkey = "" Then
            bingkey = ""
        End If
        geocodio_key = System.Configuration.ConfigurationManager.AppSettings.Get("geocodio_key")
        If geocodio_key = "" Then
            geocodio_key = ""
        End If

        ' ============================================
        ' Open log file if applicable
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            logfile = "C:\Logs\GeoencodeAddress.log"
            Try
                log4net.GlobalContext.Properties("GALogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut
            End Try

            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & Debug)
                mydebuglog.Debug("  Address:" & Addr)
                mydebuglog.Debug("  bingkey:" & bingkey)
                mydebuglog.Debug("  geocodio_key:" & geocodio_key & vbCrLf)
            End If
        End If

        ' ============================================
        ' Execute function
        Try
            'Dim Geocoder As New GeocodeRequest()

            ' Set the credentials using a valid Bing Maps Key
            'Dim GeoCredentials As New Credentials()
            'Geocoder.Credentials = New Credentials() With {.ApplicationId = bingkey}

            ' Set address to query
            'Geocoder.Query = Addr

            ' Set the options to only return high confidence results
            'Dim filters() As ConfidenceFilter = {New ConfidenceFilter() With {.MinimumConfidence = Confidence.Medium}}

            'Dim geocodeOptions As New GeocodeOptions() _
            'With {.Filters = filters}

            'Geocoder.Options = GeocodeOptions

            ' Make the geocode request
            'Dim GeocodeService As New GeocodeServiceClient("BasicHttpBinding_IGeocodeService")
            'Dim geocodeResponse = GeocodeService.Geocode(Geocoder)

            ' Use the results in your application.
            'If geocodeResponse.Results.Length <> 0 Then
            'results = GeocodeResponse.Results(0).DisplayName
            'If results.Length > 0 Then
            'S_LAT = GeocodeResponse.Results(0).Locations(0).Latitude
            'S_LON = GeocodeResponse.Results(0).Locations(0).Longitude
            'End If
            'End If
            'Create REST Services geocode request using Locations API

            ' Remove objects created
            'Geocoder = Nothing
            'GeoCredentials = Nothing
            'GeocodeService = Nothing

            ' Bing
            'Dim geocodeRequest As String = "http://dev.virtualearth.net/REST/v1/Locations?q=" & Addr.Replace(" ", "+") & "&o=xml&key=" & bingkey

            'Make the request and get the response
            'Dim geocodeResponse As XmlDocument = GetXmlResponse(geocodeRequest)

            'Create namespace manager
            'Dim nsmgr As New XmlNamespaceManager(geocodeResponse.NameTable)
            'nsmgr.AddNamespace("rest", "http://schemas.microsoft.com/search/local/ws/rest/v1")

            'Get all locations in the response and then extract the coordinates for the top location
            'Dim locationElements As XmlNodeList = geocodeResponse.SelectNodes("//rest:Location", nsmgr)
            'If locationElements.Count = 0 Then
            'Addr = Right(Addr, Len(Addr) - InStr(Addr, ","))
            'GeocodeRequest = "http://dev.virtualearth.net/REST/v1/Locations?q=" & Addr.Replace(" ", "+") & "&o=xml&key=" & bingkey
            'GeocodeResponse = GetXmlResponse(GeocodeRequest)
            'locationElements = GeocodeResponse.SelectNodes("//rest:Location", nsmgr)
            'End If

            'If locationElements.Count = 0 Then
            'errmsg = errmsg & vbCrLf & "The location you entered could not be geocoded."
            'S_LON = "0"
            'S_LAT = "0"
            'Else
            'Dim displayGeocodePoints As XmlNodeList = locationElements(0).SelectNodes(".//rest:GeocodePoint/rest:UsageType[.='Display']/parent::node()", nsmgr)
            'S_LAT = displayGeocodePoints(0).SelectSingleNode(".//rest:Latitude", nsmgr).InnerText
            'S_LON = displayGeocodePoints(0).SelectSingleNode(".//rest:Longitude", nsmgr).InnerText
            'If Debug = "Y" Then mydebuglog.Debug("  Geocoding Results: " & S_LAT & ", " & S_LON & vbCrLf)
            'displayGeocodePoints = Nothing
            'End If

            ' Remove objects created
            'GeocodeResponse = Nothing
            'locationElements = Nothing

            ' Geocodeio
            Dim JsonSerial As New JavaScriptSerializer
            Dim http As New simplehttp()

            ' Prepare Address
            Addr = Addr.Replace(" ", "+")
            Addr = Addr.Replace(",", "")
            Addr = Addr.Replace("#", "")
            Addr = Addr.Replace("%", "")
            Addr = Addr.Replace(",", "")

            ' Prepare URL
            Dim SvcURL As String
            Dim addrstr As String
            SvcURL = System.Configuration.ConfigurationManager.AppSettings("GeocodeUrl")
            addrstr = "q=" & Addr.Replace(" ", "+")
            addrstr = addrstr & "&api_key=" & geocodio_key
            If Debug = "Y" Then mydebuglog.Debug("  Geocode SvcURL: " & SvcURL & addrstr)

            ' Generate results
            Dim georesults As String
            'georesults = http.geturl(SvcURL & addrstr, System.Configuration.ConfigurationManager.AppSettings("Geocode_proxyIP"), 80, "", "")
            georesults = http.geturl(SvcURL & Replace(addrstr, "#", ""), System.Configuration.ConfigurationManager.AppSettings("Geocode_proxyIP"), 80, "", "")
            If results.Length > 0 Then

                ' Deserialize
                Dim JsonObj As GeocodioObj = JsonSerial.Deserialize(Of GeocodioObj)(georesults)

                ' Locate LAT/LON
                If JsonObj.results.Length > 0 Then
                    S_LAT = JsonObj.results(0).Location.Lat
                    If Debug = "Y" Then mydebuglog.Debug("   .. Geocode LAT: " & S_LAT.ToString)
                    S_LON = JsonObj.results(0).Location.Lng
                End If

                JsonObj = Nothing
            End If
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("  Unable to geocode: " & ex.Message)
        End Try

CloseOut:
        ' ============================================
        ' Return the standardized information as an XML document:
        '   <Location>
        '       <Lat>   
        '       <Long>   
        '   </Location>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("Location")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            If Debug <> "T" And S_LAT <> "" Then
                AddXMLChild(odoc, resultsRoot, "Lat", If(S_LAT = "", " ", HttpUtility.UrlEncode(S_LAT)))
                AddXMLChild(odoc, resultsRoot, "Long", If(S_LON = "", " ", HttpUtility.UrlEncode(S_LON)))
            Else
                If S_LAT <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("GeoencodeAddress : Error: " & Trim(errmsg))
        If Debug <> "T" Then myeventlog.Info("GeoencodeAddress : Results: " & results & " for '" & Addr & "' generated latitude " & S_LAT & " and longitude " & S_LON)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for '" & Addr & "' generated latitude " & S_LAT & " and longitude " & S_LON)
                If Debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Log Performance Data
        Dim VersionNum As String = "100"
        If Debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function

    Private Function GetXmlResponse(ByVal requestUrl As String) As XmlDocument
        Dim request As HttpWebRequest = TryCast(WebRequest.Create(requestUrl), HttpWebRequest)
        Using response As HttpWebResponse = TryCast(request.GetResponse(), HttpWebResponse)
            If response.StatusCode <> HttpStatusCode.OK Then
                Throw New Exception(String.Format("Server error (HTTP {0}: {1}).", response.StatusCode, response.StatusDescription))
            End If
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(response.GetResponseStream())
            Return xmlDoc
        End Using
    End Function

    <WebMethod(Description:="Clean a contact record on request")> _
    Public Function CleanContact(ByVal sXML As String) As XmlDocument
        ' This function attempts to "clean" a contact supplied by first standardizing it and
        ' generating match codes, and then locating matching records. If a match is
        ' found it returns the the matching record and it's id, otherwise it returns.  It will 
        ' also optionally update the matching record with changes from the supplied contact.

        ' The input parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <Contacts>
        '       <Contact>
        '           <Debug></Debug>         - A flag to indicate the service is to run in Debug mode or not
        '                                       "Y"  - Yes for debug mode on.. logging on
        '                                       "N"  - No for debug mode off.. logging off
        '                                       "T"  - Test mode on.. logging off
        '           <Database></Database>   - "C" create S_CONTACT record(s), 
        '                                       "U" update record, other do nothing
        '           <ConId></ConId>         - The Id of an existing contact, if applicable
        '           <PartId></PartId>      - Participant Id
        '           <FirstName></FirstName> - First name of contact
        '           <MidName></MidName>     - Middle name of contact
        '           <LastName></LastName>   - Last name of contact
        '           <Gender></Gender>       - Gender of contact
        '           <FullName></Fullname>   - Full name of contact
        '           <DOB></DOB>             - Date of Birth
        '           <WorkPhone></WorkPhone> - Work phone
        '           <SSN></SSN>             - Social Security Number
        '           <EmailAddr></EmailAddr> - Email address
        '           <AddrId></AddrId>       - The Id of an existing address, if applicable
        '           <AddrMatch></AddrMatch> - Address match code
        '           <AddrType></AddrType>   - Address type code ("P"ersonal, "B"usiness)
        '           <OrgId></OrgId>         - The Id of an existing organization, if applicable
        '           <OrgMatch></OrgMatch>   - Organization match code
        '           <OrgPhone></OrgPhone>   - Organization main phone number
        '           <RegNum></RegNum>       - Web registration id
        '           <SubConId></SubConId>   - Subscription contact record
        '           <TrainerNo></TrainerNo> - Trainer number 
        '           <HomePhone></HomePhone> - Home phone number
        '           <JobTitle></JobTitle>   - Job Title
        '           <Source></Source>       - Source of contact (matches to a contact category)
        '           <PerTitle></PerTitle>   - Personal Title
        '           <Confidence></Confidence>-Confidence - number of partial matches required
        '       </Contact>
        '   </Contacts>

        ' web.config Parameters used:
        '   hcidb        - connection string to siebeldb database

        ' Generic variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim rDoc As XmlDocument
        Dim rNodeList As XmlNodeList
        Dim i, j As Integer
        Dim mypath, debug, errmsg, warnmsg, logging, wp As String

        ' Database declarations
        Dim con As SqlConnection, con_ro As SqlConnection
        Dim cmd As SqlCommand, cmd_ro As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String, ConnS_ro As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("CCDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service
        Dim BasicService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim FST_NAME, MID_NAME, LAST_NAME, GENDER, MATCH_CODE, FULL_NAME, CON_ID As String
        Dim DOB, WORK_PHONE, SSN, EMAIL_ADDR, ADDR_MATCH, ORG_MATCH, ORG_PHONE As String
        Dim REG_NUM, SUB_CON_ID, TRAINER_NO, HOME_PHONE, PART_ID, ADDR_TYPE As String
        Dim ORG_ID, ADDR_ID, JOB_TITLE, SOURCE, SOURCE_ID, CON_CHRCTR_ID As String
        Dim S_FST_NAME, S_MID_NAME, S_LAST_NAME, S_GENDER, S_MATCH_CODE, S_FULL_NAME, S_CON_ID, pLAST_NAME, pFST_NAME As String
        Dim S_DOB, S_WORK_PHONE, S_SSN, S_EMAIL_ADDR, S_ADDR_MATCH, S_ORG_MATCH, S_ORG_PHONE As String
        Dim S_REG_NUM, S_SUB_CON_ID, S_TRAINER_NO, S_HOME_PHONE, S_PART_ID, S_ADDR_TYPE, O_PART_ID As String
        Dim S_ORG_ID, S_ADDR_ID, S_JOB_TITLE, PER_TITLE, S_PER_TITLE, S_PER_ADDR_ID As String
        Dim temp, database As String
        Dim Confidence, match_count, high_count As Integer
        Dim MatchNew_count As Integer
        Dim tFST_NAME, tMID_NAME, tLAST_NAME, tGENDER, tMATCH_CODE, tFULL_NAME, tCON_ID As String
        Dim tDOB, tWORK_PHONE, tSSN, tEMAIL_ADDR, tADDR_MATCH, tORG_MATCH, tORG_PHONE As String
        Dim tREG_NUM, tSUB_CON_ID, tTRAINER_NO, tHOME_PHONE, tPART_ID, tADDR_TYPE, tPADDR_MATCH, tOADDR_MATCH As String
        Dim tORG_ID, tADDR_ID, tJOB_TITLE, tSOURCE, tSOURCE_ID, tCON_CHRCTR_ID, tPER_ADDR_ID, tPER_TITLE As String

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        warnmsg = ""
        results = "Success"
        FST_NAME = "Christopher"
        MID_NAME = "L"
        LAST_NAME = "Bobbitt"
        pLAST_NAME = ""
        pFST_NAME = ""
        GENDER = ""
        MATCH_CODE = ""
        FULL_NAME = ""
        CON_ID = ""
        PART_ID = ""
        O_PART_ID = ""
        DOB = ""
        WORK_PHONE = "703-524-1200"
        SSN = ""
        EMAIL_ADDR = "bobbittc@gettips.com"
        ADDR_ID = ""
        ADDR_MATCH = ""
        ADDR_TYPE = ""
        ORG_ID = ""
        ORG_MATCH = ""
        ORG_PHONE = ""
        REG_NUM = ""
        SUB_CON_ID = ""
        TRAINER_NO = ""
        HOME_PHONE = ""
        JOB_TITLE = ""
        SOURCE = ""
        SOURCE_ID = ""
        CON_CHRCTR_ID = ""
        PER_TITLE = "Mr."
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0
        Confidence = 5
        match_count = 0
        S_FST_NAME = ""
        S_MID_NAME = ""
        S_LAST_NAME = ""
        S_GENDER = ""
        S_MATCH_CODE = ""
        S_FULL_NAME = ""
        S_CON_ID = ""
        S_DOB = ""
        S_WORK_PHONE = ""
        S_SSN = ""
        S_EMAIL_ADDR = ""
        S_ADDR_MATCH = ""
        S_ORG_MATCH = ""
        S_ORG_PHONE = ""
        S_REG_NUM = ""
        S_SUB_CON_ID = ""
        S_TRAINER_NO = ""
        S_HOME_PHONE = ""
        S_PART_ID = ""
        S_ADDR_TYPE = ""
        S_ORG_ID = ""
        S_ADDR_ID = ""
        S_PER_ADDR_ID = ""
        S_JOB_TITLE = ""
        S_PER_TITLE = ""
        tFST_NAME = ""
        tMID_NAME = ""
        tLAST_NAME = ""
        tGENDER = ""
        tMATCH_CODE = ""
        tFULL_NAME = ""
        tCON_ID = ""
        tDOB = ""
        tWORK_PHONE = ""
        tSSN = ""
        tEMAIL_ADDR = ""
        tADDR_MATCH = ""
        tORG_MATCH = ""
        tORG_PHONE = ""
        tREG_NUM = ""
        tSUB_CON_ID = ""
        tTRAINER_NO = ""
        tHOME_PHONE = ""
        tPART_ID = ""
        tADDR_TYPE = ""
        tORG_ID = ""
        tADDR_ID = ""
        tJOB_TITLE = ""
        tSOURCE = ""
        tSOURCE_ID = ""
        tCON_CHRCTR_ID = ""
        tPER_ADDR_ID = ""
        tPER_TITLE = ""
        tPADDR_MATCH = ""
        tOADDR_MATCH = ""
        ConnS = ""
        ConnS_ro = ""
        high_count = 0
        MatchNew_count = 0

        ' ============================================
        ' Check parameters
        debug = "N"
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        HttpUtility.UrlDecode(sXML)
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//Contacts/Contact")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        debug = UCase(debug)

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
            ConnS_ro = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb_ro").ConnectionString
            If ConnS_ro = "" Then ConnS_ro = "server=;"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("CleanContact_debug").ToUpper()
            If temp = "Y" And debug <> "T" Then debug = temp
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\clean_contact_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening XML Log. "
                GoTo CloseOut2
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\CleanContact.log"
            Try
                log4net.GlobalContext.Properties("CCLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        ' ============================================; 2020-05-18; Ren Hou; Added for read-only per Chris;
        ' Open read-only database connection 
        errmsg = OpenDBConnection(ConnS_ro, con_ro, cmd_ro)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        For i = 0 To oNodeList.Count - 1
            errmsg = ""
            ' ============================================
            ' Extract data from parameter string
            If debug <> "T" Then
                Dim regExp As String = "[^-_ ',A-Za-z\u00c0-\u00ff]|[\u00fe\u00de]"
                FST_NAME = Trim(RemovePluses(HttpUtility.UrlDecode(Trim(GetNodeValue("FirstName", oNodeList.Item(i))))))
                FST_NAME = Regex.Replace(FST_NAME, regExp, "") ' 4/2/2020; Ren Hou; remove special characters per Chris;
                MID_NAME = Trim(RemovePluses(HttpUtility.UrlDecode(Trim(GetNodeValue("MidName", oNodeList.Item(i))))))
                MID_NAME = Regex.Replace(MID_NAME, regExp, "") ' 4/2/2020; Ren Hou; remove special characters per Chris;
                LAST_NAME = Trim(RemovePluses(HttpUtility.UrlDecode(Trim(GetNodeValue("LastName", oNodeList.Item(i))))))
                LAST_NAME = Regex.Replace(LAST_NAME, regExp, "") ' 4/2/2020; Ren Hou; remove special characters per Chris;
                GENDER = Trim(RemovePluses(GetNodeValue("Gender", oNodeList.Item(i))))
                FULL_NAME = Trim(RemovePluses(HttpUtility.UrlDecode(GetNodeValue("FullName", oNodeList.Item(i)))))
                FULL_NAME = Regex.Replace(FULL_NAME, regExp, "") ' 4/2/2020; Ren Hou; remove special characters per Chris;
                If FULL_NAME = "" Then
                    If MID_NAME = "" Then
                        FULL_NAME = FST_NAME & " " & LAST_NAME
                    Else
                        FULL_NAME = FST_NAME & " " & MID_NAME & " " & LAST_NAME
                    End If
                End If
                CON_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("ConId", oNodeList.Item(i))))
                CON_ID = KeySpace(CON_ID)
                database = Left(GetNodeValue("Database", oNodeList.Item(i)), 1)
                PART_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("PartId", oNodeList.Item(i))))
                PART_ID = KeySpace(PART_ID)
                O_PART_ID = PART_ID
                DOB = Trim(HttpUtility.UrlDecode(GetNodeValue("DOB", oNodeList.Item(i))))
                DOB = StndDate(DOB)
                Try
                    WORK_PHONE = RemovePluses(HttpUtility.UrlDecode(GetNodeValue("WorkPhone", oNodeList.Item(i))))
                    If WORK_PHONE <> "" Then WORK_PHONE = StndPhone(WORK_PHONE)
                Catch ex As Exception
                End Try
                SSN = Trim(HttpUtility.UrlDecode(GetNodeValue("SSN", oNodeList.Item(i))))
                SSN = StndSSN(SSN)
                EMAIL_ADDR = Trim(HttpUtility.UrlDecode(GetNodeValue("EmailAddr", oNodeList.Item(i))))
                ADDR_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("AddrId", oNodeList.Item(i))))
                ADDR_ID = KeySpace(ADDR_ID)
                ADDR_MATCH = Trim(HttpUtility.UrlDecode(GetNodeValue("AddrMatch", oNodeList.Item(i))))
                ADDR_TYPE = Left(GetNodeValue("AddrType", oNodeList.Item(i)), 1)
                If ADDR_TYPE = "O" Then ADDR_TYPE = "B"
                ORG_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("OrgId", oNodeList.Item(i))))
                ORG_ID = KeySpace(ORG_ID)
                ORG_MATCH = Trim(HttpUtility.UrlDecode(GetNodeValue("OrgMatch", oNodeList.Item(i))))
                Try
                    ORG_PHONE = RemovePluses(HttpUtility.UrlDecode(GetNodeValue("OrgPhone", oNodeList.Item(i))))
                    If ORG_PHONE <> "" Then ORG_PHONE = StndPhone(ORG_PHONE)
                Catch ex As Exception
                End Try
                REG_NUM = HttpUtility.UrlDecode(GetNodeValue("RegNum", oNodeList.Item(i)))
                SUB_CON_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("SubConId", oNodeList.Item(i))))
                TRAINER_NO = Trim(HttpUtility.UrlDecode(GetNodeValue("TrainerNo", oNodeList.Item(i)))) '6/27/21; Ren Hou; Added to trim Teainer_No;
                Try
                    HOME_PHONE = Trim(RemovePluses(HttpUtility.UrlDecode(GetNodeValue("HomePhone", oNodeList.Item(i)))))
                    If HOME_PHONE <> "" Then HOME_PHONE = StndPhone(HOME_PHONE)
                Catch ex As Exception
                End Try
                JOB_TITLE = Trim(RemovePluses(HttpUtility.UrlDecode(GetNodeValue("JobTitle", oNodeList.Item(i)))))
                SOURCE = Trim(RemovePluses(HttpUtility.UrlDecode(GetNodeValue("Source", oNodeList.Item(i)))))
                PER_TITLE = Trim(RemovePluses(HttpUtility.UrlDecode(GetNodeValue("PerTitle", oNodeList.Item(i)))))
                If GENDER <> "" And PER_TITLE = "" Then
                    If GENDER = "M" Then PER_TITLE = "Mr."
                    If GENDER = "F" Then PER_TITLE = "Ms."
                End If
                temp = Trim(GetNodeValue("Confidence", oNodeList.Item(i)))
                If temp <> "" And IsNumeric(temp) Then Confidence = temp
            End If
            If debug = "Y" Then
                mydebuglog.Debug("  database: " & database)
                mydebuglog.Debug("  FirstName: " & FST_NAME)
                mydebuglog.Debug("  MidName: " & MID_NAME)
                mydebuglog.Debug("  LastName: " & LAST_NAME)
                mydebuglog.Debug("  Gender: " & GENDER)
                mydebuglog.Debug("  FullName: " & FULL_NAME)
                mydebuglog.Debug("  PartId: " & PART_ID)
                'mydebuglog.Debug("  DOB: " & DOB)
                mydebuglog.Debug("  DOB: " & "Hidden")
                'mydebuglog.Debug("  SSN: " & SSN)
                mydebuglog.Debug("  SSN: " & "Hidden")
                mydebuglog.Debug("  WorkPhone: " & WORK_PHONE)
                mydebuglog.Debug("  EmailAddr: " & EMAIL_ADDR)
                mydebuglog.Debug("  ConId: " & CON_ID)
                mydebuglog.Debug("  AddrId: " & ADDR_ID)
                mydebuglog.Debug("  AddrMatch: " & ADDR_MATCH)
                mydebuglog.Debug("  AddrType: " & ADDR_TYPE)
                mydebuglog.Debug("  OrgId: " & ORG_ID)
                mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
                mydebuglog.Debug("  OrgPhone: " & ORG_PHONE)
                mydebuglog.Debug("  RegNum: " & REG_NUM)
                mydebuglog.Debug("  SubConId: " & SUB_CON_ID)
                mydebuglog.Debug("  TrainerNo: " & TRAINER_NO)
                mydebuglog.Debug("  HomePhone: " & HOME_PHONE)
                mydebuglog.Debug("  JobTitle: " & JOB_TITLE)
                mydebuglog.Debug("  PerTitle: " & PER_TITLE)
                mydebuglog.Debug("  Source: " & SOURCE)
                mydebuglog.Debug("  Confidence: " & Confidence)
            End If

            ' ============================================
            ' Do not clean the record if there is no contact to clean
            ' Ignore certain names like "ACCOUNTS PAYABLE"; 9/16/2020; Ren Hou;
            Dim skipClean As Boolean = False
            If database <> "C" Then
                'DIRECTOR','HUMAN','DISTRICT','SERVICE','F&B' -- First_name
                Dim skipFnameArr() As String = "DIRECTOR,HUMAN,DISTRICT,SERVICE,F&B,HR".Split(",")
                'MANAGER','RESOURCES','PAYABLE','ADMINISTRATOR' -- Last_name
                Dim skipLnameArr() As String = "MANAGER,RESOURCES,PAYABLE,ADMINISTRATOR,DIRECTOR".Split(",")
                For Each fname In skipFnameArr
                    If FST_NAME = fname Then skipClean = True
                Next
                For Each lname In skipLnameArr
                    If LAST_NAME = lname Then skipClean = True
                Next
            End If
            If debug = "Y" Then mydebuglog.Debug("--skipClean bypass: " & skipClean.ToString + vbCrLf)
            If FST_NAME <> "" And LAST_NAME <> "" And skipClean <> True Then  ' Ignore "ACCOUNTS PAYABLE"; 9/16/2020; Ren Hou;

                ' ============================================
                ' Call StandardizeContact to update the record if the contact is unknown
                If CON_ID = "" Then
                    wp = "<Contacts><Contact>"
                    wp = wp & "<Debug>N</Debug>"
                    wp = wp & "<Database>X</Database>"
                    wp = wp & "<ConId>" & HttpUtility.UrlEncode(CON_ID) & "</ConId>"
                    wp = wp & "<FirstName>" & HttpUtility.UrlEncode(FST_NAME) & "</FirstName>"
                    wp = wp & "<MidName>" & HttpUtility.UrlEncode(MID_NAME) & "</MidName>"
                    wp = wp & "<LastName>" & HttpUtility.UrlEncode(LAST_NAME) & "</LastName>"
                    wp = wp & "<Gender>" & GENDER & "</Gender>"
                    wp = wp & "<FullName>" & HttpUtility.UrlEncode(FULL_NAME) & "</FullName>"
                    wp = wp & "</Contact></Contacts>"
                    Try
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "==========================" & vbCrLf & "Calling StandardizeContact " & vbCrLf & "  sXML: " & wp)
                        rDoc = StandardizeContact(wp)
                        If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
                        rNodeList = rDoc.SelectNodes("//Contact")
                        For j = 0 To rNodeList.Count - 1
                            Try
                                If debug = "Y" Then mydebuglog.Debug("  Found node: " & j.ToString)
                                FST_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("FirstName", rNodeList.Item(j))))
                                MID_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("MidName", rNodeList.Item(j))))
                                LAST_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("LastName", rNodeList.Item(j))))
                                GENDER = Trim(RemovePluses(GetNodeValue("Gender", rNodeList.Item(j))))
                                FULL_NAME = HttpUtility.UrlDecode(GetNodeValue("FullName", rNodeList.Item(j)))
                                If FULL_NAME = "" Then
                                    If MID_NAME = "" Then
                                        FULL_NAME = FST_NAME & " " & LAST_NAME
                                    Else
                                        FULL_NAME = FST_NAME & " " & MID_NAME & " " & LAST_NAME
                                    End If
                                End If
                                CON_ID = HttpUtility.UrlDecode(Trim(GetNodeValue("ConId", rNodeList.Item(j))))
                                MATCH_CODE = HttpUtility.UrlDecode(GetNodeValue("MatchCode", rNodeList.Item(j)))
                                If GENDER <> "" And PER_TITLE = "" Then
                                    If GENDER = "M" Then PER_TITLE = "Mr."
                                    If GENDER = "F" Then PER_TITLE = "Ms."
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                                results = "Failure"
                                GoTo CloseOut2
                            End Try
                        Next
                        If debug = "Y" Then mydebuglog.Debug("  Standardized: " & results)
                        If results <> "Success" Then GoTo CloseOut

                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
                    End Try
                    If debug = "Y" Then mydebuglog.Debug("  MatchCode : " & MATCH_CODE)
                End If

                ' ============================================
                ' Create SQL-safe strings for testing
                Dim sFST_NAME, sLAST_NAME, sMID_NAME, sJOB_TITLE As String
                sFST_NAME = SqlString(FST_NAME)
                sLAST_NAME = SqlString(LAST_NAME)
                sMID_NAME = SqlString(MID_NAME)
                sJOB_TITLE = SqlString(JOB_TITLE)
                pLAST_NAME = UCase(SqlString(Left(LAST_NAME, 3)))
                pFST_NAME = UCase(SqlString(Left(FST_NAME, 3)))

                ' ============================================
                ' Locate match
                ' If contact id exists, assume we already have a record and don't need to load it
                If debug = "Y" Then mydebuglog.Debug("  Get Match Candidates: " & CON_ID)
                If CON_ID = "" Then

                    ' Look for an existing record at least one was found
                    'If returnv > 0 Then
                    SqlS = "SELECT TOP 250 C.ROW_ID, O.ROW_ID, O.NAME, O.LOC, O.MAIN_PH_NUM, O.X_ACCOUNT_NUM, C.X_PART_ID, " & _
                    "C.X_REGISTRATION_NUM, SC.ROW_ID, C.BIRTH_DT, C.SOC_SECURITY_NUM, O.DEDUP_TOKEN, C.X_MATCH_CD, " & _
                    "A.ROW_ID, A.ADDR, A.CITY, A.STATE, A.ZIPCODE, A.COUNTRY, A.X_MATCH_CD, " & _
                    "P.ROW_ID, P.ADDR, P.CITY, P.STATE, P.ZIPCODE, P.COUNTRY, P.X_MATCH_CD, C.X_MATCH_CD, C.WORK_PH_NUM, " & _
                    "C.EMAIL_ADDR, C.HOME_PH_NUM, C.X_TRAINER_NUM, UPPER(C.FST_NAME), UPPER(C.LAST_NAME), C.JOB_TITLE, " & _
                    "C.PER_TITLE, UPPER(C.MID_NAME), C.SEX_MF, OA.X_MATCH_CD  " & _
                    "FROM siebeldb.dbo.S_CONTACT C WITH (INDEX([S_CONTACT_P1])) " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT O WITH (INDEX([S_ORG_EXT_P1])) ON O.ROW_ID=C.PR_DEPT_OU_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG OA WITH (INDEX([S_ADDR_ORG_P1])) ON OA.ROW_ID=O.PR_ADDR_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_SUB_CON SC WITH (INDEX([CX_SUB_CON_C1])) ON SC.CON_ID=C.ROW_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A WITH (INDEX([S_ADDR_ORG_P1])) ON A.ROW_ID=C.PR_OU_ADDR_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER P WITH (INDEX([S_ADDR_PER_P1])) ON P.ROW_ID=C.PR_PER_ADDR_ID " & _
                    "WHERE "
                    If LAST_NAME <> "" And FST_NAME <> "" Then
                        SqlS = SqlS & "(UPPER(C.LAST_NAME)='" & UCase(Trim(sLAST_NAME)) & "' AND UPPER(LEFT(C.FST_NAME,3))='" & pFST_NAME & "') OR "
                    End If
                    If MID_NAME <> "" Then SqlS = SqlS & "(C.MID_NAME='" & sMID_NAME & "' AND UPPER(C.LAST_NAME)='" & UCase(Trim(sLAST_NAME)) & "' AND UPPER(LEFT(C.FST_NAME,3))='" & pFST_NAME & "') OR "
                    If WORK_PHONE <> "" Then SqlS = SqlS & "(C.WORK_PH_NUM='" & WORK_PHONE & "' AND UPPER(C.LAST_NAME)='" & UCase(Trim(sLAST_NAME)) & "' AND UPPER(LEFT(C.FST_NAME,3))='" & pFST_NAME & "') OR "
                    'If SSN <> "" And Len(SSN) = 11 Then SqlS = SqlS & "(C.SOC_SECURITY_NUM='" & SSN & "' AND C.SOC_SECURITY_NUM IS NOT NULL AND C.SOC_SECURITY_NUM<>'') OR "
                    If SSN <> "" And Len(SSN) = 11 Then SqlS = SqlS & " (C.SSN_IDX=HASHBYTES('SHA2_256','" & SSN & "') AND C.SSN_IDX<>HASHBYTES('SHA2_256','') AND C.SSN_IDX is not null) OR "
                    'If DOB <> "" Then SqlS = SqlS & "(UPPER(C.LAST_NAME)='" & UCase(Trim(sLAST_NAME)) & "' AND C.BIRTH_DT='" & DOB & "' AND C.BIRTH_DT IS NOT NULL) OR "
                    If DOB <> "" Then SqlS = SqlS & "(UPPER(C.LAST_NAME)='" & UCase(Trim(sLAST_NAME)) & "' AND C.DOB_IDX=HASHBYTES('SHA2_256',CONVERT(varchar,CONVERT(datetime,'" & DOB & "'))) AND C.DOB_IDX IS NOT NULL) OR "
                    If PART_ID <> "" Then SqlS = SqlS & "(C.X_PART_ID='" & PART_ID & "') OR "
                    If ORG_ID <> "" Then SqlS = SqlS & "(C.PR_DEPT_OU_ID='" & ORG_ID & "' AND (UPPER(C.LAST_NAME) LIKE '%" & UCase(Trim(sLAST_NAME)) & "%' OR UPPER(C.FST_NAME) LIKE '%" & UCase(Trim(sFST_NAME)) & "%')) OR "
                    If MATCH_CODE <> "" And MATCH_CODE <> "$$$$$$$$$$$$$$$" Then SqlS = SqlS & "(C.X_MATCH_CD='" & MATCH_CODE & "') OR "
                    If EMAIL_ADDR <> "" Then
                        If Right(SqlS, 6) <> "WHERE " And Right(SqlS, 4) <> " OR " Then SqlS = SqlS & " OR "
                        SqlS = SqlS & "(UPPER(C.LAST_NAME)='" & UCase(Trim(sLAST_NAME)) & "' AND UPPER(C.EMAIL_ADDR)='" & SqlString(UCase(EMAIL_ADDR)) & "') OR "
                        SqlS = SqlS & "(UPPER(C.FST_NAME)='" & UCase(Trim(sFST_NAME)) & "' AND UPPER(C.EMAIL_ADDR)='" & SqlString(UCase(EMAIL_ADDR)) & "') OR "
                    End If
                    If ADDR_MATCH <> "" Then
                        Select Case ADDR_TYPE
                            Case "P"
                                SqlS = SqlS & "(P.X_MATCH_CD='" & ADDR_MATCH & "' AND UPPER(C.LAST_NAME) LIKE '%" & UCase(Trim(sLAST_NAME)) & "%') OR "
                            Case "B"
                                SqlS = SqlS & "(A.X_MATCH_CD='" & ADDR_MATCH & "' AND UPPER(C.LAST_NAME) LIKE '%" & UCase(Trim(sLAST_NAME)) & "%') OR "
                        End Select
                    End If
                    If ORG_MATCH <> "" Then
                        SqlS = SqlS & "(O.DEDUP_TOKEN='" & ORG_MATCH & "' AND UPPER(C.LAST_NAME) LIKE '%" & UCase(Trim(sLAST_NAME)) & "%' AND UPPER(C.FST_NAME) LIKE '%" & UCase(Trim(sFST_NAME)) & "%') OR "
                    End If
                    If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                    If debug = "Y" Then mydebuglog.Debug("  Get contact matches: " & SqlS & vbCrLf)
                    returnv = 0
                    Try
                        'cmd.CommandText = SqlS 
                        'dr = cmd.ExecuteReader()
                        '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                        cmd_ro.CommandText = SqlS
                        dr = cmd_ro.ExecuteReader()

                        Dim enc_con As SqlConnection = New SqlConnection(ConnS)
                        enc_con.Open()
                        Dim enc_dr As SqlDataReader
                        Dim enc_cmd As SqlCommand = New SqlCommand()
                        enc_cmd.Connection = enc_con
                        If Not dr Is Nothing Then
                            ' Check for best match based on confidence
                            While dr.Read()
                                Try
                                    match_count = 0         ' Reset match count for each record
                                    returnv = returnv + 1

                                    ' ----- 
                                    ' Store record in temp fields for testing
                                    tCON_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                    '*** Get decrypted SSN and DOB; 12/19/2019; Ren Hou; modified for decripting SSN and DOB and improving query efficiency;

                                    enc_cmd.CommandText = "EXEC reports.dbo.OpenHCIKeys;" & _
                                                        "SELECT top 1 CONVERT(datetime, reports.dbo.HCI_Decrypt(ENC_BIRTH_DT)) BIRTH_DT " & _
                                                        "   , reports.dbo.HCI_Decrypt(ENC_SOC_SECURITY_NUM) SOC_SECURITY_NUM " & _
                                                        "FROM siebeldb.dbo.S_CONTACT " & _
                                                        "WHERE ROW_ID = '" & tCON_ID & "'; " & _
                                                        "EXEC reports.dbo.CloseHCIKeys;"
                                    enc_dr = enc_cmd.ExecuteReader()
                                    While enc_dr.Read()
                                        tDOB = StndDate(Trim(CheckDBNull(enc_dr(0), enumObjectType.StrType)))
                                        tSSN = StndSSN(Trim(CheckDBNull(enc_dr(1), enumObjectType.StrType)))
                                    End While
                                    enc_dr.Dispose()
                                    '*** end modification ***
                                    tFST_NAME = Trim(CheckDBNull(dr(32), enumObjectType.StrType))
                                    tLAST_NAME = Trim(CheckDBNull(dr(33), enumObjectType.StrType))
                                    tMID_NAME = Trim(CheckDBNull(dr(36), enumObjectType.StrType))
                                    tFULL_NAME = tFST_NAME & " "
                                    If tMID_NAME <> "" Then tFULL_NAME = tFULL_NAME & tMID_NAME & " "
                                    tFULL_NAME = tFULL_NAME & tLAST_NAME
                                    tEMAIL_ADDR = Trim(CheckDBNull(dr(29), enumObjectType.StrType))
                                    tPART_ID = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                    tREG_NUM = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                    tSUB_CON_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                    tTRAINER_NO = Trim(CheckDBNull(dr(31), enumObjectType.StrType))
                                    tWORK_PHONE = Trim(CheckDBNull(dr(28), enumObjectType.StrType))
                                    'If debug = "Y" Then mydebuglog.Debug("    .. tWORK_PHONE: " & tWORK_PHONE)
                                    If tWORK_PHONE <> "" Then tWORK_PHONE = StndPhone(tWORK_PHONE)
                                    'If debug = "Y" Then mydebuglog.Debug("    .. tWORK_PHONE: " & tWORK_PHONE)
                                    tHOME_PHONE = Trim(CheckDBNull(dr(30), enumObjectType.StrType))
                                    If tHOME_PHONE <> "" Then tHOME_PHONE = StndPhone(tHOME_PHONE)
                                    tMATCH_CODE = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                                    tORG_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                    tADDR_ID = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                                    tADDR_MATCH = Trim(CheckDBNull(dr(19), enumObjectType.StrType))
                                    tPER_ADDR_ID = Trim(CheckDBNull(dr(20), enumObjectType.StrType))
                                    tPADDR_MATCH = Trim(CheckDBNull(dr(26), enumObjectType.StrType))
                                    'tDOB = StndDate(Trim(CheckDBNull(dr(9), enumObjectType.StrType)))
                                    'tSSN = StndSSN(Trim(CheckDBNull(dr(10), enumObjectType.StrType)))
                                    tJOB_TITLE = Trim(CheckDBNull(dr(34), enumObjectType.StrType))
                                    tPER_TITLE = Trim(CheckDBNull(dr(35), enumObjectType.StrType))
                                    tGENDER = Trim(CheckDBNull(dr(37), enumObjectType.StrType))
                                    tORG_MATCH = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                    tORG_PHONE = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                    If tORG_PHONE <> "" Then tORG_PHONE = StndPhone(tORG_PHONE)
                                    tOADDR_MATCH = Trim(CheckDBNull(dr(38), enumObjectType.StrType))

                                    ' -----
                                    ' If contact id match, then assume correct
                                    If CON_ID <> "" And CON_ID = tCON_ID Then
                                        If debug = "Y" Then mydebuglog.Debug("   > Found exact match based on id: " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)) & "" & Trim(CheckDBNull(dr(32), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(33), enumObjectType.StrType)))

                                        ' Save the values from the exact match
                                        S_CON_ID = tCON_ID
                                        S_FST_NAME = tFST_NAME
                                        S_LAST_NAME = tLAST_NAME
                                        S_MID_NAME = tMID_NAME
                                        S_EMAIL_ADDR = tEMAIL_ADDR
                                        S_PART_ID = tPART_ID
                                        S_REG_NUM = tREG_NUM
                                        S_SUB_CON_ID = tSUB_CON_ID
                                        S_TRAINER_NO = tTRAINER_NO
                                        S_WORK_PHONE = tWORK_PHONE
                                        S_HOME_PHONE = tHOME_PHONE
                                        S_MATCH_CODE = tMATCH_CODE
                                        S_ORG_ID = tORG_ID
                                        S_ADDR_ID = tADDR_ID
                                        S_PER_ADDR_ID = tPER_ADDR_ID
                                        S_DOB = tDOB
                                        S_SSN = tSSN
                                        S_JOB_TITLE = tJOB_TITLE
                                        S_PER_TITLE = tPER_TITLE
                                        S_GENDER = tGENDER
                                        high_count = Confidence
                                        GoTo UpdConMatch
                                    End If

                                    ' Check/fix Match Code
                                    If tMATCH_CODE = "" And MATCH_CODE <> "" Then
                                        If UCase(FST_NAME) = UCase(tLAST_NAME) And UCase(LAST_NAME) = UCase(tLAST_NAME) And UCase(MID_NAME) = UCase(tMID_NAME) Then tMATCH_CODE = MATCH_CODE
                                    End If

                                    ' -----
                                    ' Check matches when there is a match code
                                    If MATCH_CODE <> "" And MATCH_CODE <> "$$$$$$$$$$$$$$$" Then
                                        ' If the contact match code exists, then check secondary factors
                                        ' Match based on match code to an existing record
                                        If debug = "Y" Then mydebuglog.Debug("  Matchcode found: " & tMATCH_CODE)
                                        If MATCH_CODE = tMATCH_CODE Then
                                            If debug = "Y" Then mydebuglog.Debug("  * Testing matchcode candidate record ID: " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(32), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(33), enumObjectType.StrType)))
                                            match_count = 2
                                            ' -----
                                            ' The match code matches - check for sub-matches
                                            ' Check to see if the address is the same
                                            If ADDR_MATCH <> "" And ADDR_MATCH <> "$$$$$$$$$$$$$$$" Then
                                                If ADDR_MATCH = tADDR_MATCH Or ADDR_MATCH = tPADDR_MATCH Or ADDR_MATCH = tOADDR_MATCH Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found address based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If
                                            If ADDR_ID <> "" Then
                                                If ADDR_ID = tADDR_ID Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found address based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check to see if the organization is the same
                                            If ORG_MATCH <> "" And ORG_MATCH <> "$$$$$$$$$$$$$$$" Then
                                                If ORG_MATCH = tORG_MATCH Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found organization based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If
                                            If ORG_ID <> "" Then
                                                If ORG_ID = tORG_ID Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found organization based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check to see if contact work phone is the same
                                            If WORK_PHONE <> "" Then
                                                If WORK_PHONE = tWORK_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found work phone to work phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                                If WORK_PHONE = tHOME_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found work phone to home phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                                If WORK_PHONE = tORG_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found work phone to org phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check to see if contact home phone is the same
                                            If HOME_PHONE <> "" Then
                                                If HOME_PHONE = tWORK_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found home phone to work phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                                If HOME_PHONE = tHOME_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found home phone to home phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                                If HOME_PHONE = tORG_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found home phone to org phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for contact's employer phone number matches
                                            If ORG_PHONE <> "" Then
                                                If ORG_PHONE = tORG_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found employer phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for email address matches
                                            If EMAIL_ADDR <> "" And UCase(EMAIL_ADDR) = UCase(tEMAIL_ADDR) Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found email address based match for contact")
                                                match_count = match_count + 1
                                            End If

                                            ' Check for participant id matches
                                            If PART_ID <> "" Then
                                                If PART_ID = tPART_ID Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found participant match for contact")
                                                    match_count = match_count + 1
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > Participant does not match for contact")
                                                    match_count = match_count - 1
                                                End If
                                            End If

                                            ' Check for trainer number matches
                                            If TRAINER_NO <> "" Then
                                                If TRAINER_NO = tTRAINER_NO Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found trainer match for contact")
                                                    match_count = match_count + 5
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > Trainer does not match for contact")
                                                    match_count = match_count - 5
                                                End If
                                            End If

                                            ' Check for DOB matches
                                            If DOB <> "" And tDOB <> "" Then
                                                If DOB = tDOB Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found DOB match for contact")
                                                    If UCase(FST_NAME) = UCase(tFST_NAME) Or SSN = tSSN Then
                                                        match_count = match_count + 2
                                                    Else
                                                        match_count = match_count + 1
                                                    End If
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > DOB does not match for contact")
                                                    match_count = match_count - 1
                                                End If
                                            End If

                                            ' Check for SSN matches
                                            If SSN <> "" And tSSN <> "" Then
                                                If SSN = tSSN Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found SSN match for contact")
                                                    If UCase(FST_NAME) = UCase(tFST_NAME) Or DOB = tDOB Then
                                                        match_count = match_count + 2
                                                    Else
                                                        match_count = match_count + 1
                                                    End If
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > SSN does not match for contact")
                                                    match_count = match_count - 1
                                                End If
                                            End If

                                            ' Check for registration number matches
                                            If REG_NUM <> "" Then
                                                If REG_NUM = tREG_NUM Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found registration number match for contact")
                                                    match_count = match_count + 2
                                                End If
                                            End If

                                            ' Save the values from the highest match
                                            If match_count > high_count Then
                                                S_CON_ID = tCON_ID
                                                S_FST_NAME = tFST_NAME
                                                S_LAST_NAME = tLAST_NAME
                                                S_MID_NAME = tMID_NAME
                                                S_EMAIL_ADDR = tEMAIL_ADDR
                                                S_PART_ID = tPART_ID
                                                S_REG_NUM = tREG_NUM
                                                S_SUB_CON_ID = tSUB_CON_ID
                                                S_TRAINER_NO = tTRAINER_NO
                                                S_WORK_PHONE = tWORK_PHONE
                                                S_HOME_PHONE = tHOME_PHONE
                                                S_MATCH_CODE = tMATCH_CODE
                                                S_ORG_ID = tORG_ID
                                                S_ADDR_ID = tADDR_ID
                                                S_PER_ADDR_ID = tPER_ADDR_ID
                                                S_DOB = tDOB
                                                S_SSN = tSSN
                                                S_JOB_TITLE = tJOB_TITLE
                                                S_PER_TITLE = tPER_TITLE
                                                S_GENDER = tGENDER
                                                high_count = match_count
                                            End If
                                        Else
                                            ' -----
                                            ' The match code doesn't match
                                            ' Try to match without using the match code.  
                                            ' This automatically reduces the match score
                                            match_count = -1
                                            If debug = "Y" Then mydebuglog.Debug("  * Testing non-matchcode candidate record ID: " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(32), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(33), enumObjectType.StrType)))
                                            ' Check for full name match
                                            'If debug = "Y" Then mydebuglog.Debug("   - Checking FST_NAME : " & FST_NAME & " - " & tFST_NAME)
                                            If UCase(FST_NAME) = UCase(tFST_NAME) Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found first name match for contact")
                                                match_count = match_count + 1
                                            End If
                                            'If debug = "Y" Then mydebuglog.Debug("   - Checking LAST_NAME : " & LAST_NAME & " - " & tLAST_NAME)
                                            If UCase(LAST_NAME) = UCase(tLAST_NAME) Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found last name match for contact")
                                                match_count = match_count + 1
                                            End If

                                            ' Check for address matches
                                            If ADDR_MATCH <> "" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking ADDR_MATCH : " & ADDR_MATCH & " - " & tADDR_MATCH & " / " & tPADDR_MATCH & " / " & tOADDR_MATCH)
                                                If ADDR_MATCH = tADDR_MATCH Or ADDR_MATCH = tPADDR_MATCH Or ADDR_MATCH = tOADDR_MATCH Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found address based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If
                                            If ADDR_ID <> "" And ADDR_MATCH <> "$$$$$$$$$$$$$$$" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking ADDR_ID : " & ADDR_ID & " - " & tADDR_ID)
                                                If ADDR_ID = tADDR_ID Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found address based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for organization based matches
                                            If ORG_MATCH <> "" And ORG_MATCH <> "$$$$$$$$$$$$$$$" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking ORG_MATCH : " & ORG_MATCH & " - " & tORG_MATCH)
                                                If ORG_MATCH = tORG_MATCH Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found organization based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If
                                            If ORG_ID <> "" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking ORG_ID : " & ORG_ID & " - " & tORG_ID)
                                                If ORG_ID = tORG_ID Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found organization based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for work phone matches
                                            If WORK_PHONE <> "" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking WORK_PHONE : " & WORK_PHONE & " - " & tWORK_PHONE)
                                                If WORK_PHONE = tWORK_PHONE Then
                                                    match_count = match_count + 1
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found work phone to work phone based match for contact")
                                                End If
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking WORK_PHONE : " & WORK_PHONE & " - " & tHOME_PHONE)
                                                If WORK_PHONE = tHOME_PHONE Then
                                                    match_count = match_count + 1
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found work phone to home phone based match for contact")
                                                End If
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking WORK_PHONE : " & WORK_PHONE & " - " & tORG_PHONE)
                                                If WORK_PHONE = tORG_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found work phone to org phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for home phone matches
                                            If HOME_PHONE <> "" Then
                                                match_count = match_count + 1
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking HOME_PHONE : " & HOME_PHONE & " - " & tWORK_PHONE)
                                                If HOME_PHONE = tWORK_PHONE Then
                                                    match_count = match_count + 1
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found home phone to work phone based match for contact")
                                                End If
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking HOME_PHONE : " & HOME_PHONE & " - " & tHOME_PHONE)
                                                If HOME_PHONE = tHOME_PHONE Then
                                                    match_count = match_count + 1
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found home phone to home phone based match for contact")
                                                End If
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking HOME_PHONE : " & HOME_PHONE & " - " & tORG_PHONE)
                                                If HOME_PHONE = tORG_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found home phone to org phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for contact's employer phone number matches
                                            If ORG_PHONE <> "" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking ORG_PHONE : " & ORG_PHONE & " - " & tORG_PHONE)
                                                If ORG_PHONE = tORG_PHONE Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found employer phone based match for contact")
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for email address matches
                                            'If debug = "Y" Then mydebuglog.Debug("   - Checking EMAIL_ADDR : " & tEMAIL_ADDR)
                                            If EMAIL_ADDR <> "" And UCase(EMAIL_ADDR) = UCase(tEMAIL_ADDR) Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found email address based match for contact")
                                                If UCase(FST_NAME) = UCase(tFST_NAME) And UCase(LAST_NAME) = UCase(tLAST_NAME) Then
                                                    ' If the name also matches then give a higher weight to this match
                                                    match_count = match_count + 2
                                                Else
                                                    match_count = match_count + 1
                                                End If
                                            End If

                                            ' Check for participant id matches
                                            If PART_ID <> "" Then
                                                ' If debug = "Y" Then mydebuglog.Debug("   - Checking PART_ID : " & tPART_ID)
                                                If PART_ID = tPART_ID Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found participant match for contact")
                                                    If UCase(FST_NAME) = UCase(tFST_NAME) And UCase(LAST_NAME) = UCase(tLAST_NAME) Then
                                                        ' If the name also matches then give a higher weight to this match
                                                        match_count = match_count + 2
                                                    Else
                                                        match_count = match_count + 1
                                                    End If
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > Participant does not match for contact")
                                                    match_count = match_count - 1
                                                End If
                                            End If

                                            ' Check for trainer number matches
                                            If TRAINER_NO <> "" Then
                                                If TRAINER_NO = tTRAINER_NO Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found trainer match for contact")
                                                    match_count = match_count + 5
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > Trainer does not match for contact")
                                                    match_count = match_count - 5
                                                End If
                                            End If

                                            ' Check for DOB matches
                                            If DOB <> "" And tDOB <> "" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking DOB : " & tDOB)
                                                If DOB = tDOB Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found DOB match for contact")
                                                    If UCase(FST_NAME) = UCase(tFST_NAME) Or SSN = tSSN Then
                                                        match_count = match_count + 2
                                                    Else
                                                        match_count = match_count + 1
                                                    End If
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > DOB does not match for contact")
                                                    match_count = match_count - 1
                                                End If
                                            End If

                                            ' Check for SSN matches
                                            If SSN <> "" And tSSN <> "" Then
                                                'If debug = "Y" Then mydebuglog.Debug("   - Checking SSN : " & tSSN)
                                                If SSN = tSSN Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found SSN match for contact")
                                                    If UCase(FST_NAME) = UCase(tFST_NAME) Or DOB = tDOB Then
                                                        match_count = match_count + 2
                                                    Else
                                                        match_count = match_count + 1
                                                    End If
                                                Else
                                                    If debug = "Y" Then mydebuglog.Debug("   > SSN does not match for contact")
                                                    match_count = match_count - 1
                                                End If
                                            End If

                                            ' Check for registration number matches
                                            If REG_NUM <> "" Then
                                                If REG_NUM = tREG_NUM Then
                                                    If debug = "Y" Then mydebuglog.Debug("   > Found registration number match for contact")
                                                    match_count = match_count + 2
                                                End If
                                            End If

                                            ' Save the values from the highest match
                                            If match_count > high_count Then
                                                S_CON_ID = tCON_ID
                                                S_FST_NAME = tFST_NAME
                                                S_LAST_NAME = tLAST_NAME
                                                S_MID_NAME = tMID_NAME
                                                S_EMAIL_ADDR = tEMAIL_ADDR
                                                S_PART_ID = tPART_ID
                                                S_REG_NUM = tREG_NUM
                                                S_SUB_CON_ID = tSUB_CON_ID
                                                S_TRAINER_NO = tTRAINER_NO
                                                S_WORK_PHONE = tWORK_PHONE
                                                S_HOME_PHONE = tHOME_PHONE
                                                S_MATCH_CODE = tMATCH_CODE
                                                S_ORG_ID = tORG_ID
                                                S_ADDR_ID = tADDR_ID
                                                S_PER_ADDR_ID = tPER_ADDR_ID
                                                S_DOB = tDOB
                                                S_SSN = tSSN
                                                S_JOB_TITLE = tJOB_TITLE
                                                S_PER_TITLE = tPER_TITLE
                                                S_GENDER = tGENDER
                                                high_count = match_count
                                            End If

                                        End If
                                    Else
                                        match_count = -1
                                        ' Match based on other things
                                        ' There is less confidence if there is no match code
                                        ' If the first and last name match then try to confirm based on other criteria
                                        If debug = "Y" Then mydebuglog.Debug("  * Testing miscellaneous candidate record ID: " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(32), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(33), enumObjectType.StrType)))

                                        If UCase(FST_NAME) = UCase(tFST_NAME) Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found first name match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If UCase(LAST_NAME) = UCase(tLAST_NAME) Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found last name match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If WORK_PHONE <> "" And WORK_PHONE = tWORK_PHONE Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found work phone to work phone based match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If WORK_PHONE <> "" And WORK_PHONE = tHOME_PHONE Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found work phone to home phone based match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If WORK_PHONE <> "" And WORK_PHONE = tORG_PHONE Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found work phone to org phone based match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If HOME_PHONE <> "" And HOME_PHONE = tWORK_PHONE Then
                                            match_count = match_count + 1
                                            If debug = "Y" Then mydebuglog.Debug("   > Found home phone to work phone based match for contact")
                                        End If

                                        If HOME_PHONE <> "" And HOME_PHONE = tHOME_PHONE Then
                                            match_count = match_count + 1
                                            If debug = "Y" Then mydebuglog.Debug("   > Found home phone to home phone based match for contact")
                                        End If

                                        If HOME_PHONE <> "" And HOME_PHONE = tORG_PHONE Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found home phone to org phone based match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If EMAIL_ADDR <> "" And UCase(EMAIL_ADDR) = UCase(tEMAIL_ADDR) Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found email address based match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If PART_ID <> "" Then
                                            If PART_ID = tPART_ID Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found participant match for contact")
                                                match_count = match_count + 2
                                            Else
                                                If debug = "Y" Then mydebuglog.Debug("   > Participant does not match for contact")
                                                match_count = match_count - 1
                                            End If
                                        End If

                                        If TRAINER_NO <> "" Then
                                            If TRAINER_NO = tTRAINER_NO Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found trainer match for contact")
                                                match_count = match_count + 5
                                            Else
                                                If debug = "Y" Then mydebuglog.Debug("   > Trainer does not match for contact")
                                                match_count = match_count - 5
                                            End If
                                        End If

                                        If DOB <> "" And tDOB <> "" Then
                                            If DOB = tDOB Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found DOB match for contact")
                                                match_count = match_count + 2
                                            Else
                                                If debug = "Y" Then mydebuglog.Debug("   > DOB does not match for contact")
                                                match_count = match_count - 1
                                            End If
                                        End If

                                        If SSN <> "" And tSSN <> "" Then
                                            If SSN = tSSN Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found SSN match for contact")
                                                match_count = match_count + 2
                                            Else
                                                If debug = "Y" Then mydebuglog.Debug("   > SSN does not match for contact")
                                                match_count = match_count - 1
                                            End If
                                        End If

                                        If ADDR_MATCH <> "" And ADDR_MATCH <> "$$$$$$$$$$$$$$$" Then
                                            If ADDR_MATCH = tADDR_MATCH Or ADDR_MATCH = tPADDR_MATCH Or ADDR_MATCH = tOADDR_MATCH Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found address based match for contact")
                                                match_count = match_count + 1
                                            End If
                                        End If

                                        If ADDR_ID <> "" Then
                                            If ADDR_ID = tADDR_ID Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found address based match for contact")
                                                match_count = match_count + 1
                                            End If
                                        End If

                                        If ORG_ID <> "" Then
                                            If ORG_ID = tORG_ID Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found organization based match for contact")
                                                match_count = match_count + 1
                                            End If
                                        End If

                                        If ORG_PHONE <> "" And ORG_PHONE = tORG_PHONE Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found employer phone based match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If ORG_MATCH <> "" And ORG_MATCH = tORG_MATCH Then
                                            If debug = "Y" Then mydebuglog.Debug("   > Found employer match code based match for contact")
                                            match_count = match_count + 1
                                        End If

                                        If REG_NUM <> "" Then
                                            If REG_NUM = tREG_NUM Then
                                                If debug = "Y" Then mydebuglog.Debug("   > Found registration number match for contact")
                                                match_count = match_count + 2
                                            End If
                                        End If

                                        ' Save the values from the highest match
                                        If match_count > high_count Then
                                            S_CON_ID = tCON_ID
                                            S_FST_NAME = tFST_NAME
                                            S_LAST_NAME = tLAST_NAME
                                            S_MID_NAME = tMID_NAME
                                            S_EMAIL_ADDR = tEMAIL_ADDR
                                            S_PART_ID = tPART_ID
                                            S_REG_NUM = tREG_NUM
                                            S_SUB_CON_ID = tSUB_CON_ID
                                            S_TRAINER_NO = tTRAINER_NO
                                            S_WORK_PHONE = tWORK_PHONE
                                            S_HOME_PHONE = tHOME_PHONE
                                            S_MATCH_CODE = tMATCH_CODE
                                            S_ORG_ID = tORG_ID
                                            S_ADDR_ID = tADDR_ID
                                            S_PER_ADDR_ID = tPER_ADDR_ID
                                            S_DOB = tDOB
                                            S_SSN = tSSN
                                            S_JOB_TITLE = tJOB_TITLE
                                            S_PER_TITLE = tPER_TITLE
                                            S_GENDER = tGENDER
                                            high_count = match_count
                                        End If
                                    End If
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading S_CONTACT: " & ex.ToString & vbCrLf
                                End Try
                                If debug = "Y" Then mydebuglog.Debug("     >>Match score: " & match_count.ToString & vbCrLf)
                            End While
                            enc_cmd.Dispose()
                            enc_con.Close()
                            enc_con.Dispose()

                            ' If the high count exceeds the confidence, declare a match
                            If high_count >= Confidence Then GoTo UpdConMatch

                            ' If not, then declare no match found
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "No contact matched.  Score: " & match_count.ToString & vbCrLf)
                            GoTo FinishMatch
UpdConMatch:
                            ' Extract data from matching contact
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Contact matched.  Record id: " & S_CON_ID & "  Score: " & high_count.ToString & vbCrLf)
                            CON_ID = S_CON_ID
                            If EMAIL_ADDR = "" Then EMAIL_ADDR = S_EMAIL_ADDR
                            PART_ID = S_PART_ID
                            REG_NUM = S_REG_NUM
                            SUB_CON_ID = S_SUB_CON_ID
                            If IsNumeric(S_TRAINER_NO) And Val(S_TRAINER_NO) > 0 Then
                                TRAINER_NO = S_TRAINER_NO
                            End If
                            If WORK_PHONE <> "" Then WORK_PHONE = StndPhone(WORK_PHONE)
                            If HOME_PHONE <> "" Then HOME_PHONE = StndPhone(HOME_PHONE)
                            MATCH_CODE = S_MATCH_CODE
                            ORG_ID = S_ORG_ID
                            ADDR_ID = S_ADDR_ID
                            JOB_TITLE = S_JOB_TITLE
                            PER_TITLE = S_PER_TITLE
                            GENDER = S_GENDER
                            If DOB = "" And S_DOB <> "" And S_DOB <> "1/1/1900" Then DOB = S_DOB
                            If SSN = "" And S_SSN <> "" Then SSN = S_SSN
                        End If

                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error locating contact record. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                    'End If
                End If

FinishMatch:
                Try
                    dr.Close()
                Catch ex As Exception
                End Try

                ' Debug output saved record
                If debug = "Y" And S_CON_ID <> "" Then
                    mydebuglog.Debug(vbCrLf & "==SAVED RECORD")
                    mydebuglog.Debug("  FirstName: " & S_FST_NAME)
                    mydebuglog.Debug("  MidName: " & S_MID_NAME)
                    mydebuglog.Debug("  LastName: " & S_LAST_NAME)
                    mydebuglog.Debug("  Gender: " & S_GENDER)
                    mydebuglog.Debug("  PartId: " & S_PART_ID)
                    'mydebuglog.Debug("  DOB: " & S_DOB)
                    mydebuglog.Debug("  DOB: " & "Hidden")
                    'mydebuglog.Debug("  SSN: " & S_SSN)
                    mydebuglog.Debug("  SSN: " & "Hidden")
                    mydebuglog.Debug("  WorkPhone: " & S_WORK_PHONE)
                    mydebuglog.Debug("  EmailAddr: " & S_EMAIL_ADDR)
                    mydebuglog.Debug("  ConId: " & S_CON_ID)
                    mydebuglog.Debug("  AddrId: " & S_ADDR_ID)
                    mydebuglog.Debug("  PerAddrId: " & S_PER_ADDR_ID)
                    mydebuglog.Debug("  OrgId: " & S_ORG_ID)
                    mydebuglog.Debug("  RegNum: " & S_REG_NUM)
                    mydebuglog.Debug("  SubConId: " & S_SUB_CON_ID)
                    mydebuglog.Debug("  TrainerNo: " & S_TRAINER_NO)
                    mydebuglog.Debug("  HomePhone: " & S_HOME_PHONE)
                    mydebuglog.Debug("  JobTitle: " & S_JOB_TITLE)
                    mydebuglog.Debug("  PerTitle: " & S_PER_TITLE & vbCrLf)
                End If

                ' ============================================
                ' Database operations
                If database = "U" And CON_ID = "" Then database = "C"

                ' Create record
                If database = "C" Then
                    If CON_ID = "" Then
                        ' Generate a new contact id
                        CON_ID = BasicService.GenerateRecordId("S_CONTACT", "N", debug)
                        If debug = "Y" Then mydebuglog.Debug("  Generated CON_ID: " & CON_ID)

                        ' Create contact record 
                        SqlS = "INSERT siebeldb.dbo.S_CONTACT (ROW_ID,CREATED,CREATED_BY," & _
                        "LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,CONFLICT_ID,BU_ID," & _
                        "FST_NAME,LAST_NAME,MID_NAME,SEX_MF,X_MATCH_CD,X_MATCH_DT,LOGIN," & _
                        "EMAIL_ADDR,COMMENTS,CONSUMER_FLG,CON_CD,PR_POSTN_ID) " & _
                        "SELECT TOP 1 '" & CON_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0,0,'0-R9NH'," & _
                        "'" & SqlString(FST_NAME) & "','" & SqlString(LAST_NAME) & "','" & SqlString(MID_NAME) & _
                        "','" & GENDER & "','" & MATCH_CODE & "',GETDATE(),'" & SqlString(FULL_NAME) & "','" & _
                        SqlString(EMAIL_ADDR) & "','From CleanContact web service','N','Business','0-5220' " & _
                        "FROM siebeldb.dbo.S_CONTACT WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & CON_ID & "')"
                        temp = ExecQuery("Create", "Contact", cmd, SqlS, mydebuglog, debug)
                        errmsg = errmsg & temp

                        ' Enrich the new contact with additional information
                        SqlS = ""
                        If DOB <> "" And IsDate(DOB) Then SqlS = SqlS & ",BIRTH_DT='" & DOB & "'"
                        If WORK_PHONE <> "" Then SqlS = SqlS & ",WORK_PH_NUM='" & Left(WORK_PHONE, 40) & "'"
                        If SSN <> "" Then SqlS = SqlS & ",SOC_SECURITY_NUM='" & Left(SSN, 11) & "'"
                        If GENDER <> "" Then SqlS = SqlS & ",SEX_MF='" & Left(GENDER, 1) & "'"
                        If ADDR_TYPE = "P" Then
                            If ADDR_ID <> "" And ADDR_ID <> "+" Then SqlS = SqlS & ",PR_PER_ADDR_ID='" & Left(ADDR_ID, 15) & "'"
                        Else
                            If ADDR_ID <> "" And ADDR_ID <> "+" Then SqlS = SqlS & ",PR_OU_ADDR_ID='" & Left(ADDR_ID, 15) & "'"
                        End If
                        If EMAIL_ADDR <> "" Then
                            SqlS = SqlS & ",PREF_COMM_MEDIA_CD='Email'"
                        Else
                            If WORK_PHONE <> "" Or HOME_PHONE <> "" Then
                                SqlS = SqlS & ",PREF_COMM_MEDIA_CD='Email'"
                            Else
                                If ADDR_ID <> "" Then SqlS = SqlS & ",PREF_COMM_MEDIA_CD='Direct Mail'"
                            End If
                        End If
                        If ORG_ID <> "" And ORG_ID <> "+" Then SqlS = SqlS & ",PR_DEPT_OU_ID='" & Left(ORG_ID, 15) & "'"
                        If PART_ID <> "" Then SqlS = SqlS & ",X_PART_ID='" & PART_ID & "',CONSUMER_FLG='Y'"
                        If REG_NUM <> "" Then SqlS = SqlS & ",X_REGISTRATION_NUM='" & Left(REG_NUM, 15) & "'"
                        If HOME_PHONE <> "" Then SqlS = SqlS & ",HOME_PH_NUM='" & Left(HOME_PHONE, 40) & "'"
                        If JOB_TITLE <> "" Then SqlS = SqlS & ",JOB_TITLE='" & Left(SqlString(JOB_TITLE), 75) & "'"
                        If PER_TITLE <> "" And PER_TITLE <> " " Then SqlS = SqlS & ",PER_TITLE='" & Left(PER_TITLE, 15) & "'"
                        If SqlS <> "" Then
                            SqlS = "UPDATE siebeldb.dbo.S_CONTACT SET LAST_UPD=GETDATE()" & SqlS & " WHERE ROW_ID='" & CON_ID & "'"
                            If TRAINER_NO <> "" Then
                                mydebuglog.Debug("  Is TRAINER: Y (TRAINER_NO: " & TRAINER_NO & ")")
                                Dim oldRecSqlS As String = "EXEC reports.dbo.OpenHCIKeys;Select reports.dbo.HCI_Decrypt(ENC_SOC_SECURITY_NUM) 'SOC_SECURITY_NUM', reports.dbo.HCI_Decrypt(ENC_PASSWORD) 'X_PASSWORD', reports.dbo.HCI_Decrypt(ENC_BIRTH_DT) 'BIRTH_DT', * from siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & CON_ID & "'"
                                LogTrainerDataChanges(con_ro, TRAINER_NO, SqlS, oldRecSqlS, "C:\Logs\TrainerChanges.log", mydebuglog)
                            Else
                                mydebuglog.Debug("  Is TRAINER: N")
                            End If
                            temp = ExecQuery("Update", "Contact", cmd, SqlS, mydebuglog, debug)
                            errmsg = errmsg & temp
                        End If

                        ' Process a source
                        If Trim(SOURCE) <> "" And InStr(SOURCE, "<", CompareMethod.Text) = 0 And SOURCE <> "Registration - " Then
                            ' Locate source category id
                            SqlS = "SELECT C.ROW_ID, CC.ROW_ID  " & _
                            "FROM siebeldb.dbo.S_CHRCTR C " & _
                            "LEFT OUTER JOIN siebeldb.dbo.S_CON_CHRCTR CC ON CC.CHRCTR_ID=C.ROW_ID AND CC.CONTACT_ID='" & CON_ID & "' " & _
                            "WHERE C.OBJ_TYPE_CD='Contact' AND C.NAME='" & SOURCE & "'"
                            If debug = "Y" Then mydebuglog.Debug("  Looking for source: " & SqlS)
                            Try
                                'cmd.CommandText = SqlS
                                'dr = cmd.ExecuteReader()
                                '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                                cmd_ro.CommandText = SqlS
                                dr = cmd_ro.ExecuteReader()

                                If Not dr Is Nothing Then
                                    While dr.Read()
                                        Try
                                            SOURCE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                            CON_CHRCTR_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                        Catch ex As Exception
                                            errmsg = errmsg & "Error reading source: " & ex.ToString & vbCrLf
                                        End Try
                                    End While
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error reading source. " & ex.ToString
                                GoTo CloseOut
                            End Try
                            dr.Close()

                            ' Create the contact category record from the id
                            If SOURCE_ID <> "" And CON_CHRCTR_ID = "" Then
                                SqlS = "INSERT siebeldb.dbo.S_CON_CHRCTR " & _
                                "(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " & _
                                "MODIFICATION_NUM, CONFLICT_ID, CHRCTR_ID, PRIV_FLG, CONTACT_ID) " & _
                                "SELECT TOP 1 '" & CON_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0,0,'" & SOURCE_ID & "','N','" & CON_ID & "' " & _
                                "FROM siebeldb.dbo.S_CON_CHRCTR WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_CON_CHRCTR WHERE ROW_ID='" & CON_ID & "')"
                                temp = ExecQuery("Create", "Contact category", cmd, SqlS, mydebuglog, debug)
                                'errmsg = errmsg & temp
                            End If
                        End If

                        ' Verify the record was written
                        SqlS = "SELECT COUNT(*) FROM siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & CON_ID & "'"
                        If debug = "Y" Then mydebuglog.Debug("  Verifying uniqueness: " & SqlS)
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            '2020-05-18; Ren Hou; Added for read-only per Chris;
                            '2021-07-08: Rebecca; The AG does not replicate fast enough to permit the RO member to be used for this check
                            'cmd_ro.CommandText = SqlS
                            'dr = cmd_ro.ExecuteReader()

                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        returnv = CheckDBNull(dr(0), enumObjectType.IntType)
                                    Catch ex As Exception
                                        errmsg = errmsg & "Error verifying new S_CONTACT: " & ex.ToString & vbCrLf
                                    End Try
                                End While
                            End If
                        Catch ex As Exception
                            errmsg = errmsg & vbCrLf & "Error verifying new contact record. " & ex.ToString
                            results = "Failure"
                            GoTo CloseOut
                        End Try
                        dr.Close()
                        If returnv > 0 Then
                            ' Update participant record if applicable
                            If PART_ID <> "" Then
                                SqlS = "UPDATE siebeldb.dbo.CX_PARTICIPANT_X " & _
                                "SET CON_ID='" & CON_ID & "' " & _
                                "WHERE ROW_ID='" & PART_ID & "' AND (CON_ID IS NULL OR CON_ID='')"
                                temp = ExecQuery("Update", "Participant", cmd, SqlS, mydebuglog, debug)
                                errmsg = errmsg & temp
                            End If

                            ' Create extension record
                            SqlS = "INSERT INTO siebeldb.dbo.S_CONTACT_X " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM,MODIFICATION_NUM,CONFLICT_ID,PAR_ROW_ID) " & _
                            "SELECT TOP 1 '" & CON_ID & "',GETDATE(), '0-1', GETDATE(), '0-1', 0, 0, 0, '" & CON_ID & "' " & _
                            "FROM siebeldb.dbo.S_CONTACT_X WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_CONTACT_X WHERE ROW_ID='" & CON_ID & "')"
                            temp = ExecQuery("Create", "Contact extension", cmd, SqlS, mydebuglog, debug)
                            errmsg = errmsg & temp

                            ' Create contact position record
                            SqlS = "INSERT INTO siebeldb.dbo.S_POSTN_CON " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,CONFLICT_ID,CON_FST_NAME, CON_ID, CON_LAST_NAME, POSTN_ID, ROW_STATUS, ASGN_DNRM_FLG, ASGN_MANL_FLG, ASGN_SYS_FLG, STATUS) " & _
                            "SELECT TOP 1 '" & CON_ID & "',GETDATE(), '0-1', GETDATE(), '0-1', 0, 0, '" & SqlString(FST_NAME) & "', '" & _
                            CON_ID & "', '" & SqlString(LAST_NAME) & "', '0-5220', 'Y', 'N', 'Y', 'N', 'Active' " & _
                            "FROM siebeldb.dbo.S_POSTN_CON WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_POSTN_CON WHERE ROW_ID='" & CON_ID & "')"
                            temp = ExecQuery("Create", "Contact position", cmd, SqlS, mydebuglog, debug)
                            errmsg = errmsg & temp
                        End If
                    Else
                        'errmsg = errmsg & vbCrLf & ">>Contact Id already existed.  No record created. Merging instead"
                        'results = "Failure"
                        database = "M"
                    End If
                End If

                '-----
                ' Update record
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Operation: " & database & " on id '" & CON_ID & "'")
                If database = "U" Then
                    If CON_ID <> "" Then
                        Try
                            SqlS = ""
                            If FST_NAME <> "" Then SqlS = SqlS & ",FST_NAME='" & Left(SqlString(FST_NAME), 50) & "'"
                            If LAST_NAME <> "" Then SqlS = SqlS & ",LAST_NAME='" & Left(SqlString(LAST_NAME), 50) & "'"
                            If MID_NAME <> "" Then SqlS = SqlS & ",MID_NAME='" & Left(SqlString(MID_NAME), 50) & "'"
                            If GENDER <> "" Then SqlS = SqlS & ",SEX_MF='" & Left(GENDER, 1) & "'"
                            If MATCH_CODE <> "" Then SqlS = SqlS & ",X_MATCH_CD='" & MATCH_CODE & "',X_MATCH_DT=GETDATE() "
                            If DOB <> "" And IsDate(DOB) Then SqlS = SqlS & ",BIRTH_DT='" & DOB & "'"
                            If WORK_PHONE <> "" Then SqlS = SqlS & ",WORK_PH_NUM='" & Left(WORK_PHONE, 40) & "'"
                            If SSN <> "" Then SqlS = SqlS & ",SOC_SECURITY_NUM='" & Left(SSN, 11) & "'"
                            If EMAIL_ADDR <> "" Then SqlS = SqlS & ",EMAIL_ADDR='" & Left(SqlString(EMAIL_ADDR), 50) & "'"
                            If REG_NUM <> "" Then SqlS = SqlS & ",X_REGISTRATION_NUM='" & Left(REG_NUM, 15) & "'"
                            If HOME_PHONE <> "" Then SqlS = SqlS & ",HOME_PH_NUM='" & Left(HOME_PHONE, 40) & "'"
                            If JOB_TITLE <> "" Then SqlS = SqlS & ",JOB_TITLE='" & Left(SqlString(JOB_TITLE), 75) & "'"
                            If PER_TITLE <> "" Then SqlS = SqlS & ",PER_TITLE='" & Left(PER_TITLE, 15) & "'"
                            If PART_ID <> "" Then SqlS = SqlS & ",X_PART_ID='" & PART_ID & "'"
                            If ADDR_ID <> "" Then
                                If ADDR_TYPE = "P" Then
                                    SqlS = SqlS & ",PR_PER_ADDR_ID='" & ADDR_ID & "'"
                                Else
                                    SqlS = SqlS & ",PR_OU_ADDR_ID='" & ADDR_ID & "'"
                                End If
                            End If
                            If ORG_ID <> "" Then SqlS = SqlS & ",PR_DEPT_OU_ID='" & ORG_ID & "'"

                            If SqlS <> "" Then  '2012-06-29; Ren Hou; Modified the SqlS to avoid running the UPDATE stement if the SET fields are empty;
                                SqlS = "UPDATE siebeldb.dbo.S_CONTACT SET LAST_UPD=GETDATE() " & SqlS
                                SqlS = SqlS & " WHERE ROW_ID='" & CON_ID & "'"
                                If TRAINER_NO <> "" Then
                                    mydebuglog.Debug("  Is TRAINER: Y (TRAINER_NO: " & TRAINER_NO & ")")
                                    'Dim oldRecSqlS As String = "EXEC reports.dbo.OpenHCIKeys;Select reports.dbo.HCI_Decrypt(ENC_SOC_SECURITY_NUM) 'SOC_SECURITY_NUM', reports.dbo.HCI_Decrypt(ENC_PASSWORD) 'X_PASSWORD', reports.dbo.HCI_Decrypt(ENC_BIRTH_DT) 'BIRTH_DT', * from siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & CON_ID & "'"
                                    'LogTrainerDataChanges(con_ro, TRAINER_NO, SqlS, oldRecSqlS, "C:\Logs\TrainerChanges.log", mydebuglog)
                                Else
                                    mydebuglog.Debug("  Is TRAINER: N")
                                End If
                                temp = ExecQuery("Update", "Contact", cmd, SqlS, mydebuglog, debug)
                                If debug = "Y" And temp <> "" Then mydebuglog.Debug("   >> ERROR MESSAGE: " & temp)
                                If debug = "Y" Then mydebuglog.Debug("   >> Current results: " & results)
                                errmsg = errmsg & temp
                            End If
                        Catch ex As Exception
                            If debug = "Y" Then mydebuglog.Debug("   >> Error updating contact: " & ex.ToString & vbCrLf & "SqlS:" & SqlS)
                            errmsg = errmsg & vbCrLf & "Error updating contact: " & ex.ToString
                            results = "Failure"
                        End Try

                        ' Process a source
                        If Trim(SOURCE) <> "" And InStr(SOURCE, "<", CompareMethod.Text) = 0 And SOURCE <> "Registration - " Then
                            ' Locate source category id
                            SqlS = "SELECT C.ROW_ID, CC.ROW_ID  " & _
                            "FROM siebeldb.dbo.S_CHRCTR C " & _
                            "LEFT OUTER JOIN siebeldb.dbo.S_CON_CHRCTR CC ON CC.CHRCTR_ID=C.ROW_ID AND CC.CONTACT_ID='" & CON_ID & "' " & _
                            "WHERE C.OBJ_TYPE_CD='Contact' AND C.NAME='" & SOURCE & "'"
                            If debug = "Y" Then mydebuglog.Debug("  Looking for source: " & SqlS)
                            Try
                                'cmd.CommandText = SqlS
                                'dr = cmd.ExecuteReader()
                                '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                                cmd_ro.CommandText = SqlS
                                dr = cmd_ro.ExecuteReader()
                                If Not dr Is Nothing Then
                                    While dr.Read()
                                        Try
                                            SOURCE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                            CON_CHRCTR_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                        Catch ex As Exception
                                            errmsg = errmsg & "Error reading source: " & ex.ToString & vbCrLf
                                        End Try
                                    End While
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error reading source. " & ex.ToString
                                results = "Failure"
                                GoTo CloseOut
                            End Try
                            dr.Close()

                            ' Create the contact category record from the id
                            If SOURCE_ID <> "" And CON_CHRCTR_ID = "" Then
                                SqlS = "INSERT siebeldb.dbo.S_CON_CHRCTR " & _
                                "(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " & _
                                "MODIFICATION_NUM, CONFLICT_ID, CHRCTR_ID, PRIV_FLG, CONTACT_ID) " & _
                                "SELECT TOP 1 '" & CON_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0,0,'" & SOURCE_ID & "','N','" & CON_ID & "' " & _
                                "FROM siebeldb.dbo.S_CON_CHRCTR WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_CON_CHRCTR WHERE ROW_ID='" & CON_ID & "')"
                                temp = ExecQuery("Create", "Contact category record", cmd, SqlS, mydebuglog, debug)
                                'errmsg = errmsg & temp
                            End If
                        End If
                    Else
                        errmsg = errmsg & vbCrLf & "Contact Id error on updating record. "
                        results = "Failure"
                    End If
                End If

                '-----
                ' Merge records
                If database = "M" Then
                    If CON_ID <> "" Then

                        If S_PART_ID <> O_PART_ID And O_PART_ID <> "" And S_PART_ID <> "" Then
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Merging participant '" & O_PART_ID & "' into '" & S_PART_ID & "'")
                            ' Call stored procedure sp_MergeParticipants to merge the new one into the old
                            ' Parameters:
                            '  @p_deleted_part_id varchar(15) 
                            '  , @p_kept_part_id varchar(15) 
                            '  , @p_User varchar(15) = NULL --'0-1' 
                            '  , @p_merge_attribute_flg bit = 0 -- default is no merge
                            '  , @p_called_by_contact_merge bit = 0 -- Optional parameter; default is not called by sp_MergeContacts SP 
                            '   @p_kept_contact_id varchar(15) = NULL -- Optional; CON_ID used to update CON_ID field of the kept Participant. But if the @p_called_by_contact_merge parameter is 1, this is required.
                            Try
                                Dim spcmd As SqlCommand = New SqlCommand("siebeldb.dbo.sp_MergeParticipants", con)
                                spcmd.CommandType = Data.CommandType.StoredProcedure
                                spcmd.Parameters.Add("@p_deleted_part_id", Data.SqlDbType.VarChar)
                                spcmd.Parameters("@p_deleted_part_id").Value = O_PART_ID
                                spcmd.Parameters.Add("@p_kept_part_id", Data.SqlDbType.VarChar)
                                spcmd.Parameters("@p_kept_part_id").Value = S_PART_ID
                                spcmd.Parameters.Add("@p_User", Data.SqlDbType.VarChar)
                                spcmd.Parameters("@p_User").Value = "0-1"
                                spcmd.Parameters.Add("@p_merge_attribute_flg", Data.SqlDbType.Bit)
                                spcmd.Parameters("@p_merge_attribute_flg").Value = 1
                                spcmd.Parameters.Add("@p_called_by_contact_merge", Data.SqlDbType.Bit)
                                spcmd.Parameters("@p_called_by_contact_merge").Value = 0
                                spcmd.Parameters.Add("@p_kept_contact_id", Data.SqlDbType.VarChar)
                                spcmd.Parameters("@p_kept_contact_id").Value = CON_ID
                                results = Str(spcmd.ExecuteNonQuery())
                                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "    > Results: '" & results)
                                spcmd.Parameters.Clear()
                                spcmd = Nothing

                                ' Update CONVERT_CONTACT_QUEUE record if applicable
                                SqlS = "UPDATE scanner.dbo.CONVERT_CONTACT_QUEUE SET PART_ID='" & S_PART_ID & "' WHERE PART_ID='" & O_PART_ID & "'"
                                temp = ExecQuery("Update", "Convert Contact Queue record", cmd, SqlS, mydebuglog, debug)

                                ' Update CONVERT_IMAGES_QUEUE record if applicable
                                SqlS = "UPDATE scanner.dbo.CONVERT_IMAGES_QUEUE SET PART_ID='" & S_PART_ID & "' WHERE PART_ID='" & O_PART_ID & "'"
                                temp = ExecQuery("Update", "Convert Image Queue record", cmd, SqlS, mydebuglog, debug)

                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error executing sp_MergeParticipants: " & ex.ToString
                                results = "Failure"
                                GoTo CloseOut
                            End Try
                        End If
                        SqlS = ""
                        If FST_NAME <> "" And S_FST_NAME = "" Then SqlS = SqlS & ",FST_NAME='" & Left(SqlString(FST_NAME), 50) & "'"
                        If LAST_NAME <> "" And S_LAST_NAME = "" Then SqlS = SqlS & ",LAST_NAME='" & Left(SqlString(LAST_NAME), 50) & "'"
                        If MID_NAME <> "" And S_MID_NAME = "" Then SqlS = SqlS & ",MID_NAME='" & Left(SqlString(MID_NAME), 50) & "'"
                        If GENDER <> "" And S_GENDER = "" Then SqlS = SqlS & ",SEX_MF='" & Left(GENDER, 1) & "'"
                        If MATCH_CODE <> "" And S_MATCH_CODE = "" Then SqlS = SqlS & ",X_MATCH_CD='" & MATCH_CODE & "',X_MATCH_DT=GETDATE() "
                        If DOB <> "" And IsDate(DOB) And S_DOB = "" Then SqlS = SqlS & ",BIRTH_DT='" & DOB & "'"
                        If WORK_PHONE <> "" And S_WORK_PHONE = "" Then SqlS = SqlS & ",WORK_PH_NUM='" & Left(WORK_PHONE, 40) & "'"
                        If SSN <> "" And S_SSN = "" Then SqlS = SqlS & ",SOC_SECURITY_NUM='" & Left(SSN, 11) & "'"
                        If EMAIL_ADDR <> "" And S_EMAIL_ADDR = "" Then SqlS = SqlS & ",EMAIL_ADDR='" & Left(SqlString(EMAIL_ADDR), 50) & "'"
                        If REG_NUM <> "" And S_REG_NUM = "" Then SqlS = SqlS & ",X_REGISTRATION_NUM='" & Left(REG_NUM, 15) & "'"
                        If HOME_PHONE <> "" And S_HOME_PHONE = "" Then SqlS = SqlS & ",HOME_PH_NUM='" & Left(HOME_PHONE, 40) & "'"
                        If JOB_TITLE <> "" And S_JOB_TITLE = "" Then SqlS = SqlS & ",JOB_TITLE='" & Left(SqlString(JOB_TITLE), 75) & "'"
                        If PER_TITLE <> "" And S_PER_TITLE = "" Then SqlS = SqlS & ",PER_TITLE='" & Left(PER_TITLE, 15) & "'"
                        If PART_ID <> "" And S_PART_ID = "" Then SqlS = SqlS & ",X_PART_ID='" & PART_ID & "'"
                        If ADDR_ID <> "" Then
                            If ADDR_TYPE = "P" Then
                                If S_PER_ADDR_ID = "" Then
                                    SqlS = SqlS & ",PR_PER_ADDR_ID='" & ADDR_ID & "'"
                                End If
                            Else
                                If S_ADDR_ID = "" Then
                                    SqlS = SqlS & ",PR_OU_ADDR_ID='" & ADDR_ID & "'"
                                End If
                            End If
                        End If
                        If ORG_ID <> "" And S_ORG_ID = "" Then SqlS = SqlS & ",PR_DEPT_OU_ID='" & ORG_ID & "'"
                        If SqlS <> "" Then  '2012-06-29; Ren Hou; Modified the SqlS to avoid running the UPDATE stement if the SET fields are empty;
                            SqlS = "UPDATE siebeldb.dbo.S_CONTACT SET LAST_UPD=GETDATE() " & SqlS
                            SqlS = SqlS & " WHERE ROW_ID='" & CON_ID & "'"
                            If TRAINER_NO <> "" Then
                                mydebuglog.Debug("  Is TRAINER: Y (TRAINER_NO: " & TRAINER_NO & ")")
                                'Dim oldRecSqlS As String = "EXEC reports.dbo.OpenHCIKeys;Select reports.dbo.HCI_Decrypt(ENC_SOC_SECURITY_NUM) 'SOC_SECURITY_NUM', reports.dbo.HCI_Decrypt(ENC_PASSWORD) 'X_PASSWORD', reports.dbo.HCI_Decrypt(ENC_BIRTH_DT) 'BIRTH_DT', * from siebeldb.dbo.S_CONTACT WHERE ROW_ID='" & CON_ID & "'"
                                'LogTrainerDataChanges(con_ro, TRAINER_NO, SqlS, oldRecSqlS, "C:\Logs\TrainerChanges.log", mydebuglog)
                            Else
                                mydebuglog.Debug("  Is TRAINER: N")
                            End If
                            temp = ExecQuery("Merge", "Contact record", cmd, SqlS, mydebuglog, debug)
                            errmsg = errmsg & temp
                        End If

                        ' Process a source
                        If Trim(SOURCE) <> "" And InStr(SOURCE, "<", CompareMethod.Text) = 0 And SOURCE <> "Registration - " Then
                            ' Locate source category id
                            SqlS = "SELECT C.ROW_ID, CC.ROW_ID  " & _
                            "FROM siebeldb.dbo.S_CHRCTR C " & _
                            "LEFT OUTER JOIN siebeldb.dbo.S_CON_CHRCTR CC ON CC.CHRCTR_ID=C.ROW_ID AND CC.CONTACT_ID='" & CON_ID & "' " & _
                            "WHERE C.OBJ_TYPE_CD='Contact' AND C.NAME='" & SOURCE & "'"
                            If debug = "Y" Then mydebuglog.Debug("  Looking for source: " & SqlS)
                            Try
                                'cmd.CommandText = SqlS
                                'dr = cmd.ExecuteReader()
                                '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                                cmd_ro.CommandText = SqlS
                                dr = cmd_ro.ExecuteReader()
                                If Not dr Is Nothing Then
                                    While dr.Read()
                                        Try
                                            SOURCE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                            CON_CHRCTR_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                        Catch ex As Exception
                                            errmsg = errmsg & "Error reading source: " & ex.ToString & vbCrLf
                                        End Try
                                    End While
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error reading source. " & ex.ToString
                                results = "Failure"
                                GoTo CloseOut
                            End Try
                            dr.Close()

                            ' Create the contact category record from the id
                            If SOURCE_ID <> "" And CON_CHRCTR_ID = "" Then
                                SqlS = "INSERT siebeldb.dbo.S_CON_CHRCTR " & _
                                "(ROW_ID, CREATED, CREATED_BY, LAST_UPD, LAST_UPD_BY, " & _
                                "MODIFICATION_NUM, CONFLICT_ID, CHRCTR_ID, PRIV_FLG, CONTACT_ID) " & _
                                "SELECT TOP 1 '" & CON_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0,0,'" & SOURCE_ID & "','N','" & CON_ID & "' " & _
                                "FROM siebeldb.dbo.S_CON_CHRCTR WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_CON_CHRCTR WHERE ROW_ID='" & CON_ID & "')"
                                temp = ExecQuery("Create", "Contact category record", cmd, SqlS, mydebuglog, debug)
                                'errmsg = errmsg & temp
                            End If
                        End If
                    Else
                        errmsg = errmsg & vbCrLf & "Contact Id error on updating record. "
                        results = "Failure"
                    End If
                End If

            Else
                'errmsg = errmsg & vbCrLf & "Contact record is blank. "
                results = "Failure"
                GoTo CloseOut
            End If
        Next

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Close()
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
            con_ro.Close()
            con_ro.Dispose()
            con_ro = Nothing
            cmd_ro.Dispose()
            cmd_ro = Nothing
        Catch ex As Exception
        End Try

CloseOut2:
        ' ============================================
        ' Return the cleaned/deduped information as an XML document:
        '   <Contact>
        '       <FirstName>   
        '       <MidName>
        '       <LastName>
        '       <FullName>
        '       <Gender>
        '       <MatchCode>
        '       <ConId>
        '       <PartId>
        '       <DOB></DOB>             
        '       <WorkPhone></WorkPhone> 
        '       <SSN></SSN>             
        '       <EmailAddr></EmailAddr> 
        '       <RegNum>
        '       <SubId>
        '       <TrainerNo>
        '       <HomePhone>
        '       <JobTitle>
        '       <PerTitle>
        '       <Source>
        '   </Contact>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("Contact")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            If debug <> "T" And (MATCH_CODE <> "" Or CON_ID <> "") Then
                AddXMLChild(odoc, resultsRoot, "FirstName", IIf(FST_NAME = "", " ", HttpUtility.UrlEncode(FST_NAME)))
                AddXMLChild(odoc, resultsRoot, "MidName", IIf(MID_NAME = "", " ", HttpUtility.UrlEncode(MID_NAME)))
                AddXMLChild(odoc, resultsRoot, "LastName", IIf(LAST_NAME = "", " ", HttpUtility.UrlEncode(LAST_NAME)))
                AddXMLChild(odoc, resultsRoot, "FullName", IIf(FULL_NAME = "", " ", HttpUtility.UrlEncode(FULL_NAME)))
                AddXMLChild(odoc, resultsRoot, "Gender", IIf(GENDER = "", " ", GENDER))
                AddXMLChild(odoc, resultsRoot, "MatchCode", IIf(MATCH_CODE = "", " ", HttpUtility.UrlEncode(MATCH_CODE)))
                AddXMLChild(odoc, resultsRoot, "ConId", IIf(CON_ID = "", " ", HttpUtility.UrlEncode(CON_ID)))
                AddXMLChild(odoc, resultsRoot, "PartId", IIf(PART_ID = "", " ", HttpUtility.UrlEncode(PART_ID)))
                AddXMLChild(odoc, resultsRoot, "DOB", IIf(DOB = "", " ", HttpUtility.UrlEncode(DOB)))
                AddXMLChild(odoc, resultsRoot, "WorkPhone", IIf(WORK_PHONE = "", " ", HttpUtility.UrlEncode(WORK_PHONE)))
                AddXMLChild(odoc, resultsRoot, "SSN", IIf(SSN = "", " ", HttpUtility.UrlEncode(SSN)))
                AddXMLChild(odoc, resultsRoot, "EmailAddr", IIf(EMAIL_ADDR = "", " ", HttpUtility.UrlEncode(EMAIL_ADDR)))
                AddXMLChild(odoc, resultsRoot, "RegNum", IIf(REG_NUM = "", " ", HttpUtility.UrlEncode(REG_NUM)))
                AddXMLChild(odoc, resultsRoot, "SubId", IIf(SUB_CON_ID = "", " ", HttpUtility.UrlEncode(SUB_CON_ID)))
                AddXMLChild(odoc, resultsRoot, "TrainerNo", IIf(TRAINER_NO = "", " ", HttpUtility.UrlEncode(TRAINER_NO)))
                AddXMLChild(odoc, resultsRoot, "HomePhone", IIf(HOME_PHONE = "", " ", HttpUtility.UrlEncode(HOME_PHONE)))
                AddXMLChild(odoc, resultsRoot, "JobTitle", IIf(JOB_TITLE = "", " ", HttpUtility.UrlEncode(JOB_TITLE)))
                AddXMLChild(odoc, resultsRoot, "PerTitle", IIf(PER_TITLE = "", " ", HttpUtility.UrlEncode(PER_TITLE)))
                AddXMLChild(odoc, resultsRoot, "Source", IIf(SOURCE = "", " ", HttpUtility.UrlEncode(SOURCE)))
            Else
                If MATCH_CODE <> "" Or CON_ID <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("CleanContact : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("CleanContact : Results: " & results & " for " & LAST_NAME & " generated matchcode '" & MATCH_CODE & "' or contact id '" & CON_ID & "'")
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for " & LAST_NAME & " generated matchcode '" & MATCH_CODE & "' or contact id '" & CON_ID & "'")
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "101"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function

    <WebMethod(Description:="Clean an organization record on request")> _
    Public Function CleanOrganization(ByVal sXML As String) As XmlDocument
        ' This function attempts to match the organization supplied to other records. If a match is
        ' found it returns the the matching record.  It will also optionally update the matching
        ' record with changes from the supplied organization.

        ' The input parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <Organizations>
        '       <Organization>
        '           <Debug></Debug>         - A flag to indicate the service is to run in Debug mode or not
        '                                       "Y"  - Yes for debug mode on.. logging on
        '                                       "N"  - No for debug mode off.. logging off
        '                                       "T"  - Test mode on.. logging off
        '           <Database></Database>   - "C" create S_ORG_EXT record(s), 
        '                                       "U" update record, blank do nothing
        '           <Confidence></Confidence>
        '           <OrgId></OrgId>         - The Id of an existing organization, if applicable
        '           <Name></Name>           - Name of organization
        '           <Loc></Loc>             - Location of organization
        '           <FullName></Fullname>   - Full name of organization
        '           <OrgMatch></OrgMatch>   - Organization match code
        '           <OrgPhone></OrgPhone>   - Organization main phone number
        '           <AddrMatch></AddrMatch> - Address match code, if applicable
        '           <AddrId></AddrId>       - Address Id of a related address, if applicable - S_ORG_EXT.PR_ADDR_ID
        '	        <Addr></Addr>		    - Street Address
        '	        <City></City>			- City
        '	        <State></State>			- State or province
        '	        <Zipcode></Zipcode>		- Zipcode or postal code
        '	        <Country></Country>		- Country
        '           <WorkPhone></WorkPhone> - Work phone of a contact at the organization
        '           <Industry></Industry>   - Industry
        '       </Organization>
        '   </Organizations>

        ' web.config Parameters used:
        '   hcidb        - connection string to siebeldb database

        ' Generic variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim rDoc As XmlDocument
        Dim rNodeList As XmlNodeList
        Dim i, j As Integer
        Dim mypath, debug, errmsg, logging, wp As String

        ' Database declarations
        Dim con As SqlConnection, con_ro As SqlConnection
        Dim cmd As SqlCommand, cmd_ro As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String, ConnS_ro As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("CODebugLog")
        Dim logfile, temp As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim ORG_ID, NAME, LOC, FULL_NAME, ORG_MATCH, ORG_PHONE As String
        Dim ADDR_MATCH, ADDR_ID, WORK_PH_NUM, ORG_NUM, CONFLICT_ID As String
        Dim S_ORG_ID, S_NAME, S_LOC, S_FULL_NAME, S_ORG_MATCH, S_ORG_PHONE As String
        Dim S_ADDR_MATCH, S_ADDR_ID, S_WORK_PH_NUM, INDUSTRY, S_INDUSTRY, S_ORG_NUM As String
        Dim CITY, COUNTY, STATE, ADDR, ZIPCODE, COUNTRY As String
        Dim JURIS_ID, LAT, LON, DELIVERABLE, CON_ID, CON_MATCH, NEW_ORG As String
        Dim database, ntemp As String
        Dim temp_phone, temp_match As String
        Dim match_count, Confidence, high_count, ba_count As Integer

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        ORG_ID = ""
        NAME = "Health Communications Inc"
        LOC = ""
        FULL_NAME = ""
        ORG_MATCH = ""
        ORG_PHONE = "800-438-8477"
        ORG_NUM = ""
        ADDR_ID = ""
        ADDR_MATCH = ""
        WORK_PH_NUM = "800-438-8477"
        INDUSTRY = ""
        CITY = ""
        COUNTY = ""
        STATE = ""
        ADDR = ""
        ZIPCODE = ""
        COUNTRY = ""
        JURIS_ID = ""
        LAT = ""
        LON = ""
        DELIVERABLE = ""
        CON_ID = ""
        CON_MATCH = ""
        temp = ""
        database = ""
        Confidence = 2
        SqlS = ""
        returnv = 0
        ConnS = ""
        ConnS_ro = ""
        match_count = 0
        S_ORG_ID = ""
        S_NAME = ""
        S_LOC = ""
        S_FULL_NAME = ""
        S_ORG_MATCH = ""
        S_ORG_PHONE = ""
        S_ORG_NUM = ""
        S_ADDR_MATCH = ""
        S_ADDR_ID = ""
        S_WORK_PH_NUM = ""
        S_INDUSTRY = ""
        NEW_ORG = "N"
        high_count = 0
        CONFLICT_ID = "0"
        ba_count = 0
        ntemp = ""

        ' ============================================
        ' Check parameters
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        'sXML = Server.UrlDecode(sXML)
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//Organizations/Organization")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        debug = UCase(debug)

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
            ConnS_ro = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb_ro").ConnectionString
            If ConnS_ro = "" Then ConnS_ro = "server=;"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("CleanOrganization_debug").ToUpper()
            If temp = "Y" And debug <> "T" Then debug = temp
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\clean_organization_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\CleanOrganization.log"
            Try
                log4net.GlobalContext.Properties("COLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        ' ============================================; 2020-05-18; Ren Hou; Added for read-only per Chris;
        ' Open read-only database connection 
        errmsg = OpenDBConnection(ConnS_ro, con_ro, cmd_ro)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        For i = 0 To oNodeList.Count - 1
            errmsg = ""
            ' ============================================
            ' Extract data from parameter string
            If debug <> "T" Then
                ORG_ID = Trim(Server.UrlDecode(Trim(GetNodeValue("OrgId", oNodeList.Item(i)))))
                ORG_ID = KeySpace(ORG_ID)
                NAME = Server.UrlDecode(Trim(GetNodeValue("Name", oNodeList.Item(i))))
                LOC = Server.UrlDecode(Trim(GetNodeValue("Loc", oNodeList.Item(i))))
                ORG_MATCH = Trim(Server.UrlDecode(GetNodeValue("OrgMatch", oNodeList.Item(i))))
                Try
                    ORG_PHONE = RemovePluses(Trim(Server.UrlDecode(GetNodeValue("OrgPhone", oNodeList.Item(i)))))
                    If ORG_PHONE <> "" Then ORG_PHONE = StndPhone(ORG_PHONE)
                Catch ex As Exception
                End Try
                FULL_NAME = Trim(Server.UrlDecode(GetNodeValue("FullName", oNodeList.Item(i))))
                If FULL_NAME = "" Then
                    If LOC <> "" Then
                        FULL_NAME = Trim(NAME) & " " & LOC
                    Else
                        FULL_NAME = NAME
                    End If
                End If
                Try
                    WORK_PH_NUM = RemovePluses(Trim(Server.UrlDecode(Trim(GetNodeValue("WorkPhone", oNodeList.Item(i))))))
                    If WORK_PH_NUM <> "" Then WORK_PH_NUM = StndPhone(WORK_PH_NUM)
                Catch ex As Exception
                End Try
                INDUSTRY = Trim(Server.UrlDecode(GetNodeValue("Industry", oNodeList.Item(i))))
                INDUSTRY = KeySpace(INDUSTRY)
                If INDUSTRY = "<Please+select>" Then INDUSTRY = ""
                database = GetNodeValue("Database", oNodeList.Item(i))
                ADDR_ID = Server.UrlDecode(Trim(GetNodeValue("AddrId", oNodeList.Item(i))))
                ADDR_ID = KeySpace(ADDR_ID)
                ADDR_MATCH = Trim(Server.UrlDecode(GetNodeValue("AddrMatch", oNodeList.Item(i))))
                CITY = Trim(Server.UrlDecode(GetNodeValue("City", oNodeList.Item(i))))
                STATE = Trim(Server.UrlDecode(GetNodeValue("State", oNodeList.Item(i))))
                ADDR = Trim(Server.UrlDecode(GetNodeValue("Addr", oNodeList.Item(i))))
                ZIPCODE = Trim(Server.UrlDecode(GetNodeValue("Zipcode", oNodeList.Item(i))))
                COUNTRY = Trim(Server.UrlDecode(GetNodeValue("Country", oNodeList.Item(i))))
                temp = Trim(GetNodeValue("Confidence", oNodeList.Item(i)))
                If temp <> "" And IsNumeric(temp) Then
                    Confidence = Int(temp)
                Else
                    If ADDR <> "" Then Confidence = 3
                End If
            End If
            If debug = "Y" Then
                mydebuglog.Debug("  database: " & database)
                mydebuglog.Debug("  OrgId: " & ORG_ID)
                mydebuglog.Debug("  Name: " & NAME)
                mydebuglog.Debug("  Loc: " & LOC)
                mydebuglog.Debug("  FullName: " & FULL_NAME)
                mydebuglog.Debug("  AddrId: " & ADDR_ID)
                mydebuglog.Debug("  AddrMatch: " & ADDR_MATCH)
                mydebuglog.Debug("  Addr: " & ADDR)
                mydebuglog.Debug("  City: " & CITY)
                mydebuglog.Debug("  State: " & STATE)
                mydebuglog.Debug("  Zipcode: " & ZIPCODE)
                mydebuglog.Debug("  Country: " & COUNTRY)
                mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
                mydebuglog.Debug("  OrgPhone: " & ORG_PHONE)
                mydebuglog.Debug("  WorkPhone: " & WORK_PH_NUM)
                mydebuglog.Debug("  Industry: " & INDUSTRY)
                mydebuglog.Debug("  Confidence: " & Confidence)
            End If

            ' ============================================
            ' Call StandardizeOrganization with no update to update the record
            wp = "<Organizations><Organization>"
            wp = wp & "<Debug>N</Debug>"
            wp = wp & "<Database>X</Database>"
            wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
            wp = wp & "<Name>" & HttpUtility.UrlEncode(NAME) & "</Name>"
            wp = wp & "<Loc>" & HttpUtility.UrlEncode(LOC) & "</Loc>"
            wp = wp & "<FullName>" & HttpUtility.UrlEncode(FULL_NAME) & "</FullName>"
            wp = wp & "</Organization></Organizations>"
            Try
                If debug = "Y" Then mydebuglog.Debug("  sXML: " & wp)
                rDoc = StandardizeOrganization(wp)
                If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
                rNodeList = rDoc.SelectNodes("//Organization")
                For j = 0 To rNodeList.Count - 1
                    Try
                        If debug = "Y" Then mydebuglog.Debug("  Processing node: " & j.ToString)
                        NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("Name", rNodeList.Item(j))))
                        LOC = HttpUtility.UrlDecode(Trim(GetNodeValue("Loc", rNodeList.Item(j))))
                        FULL_NAME = HttpUtility.UrlDecode(GetNodeValue("FullName", rNodeList.Item(j)))
                        If FULL_NAME = "" Then
                            If LOC <> "" Then
                                FULL_NAME = Trim(NAME) & " " & LOC
                            Else
                                FULL_NAME = NAME
                            End If
                        End If
                        ORG_MATCH = HttpUtility.UrlDecode(GetNodeValue("MatchCode", rNodeList.Item(j)))
                        ORG_ID = HttpUtility.UrlDecode(Trim(GetNodeValue("OrgId", rNodeList.Item(j))))
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut2
                    End Try
                Next
                If debug = "Y" Then mydebuglog.Debug("  Standardized: " & results)
                If results <> "Success" Then GoTo CloseOut

            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
            End Try
            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "  StandardizeOrganization Results====")
                mydebuglog.Debug("  Name: " & NAME)
                mydebuglog.Debug("  Loc: " & LOC)
                mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
                mydebuglog.Debug("  FullName: " & FULL_NAME)
                mydebuglog.Debug("  =========================" & vbCrLf)
            End If

            ' ============================================
            ' If address provided but not address id 
            If ADDR <> "" And ADDR_ID = "" Then
                If ORG_ID <> "" Then
                    ' If account known then look for that address
                    Call CallCleanAddress(ADDR_ID, ORG_ID, CON_ID, "O", "N", _
                    ADDR, CITY, STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, CON_MATCH, WORK_PH_NUM, _
                    ORG_MATCH, ORG_PHONE, JURIS_ID, debug, mydebuglog, errmsg, results, "X")
                Else
                    ' Standardize the address
                    Call CallStandardizeAddress(ADDR_ID, ORG_ID, CON_ID, "O", "Y", ADDR, CITY, _
                      STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, JURIS_ID, DELIVERABLE, debug, mydebuglog, errmsg, results, "X")
                End If
            End If

            ' ============================================
            ' Verify Organization Id
            If ORG_ID <> "" Then
                SqlS = "SELECT COUNT(*) FROM siebeldb.dbo.S_ORG_EXT WHERE ROW_ID='" & ORG_ID & "'"
                If debug = "Y" Then mydebuglog.Debug("  Count orgs with specified id: " & SqlS)
                Try
                    'cmd.CommandText = SqlS
                    'dr = cmd.ExecuteReader()
                    '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                    cmd_ro.CommandText = SqlS
                    dr = cmd_ro.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                returnv = CheckDBNull(dr(0), enumObjectType.IntType)
                            Catch ex As Exception
                                errmsg = errmsg & "Error locating organization record: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error locating organization record. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If returnv = 0 Then ORG_ID = ""
            End If

            ' ============================================
            ' Locate match
            ' If contact id exists, assume we already have a record and don't need to load it
            If debug = "Y" Then mydebuglog.Debug("  Get Match Candidates: " & ORG_ID)
            If ORG_ID = "" Then
                ' Count records
                SqlS = "SELECT COUNT(*) " & _
                "FROM siebeldb.dbo.S_ORG_EXT O WITH (INDEX([S_ORG_EXT_P1])) " & _
                "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A WITH (INDEX([S_ADDR_ORG_P1])) ON A.ROW_ID=O.PR_ADDR_ID " & _
                "WHERE "
                If ORG_MATCH <> "" And ORG_MATCH <> "$$$$$$$$$$$$$$$" Then SqlS = SqlS & "O.DEDUP_TOKEN='" & ORG_MATCH & "' OR "
                If ORG_PHONE <> "" Then SqlS = SqlS & "O.MAIN_PH_NUM='" & ORG_PHONE & "' OR "
                If WORK_PH_NUM <> "" And ORG_PHONE <> WORK_PH_NUM Then SqlS = SqlS & "O.MAIN_PH_NUM='" & WORK_PH_NUM & "' OR "
                If ADDR_MATCH <> "" Then SqlS = SqlS & "A.X_MATCH_CD='" & ADDR_MATCH & "' OR "
                If ORG_ID <> "" Then SqlS = SqlS & "O.ROW_ID='" & ORG_ID & "' OR "
                If NAME <> "" Then
                    SqlS = SqlS & "(UPPER(O.NAME)='" & SqlString(UCase(NAME)) & "'"
                    If LOC <> "" Then
                        SqlS = SqlS & " AND UPPER(O.LOC)='" & SqlString(UCase(LOC)) & "'"
                    Else
                        'SqlS = SqlS & " AND (O.LOC='' OR O.LOC IS NULL"
                        If CITY <> "" Then SqlS = SqlS & ") OR (UPPER(O.NAME)='" & SqlString(UCase(NAME)) & "' AND UPPER(O.LOC) LIKE '" & SqlString(UCase(CITY)) & "%'"
                    End If
                    SqlS = SqlS & ") "
                    SqlS = SqlS & "OR (UPPER(O.NAME)+' '+UPPER(O.LOC) LIKE '" & SqlString(UCase(FULL_NAME)) & "%') "
                End If
                If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                If debug = "Y" Then mydebuglog.Debug("  Count matching orgs: " & SqlS)
                Try
                    'cmd.CommandText = SqlS
                    'dr = cmd.ExecuteReader()
                    '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                    cmd_ro.CommandText = SqlS
                    dr = cmd_ro.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                returnv = CheckDBNull(dr(0), enumObjectType.IntType)
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading organization: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error locating organization record. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If debug = "Y" Then mydebuglog.Debug("  # Matches found: " & returnv.ToString)

                ' Look for an existing record at least one was found
                If returnv > 0 Then
                    dr.Close()
                    SqlS = "SELECT (SELECT CASE WHEN O.ROW_ID IS NULL THEN '' ELSE O.ROW_ID END), " & _
                    "(SELECT CASE WHEN O.NAME IS NULL THEN '' ELSE O.NAME END), " & _
                    "(SELECT CASE WHEN O.LOC IS NULL THEN '' ELSE O.LOC END), " & _
                    "(SELECT CASE WHEN O.X_ACCOUNT_NUM IS NULL THEN '' ELSE O.X_ACCOUNT_NUM END), " & _
                    "(SELECT CASE WHEN O.DEDUP_TOKEN IS NULL THEN '' ELSE O.DEDUP_TOKEN END), " & _
                    "(SELECT CASE WHEN O.MAIN_PH_NUM IS NULL THEN '' ELSE O.MAIN_PH_NUM END), " & _
                    "(SELECT CASE WHEN A.ROW_ID IS NULL THEN '' ELSE A.ROW_ID END), " & _
                    "(SELECT CASE WHEN A.ADDR IS NULL THEN '' ELSE A.ADDR END), " & _
                    "(SELECT CASE WHEN A.CITY IS NULL THEN '' ELSE A.CITY END), " & _
                    "(SELECT CASE WHEN A.STATE IS NULL THEN '' ELSE A.STATE END), " & _
                    "(SELECT CASE WHEN A.ZIPCODE IS NULL THEN '' ELSE A.ZIPCODE END), " & _
                    "(SELECT CASE WHEN A.COUNTRY IS NULL THEN '' ELSE A.COUNTRY END), " & _
                    "(SELECT CASE WHEN A.X_MATCH_CD IS NULL THEN '' ELSE A.X_MATCH_CD END), " & _
                    "(SELECT CASE WHEN O.PR_INDUST_ID IS NULL THEN '' ELSE O.PR_INDUST_ID END), " & _
                    "(SELECT CASE WHEN O.X_ACCOUNT_NUM IS NULL THEN '' ELSE O.X_ACCOUNT_NUM END), " & _
                    "(SELECT CASE WHEN O.NAME IS NULL THEN '' ELSE O.NAME END)+(SELECT CASE WHEN O.LOC IS NULL THEN '' ELSE ' '+O.LOC END), " & _
                    "(SELECT CASE WHEN O.NAME IS NULL THEN '' ELSE O.NAME END)+(SELECT CASE WHEN A.CITY IS NULL THEN '' ELSE ' '+A.CITY END) " & _
                    "FROM siebeldb.dbo.S_ORG_EXT O WITH (INDEX([S_ORG_EXT_P1])) " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A WITH (INDEX([S_ADDR_ORG_P1])) ON A.ROW_ID=O.PR_ADDR_ID " & _
                    "WHERE "
                    If returnv > 200 And (WORK_PH_NUM <> "" Or ORG_PHONE <> "") Then
                        SqlS = SqlS & "("
                        If ORG_MATCH <> "" And ORG_MATCH <> "$$$$$$$$$$$$$$$" Then SqlS = SqlS & "O.DEDUP_TOKEN='" & ORG_MATCH & "' OR "
                        If ADDR_MATCH <> "" Then SqlS = SqlS & "A.X_MATCH_CD='" & ADDR_MATCH & "' OR "
                        If ORG_ID <> "" Then SqlS = SqlS & "O.ROW_ID='" & ORG_ID & "' OR "
                        If NAME <> "" Then
                            SqlS = SqlS & "(UPPER(O.NAME)='" & SqlString(UCase(NAME)) & "'"
                            If LOC <> "" Then
                                SqlS = SqlS & " AND UPPER(O.LOC)='" & SqlString(UCase(LOC)) & "'"
                            Else
                                'SqlS = SqlS & " AND (O.LOC='' OR O.LOC IS NULL)"
                                If CITY <> "" Then SqlS = SqlS & ") OR (UPPER(O.NAME)='" & SqlString(UCase(NAME)) & "' AND UPPER(O.LOC) LIKE '" & SqlString(UCase(CITY)) & "%'"
                            End If
                            SqlS = SqlS & ")"
                            SqlS = SqlS & "OR (UPPER(O.NAME)+' '+UPPER(O.LOC) LIKE '" & SqlString(UCase(FULL_NAME)) & "%') "
                        End If
                        If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                        SqlS = SqlS & ") AND ("
                        If ORG_PHONE <> "" Then SqlS = SqlS & "O.MAIN_PH_NUM='" & ORG_PHONE & "' OR "
                        If WORK_PH_NUM <> "" And ORG_PHONE <> WORK_PH_NUM Then SqlS = SqlS & "O.MAIN_PH_NUM='" & WORK_PH_NUM & "' OR "
                        If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                        SqlS = SqlS & ")"
                    Else
                        If ORG_MATCH <> "" And ORG_MATCH <> "$$$$$$$$$$$$$$$" Then SqlS = SqlS & "O.DEDUP_TOKEN='" & ORG_MATCH & "' OR "
                        If ORG_PHONE <> "" Then SqlS = SqlS & "O.MAIN_PH_NUM='" & ORG_PHONE & "' OR "
                        If WORK_PH_NUM <> "" And ORG_PHONE <> WORK_PH_NUM Then SqlS = SqlS & "O.MAIN_PH_NUM='" & WORK_PH_NUM & "' OR "
                        If ADDR_MATCH <> "" Then SqlS = SqlS & "A.X_MATCH_CD='" & ADDR_MATCH & "' OR "
                        If ORG_ID <> "" Then SqlS = SqlS & "O.ROW_ID='" & ORG_ID & "' OR "
                        If NAME <> "" Then
                            SqlS = SqlS & "(UPPER(O.NAME)='" & SqlString(UCase(NAME)) & "'"
                            If LOC <> "" Then
                                SqlS = SqlS & " AND UPPER(O.LOC)='" & SqlString(UCase(LOC)) & "'"
                            Else
                                'SqlS = SqlS & " AND (O.LOC='' OR O.LOC IS NULL)"
                                If CITY <> "" Then SqlS = SqlS & ") OR (UPPER(O.NAME)='" & SqlString(UCase(NAME)) & "' AND UPPER(O.LOC) LIKE '" & SqlString(UCase(CITY)) & "%'"
                            End If
                            SqlS = SqlS & ")"
                            SqlS = SqlS & "OR (UPPER(O.NAME)+' '+UPPER(O.LOC) LIKE '" & SqlString(UCase(FULL_NAME)) & "%') "
                        End If
                        If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                    End If
                    If debug = "Y" Then mydebuglog.Debug("  =========================" & vbCrLf & "  Checking matches: " & SqlS)
                    Try
                        'cmd.CommandText = SqlS
                        '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                        cmd_ro.CommandText = SqlS
                        Try
                            'dr = cmd.ExecuteReader()
                            dr = cmd_ro.ExecuteReader()
                        Catch ex2 As Exception
                            If debug = "Y" Then mydebuglog.Debug("  reader execution error: " & ex2.Message)
                        End Try
                        If Not dr Is Nothing Then

                            ' Look for match
                            Try
                                match_count = 0
                                ' Match based on match code to an existing record
                                If ORG_MATCH <> "" And ORG_MATCH <> "$$$$$$$$$$$$$$$" Then
                                    While dr.Read()
                                        temp_phone = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                        If temp_phone <> "" Then temp_phone = StndPhone(temp_phone)
                                        temp_match = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                                        'If debug = "Y" Then mydebuglog.Debug("  > getting temp_phone: " & temp_phone)
                                        'If debug = "Y" Then mydebuglog.Debug("  > getting temp_match: " & temp_match)

                                        ' If the org match codes match, find one other matching factor to declare a match
                                        'If debug = "Y" Then mydebuglog.Debug("  ... ORG_MATCH: " & ORG_MATCH)
                                        'If debug = "Y" Then mydebuglog.Debug("  ... (dr(4): " & Trim(CheckDBNull(dr(4), enumObjectType.StrType)))
                                        If Trim(ORG_MATCH) = Trim(CheckDBNull(dr(4), enumObjectType.StrType)) Then
                                            match_count = 1
                                            If debug = "Y" Then mydebuglog.Debug("    > Org match found - " & Trim(CheckDBNull(dr(1), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(2), enumObjectType.StrType)))
                                            If Trim(ADDR_MATCH) <> "" And Trim(ADDR_MATCH) = temp_match Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Addr match found")
                                                match_count = match_count + 2
                                            End If
                                            If Trim(ADDR_ID) <> "" And Trim(ADDR_ID) = Trim(CheckDBNull(dr(6), enumObjectType.StrType)) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Addr id match found")
                                                match_count = match_count + 2
                                            End If
                                            If Trim(ORG_PHONE) <> "" And Trim(ORG_PHONE) = temp_phone Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Org phone match found")
                                                match_count = match_count + 2
                                            End If
                                            If Trim(WORK_PH_NUM) <> "" And Trim(WORK_PH_NUM) = temp_phone Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Work phone match found")
                                                match_count = match_count + 2
                                            End If
                                            If debug = "Y" Then mydebuglog.Debug("     match_count: " & match_count.ToString)

                                            ' If confidence level reached then assume a match
                                            If match_count >= Confidence Then
                                                'If debug = "Y" Then mydebuglog.Debug("    > Org submatches found")
                                                high_count = match_count
                                                ' Save values retrieved from this match
                                                S_ORG_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                                S_NAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                                S_LOC = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                                If S_LOC <> "" Then
                                                    S_FULL_NAME = Trim(S_NAME) & " " & S_LOC
                                                Else
                                                    S_FULL_NAME = S_NAME
                                                End If
                                                S_ORG_MATCH = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                                S_ORG_PHONE = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                                S_INDUSTRY = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                                                S_ORG_NUM = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                                GoTo UpdActMatch
                                            End If
                                        Else
                                            If debug = "Y" Then mydebuglog.Debug("  > Org match NOT found: " & Trim(CheckDBNull(dr(1), enumObjectType.StrType)) & " " & Trim(CheckDBNull(dr(2), enumObjectType.StrType)))
                                        End If

                                        ' If org match codes DO NOT match, need multiple matching factors to declare a match
                                        If ORG_MATCH <> "" Then match_count = -1
                                        If debug = "Y" Then mydebuglog.Debug("    > Checking for non match code matches ")
                                        If Trim(NAME) <> "" Then
                                            If InStr(UCase(Trim(CheckDBNull(dr(1), enumObjectType.StrType))), Trim(UCase(NAME))) > 0 Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Name match found")
                                                match_count = match_count + 1
                                            Else
                                                match_count = match_count - 1
                                            End If
                                            If Trim(UCase(NAME)) = UCase(Trim(CheckDBNull(dr(15), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Name match found with location")
                                                match_count = match_count + 1
                                            End If
                                            If Trim(UCase(NAME)) = UCase(Trim(CheckDBNull(dr(16), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Name match found with city")
                                                match_count = match_count + 1
                                            End If
                                        End If
                                        If Trim(LOC) <> "" And Trim(UCase(LOC)) = UCase(Trim(CheckDBNull(dr(2), enumObjectType.StrType))) Then
                                            If debug = "Y" Then mydebuglog.Debug("     .. Location match found")
                                            match_count = match_count + 1
                                        End If
                                        If Trim(CITY) <> "" Then
                                            ntemp = Trim(UCase(NAME)) & " " & Trim(UCase(CITY))
                                            If ntemp = UCase(Trim(CheckDBNull(dr(1), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Name match found with city")
                                                match_count = match_count + 1
                                            End If
                                            If ntemp = UCase(Trim(CheckDBNull(dr(15), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Name match found with location")
                                                match_count = match_count + 1
                                            End If
                                            If ntemp = UCase(Trim(CheckDBNull(dr(16), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Name match found with city")
                                                match_count = match_count + 1
                                            End If
                                        End If
                                        If Trim(ADDR_MATCH) <> "" Then
                                            If Trim(ADDR_MATCH) = temp_match Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Address match code match found")
                                                match_count = match_count + 2
                                            End If
                                        Else
                                            If UCase(Trim(ADDR)) = UCase(Trim(CheckDBNull(dr(7), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Address match found")
                                                match_count = match_count + 1
                                            Else
                                                match_count = match_count - 1
                                            End If
                                            If UCase(Trim(CITY)) = UCase(Trim(CheckDBNull(dr(8), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. City match found")
                                                match_count = match_count + 1
                                            Else
                                                match_count = match_count - 1
                                            End If
                                            If UCase(Trim(ZIPCODE)) = UCase(Trim(CheckDBNull(dr(10), enumObjectType.StrType))) Then
                                                If debug = "Y" Then mydebuglog.Debug("     .. Zipcode match found")
                                                match_count = match_count + 1
                                            Else
                                                match_count = match_count - 1
                                            End If
                                        End If
                                        If Trim(ADDR_ID) <> "" And Trim(ADDR_ID) = Trim(CheckDBNull(dr(6), enumObjectType.StrType)) Then
                                            If debug = "Y" Then mydebuglog.Debug("     .. Address Id match found")
                                            match_count = match_count + 2
                                        End If
                                        If Trim(ORG_PHONE) <> "" And Trim(ORG_PHONE) = temp_phone Then
                                            If debug = "Y" Then mydebuglog.Debug("     .. Org phone match found")
                                            match_count = match_count + 2
                                        End If
                                        If Trim(WORK_PH_NUM) <> "" And Trim(WORK_PH_NUM) = temp_phone Then
                                            If debug = "Y" Then mydebuglog.Debug("     .. Work phone match found")
                                            match_count = match_count + 2
                                        End If
                                        If debug = "Y" Then mydebuglog.Debug("     match_count: " & match_count.ToString)

                                        ' Save the values from the highest match
                                        If match_count > high_count Then
                                            S_ORG_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                            S_NAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                            S_LOC = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                            If S_LOC <> "" Then
                                                S_FULL_NAME = Trim(S_NAME) & " " & S_LOC
                                            Else
                                                S_FULL_NAME = S_NAME
                                            End If
                                            S_ORG_MATCH = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                            S_ORG_PHONE = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                            S_INDUSTRY = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                                            S_ORG_NUM = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                            high_count = match_count
                                        End If
                                    End While
                                    ' If the high count exceeds the confidence, declare a match
                                    If high_count >= Confidence Then GoTo UpdActMatch
                                    If debug = "Y" Then mydebuglog.Debug("  > Exited record without finding a match. Partial matches found: " & match_count.ToString)
                                End If

                                ' Didn't find a match.. use one from the ORG_ID<>"" if available
                                If ORG_ID <> "" Then
                                    dr.Close()
                                    dr = cmd.ExecuteReader()
                                    While dr.Read()
                                        If ORG_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)) Then GoTo UpdActMatch
                                    End While
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading S_ORG_EXT: " & ex.ToString & vbCrLf
                            End Try
                            dr.Close()
                            If debug = "Y" Then mydebuglog.Debug("  > No account match found" & vbCrLf & "  =========================" & vbCrLf)
                            GoTo ExitMatch      ' No match found, just return what we standardized
UpdActMatch:
                            Try
                                If debug = "Y" Then
                                    mydebuglog.Debug(vbCrLf & "  > Full match found: " & S_ORG_ID & ". Score: " & high_count.ToString & vbCrLf)
                                    mydebuglog.Debug("   S_NAME: " & S_NAME)
                                    mydebuglog.Debug("   S_LOC: " & S_LOC)
                                    mydebuglog.Debug("   S_ORG_MATCH: " & S_ORG_MATCH)
                                    mydebuglog.Debug("   S_ORG_PHONE: " & S_ORG_PHONE)
                                    mydebuglog.Debug("   S_ORG_NUM: " & S_ORG_NUM)
                                    mydebuglog.Debug("  =========================" & vbCrLf)
                                End If
                                ORG_ID = S_ORG_ID
                                NAME = S_NAME
                                LOC = S_LOC
                                FULL_NAME = S_FULL_NAME
                                ORG_MATCH = S_ORG_MATCH
                                ORG_PHONE = S_ORG_PHONE
                                INDUSTRY = S_INDUSTRY
                                ORG_NUM = S_ORG_NUM
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading record: " & ex.ToString & vbCrLf
                            End Try
                        End If

                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error locating organization record. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                End If
            Else
                ' Get organization attributes
                SqlS = "SELECT DEDUP_TOKEN, MAIN_PH_NUM, X_ACCOUNT_NUM, PR_INDUST_ID, PR_ADDR_ID " & _
                    "FROM siebeldb.dbo.S_ORG_EXT WHERE ROW_ID='" & ORG_ID & "'"
                If debug = "Y" Then mydebuglog.Debug("  Retrieve organization attributes: " & SqlS)
                Try
                    'cmd.CommandText = SqlS
                    'dr = cmd.ExecuteReader()
                    '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                    cmd_ro.CommandText = SqlS
                    dr = cmd_ro.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                If Trim(ORG_MATCH) = "" Then ORG_MATCH = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                If Trim(ORG_PHONE) = "" Then ORG_PHONE = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                If Trim(ORG_NUM) = "" Then ORG_NUM = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                If Trim(INDUSTRY) = "" Then INDUSTRY = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                If INDUSTRY = "No Match Row Id" Then INDUSTRY = ""
                                If Trim(ADDR_ID) = "" Then ADDR_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                            Catch ex As Exception
                                errmsg = errmsg & "Error retrieving organization record: " & ex.ToString & vbCrLf
                            End Try
                        End While
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error retrieving organization record. " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try
                dr.Close()
                If debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & "  > Org Attributes found")
                    mydebuglog.Debug("   ORG_MATCH: " & ORG_MATCH)
                    mydebuglog.Debug("   ORG_PHONE: " & ORG_PHONE)
                    mydebuglog.Debug("   ORG_NUM: " & ORG_NUM)
                    mydebuglog.Debug("   INDUSTRY: " & INDUSTRY)
                    mydebuglog.Debug("   ADDR_ID: " & ADDR_ID)
                    mydebuglog.Debug("  =========================" & vbCrLf)
                End If
            End If

ExitMatch:
            ' ============================================
            ' Check address
            If ORG_ID <> "" And ADDR_MATCH <> "" And ADDR_ID = "" Then
                If ADDR <> "" Then
                    ' If address provided and address id is not known, clean address
                    Call CallCleanAddress(ADDR_ID, ORG_ID, CON_ID, "O", "N", _
                    ADDR, CITY, STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, CON_MATCH, WORK_PH_NUM, _
                    ORG_MATCH, ORG_PHONE, JURIS_ID, debug, mydebuglog, errmsg, results, database)
                Else
                    ' If address not provided and address id not known, match address
                    SqlS = "SELECT ROW_ID FROM siebeldb.dbo.S_ADDR_ORG " & _
                    "WHERE OU_ID='" & ORG_ID & "' AND X_MATCH_CD='" & ADDR_MATCH & "'"
                    If debug = "Y" Then mydebuglog.Debug("  Get address id: " & SqlS)
                    Try
                        'cmd.CommandText = SqlS
                        'dr = cmd.ExecuteReader()
                        '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                        cmd_ro.CommandText = SqlS
                        dr = cmd_ro.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    ADDR_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading adddress id: " & ex.ToString & vbCrLf
                                End Try
                            End While
                        End If
                        If debug = "Y" Then mydebuglog.Debug("    > ADDR_ID: " & ADDR_ID)
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error reading adddress id. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                End If
            End If

            ' ============================================
            ' Database operations
            ' Create record
            If database = "C" Then
                If ORG_ID = "" Then
                    ' Generate random organization id
                    SqlS = "SELECT RTRIM(CAST(MAX(X_ACCOUNT_NUM)+1 AS VARCHAR)) FROM siebeldb.dbo.S_ORG_EXT"
                    If debug = "Y" Then mydebuglog.Debug("  Get unique id: " & SqlS)
                    Try
                        'cmd.CommandText = SqlS
                        'dr = cmd.ExecuteReader()
                        '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                        cmd_ro.CommandText = SqlS
                        dr = cmd_ro.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    ORG_NUM = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading S_ORG_EXT: " & ex.ToString & vbCrLf
                                End Try
                            End While
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error locating org record. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                    If returnv > 0 Then
                        temp = ""
                        Try
                            temp = LoggingService.GenerateRecordId("S_ORG_EXT", "N", debug)
                        Catch ex As Exception
                            If debug = "Y" Then mydebuglog.Debug("  LoggingService Error: " & ex.ToString)
                        End Try
                        If temp <> "" Then ORG_ID = temp
                        If debug = "Y" Then mydebuglog.Debug("   > ORG_ID generated: " & ORG_ID)
                    End If

                    ' Verify Name Uniqueness (needed for index S_ORG_EXT_U1)
                    SqlS = "SELECT MAX(CAST(CONFLICT_ID AS INT))+1 FROM siebeldb.dbo.S_ORG_EXT " & _
                    "WHERE NAME='" & SqlString(NAME) & "' "
                    If LOC.Trim = "" Then
                        SqlS = SqlS & "AND (LOC='' OR LOC IS NULL)"
                    Else
                        SqlS = SqlS & "AND LOC='" & SqlString(LOC) & "'"
                    End If
                    If debug = "Y" Then mydebuglog.Debug("  Verifying name uniqueness: " & SqlS)
                    Try
                        'cmd.CommandText = SqlS
                        'dr = cmd.ExecuteReader()
                        '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                        cmd_ro.CommandText = SqlS
                        dr = cmd_ro.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    CONFLICT_ID = Trim(Str(CheckDBNull(dr(0), enumObjectType.IntType)))
                                Catch ex As Exception
                                    errmsg = errmsg & "Error locating S_ORG_EXT: " & ex.ToString & vbCrLf
                                End Try
                            End While
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error reading org record. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                    If CONFLICT_ID = "" Then CONFLICT_ID = "0"

InsertOrg:
                    ' Create org record and generate new id
                    If ORG_ID <> "" Then
                        SqlS = "INSERT INTO siebeldb.dbo.S_ORG_EXT " & _
                        "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM,MODIFICATION_NUM,CONFLICT_ID,BU_ID," & _
                        "DISA_CLEANSE_FLG,NAME,LOC,PROSPECT_FLG,PRTNR_FLG,ENTERPRISE_FLAG,LANG_ID,BASE_CURCY_CD," & _
                        "CREATOR_LOGIN,CUST_STAT_CD,DESC_TEXT,DISA_ALL_MAILS_FLG,FRGHT_TERMS_CD,MAIN_FAX_PH_NUM,MAIN_PH_NUM," & _
                        "PR_POSTN_ID,X_DATAFLEX_FLG,X_ACCOUNT_NUM,PR_ADDR_ID,PR_BL_ADDR_ID,PR_SHIP_ADDR_ID,PR_BL_PER_ID," & _
                        "PR_SHIP_PER_ID,DEDUP_TOKEN,X_MATCH_DT,PR_INDUST_ID) " & _
                        "SELECT TOP 1 '" & ORG_ID & "', getdate(), '0-1', getdate(), '0-1', 0, " & ORG_NUM & ", " & CONFLICT_ID & ", '0-R9NH', " & _
                        "'N', '" & SqlString(NAME) & "','" & SqlString(LOC) & "', 'Y', 'N', 'Y', 'ENU', 'USD', " & _
                        "'SADMIN', 'Prospect', 'CleanOrganization service', 'N', 'FOB', '', '" & ORG_PHONE & "', " & _
                        "'0-5220', 'N', MAX(X_ACCOUNT_NUM)+1, '" & ADDR_ID & "', '" & ADDR_ID & _
                        "', '" & ADDR_ID & "', '', '','" & ORG_MATCH & "',GETDATE(),'" & INDUSTRY & "' FROM siebeldb.dbo.S_ORG_EXT"
                    Else
                        SqlS = "INSERT INTO siebeldb.dbo.S_ORG_EXT " & _
                        "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM,MODIFICATION_NUM,CONFLICT_ID,BU_ID," & _
                        "DISA_CLEANSE_FLG,NAME,LOC,PROSPECT_FLG,PRTNR_FLG,ENTERPRISE_FLAG,LANG_ID,BASE_CURCY_CD," & _
                        "CREATOR_LOGIN,CUST_STAT_CD,DESC_TEXT,DISA_ALL_MAILS_FLG,FRGHT_TERMS_CD,MAIN_FAX_PH_NUM,MAIN_PH_NUM," & _
                        "PR_POSTN_ID,X_DATAFLEX_FLG,X_ACCOUNT_NUM,PR_ADDR_ID,PR_BL_ADDR_ID,PR_SHIP_ADDR_ID,PR_BL_PER_ID," & _
                        "PR_SHIP_PER_ID,DEDUP_TOKEN,X_MATCH_DT,PR_INDUST_ID) " & _
                        "SELECT TOP 1 'CLN'+LTRIM(RTRIM(CAST(MAX(X_ACCOUNT_NUM)+1 AS VARCHAR))), getdate(), '0-1', getdate(), '0-1', 0, " & ORG_NUM & ", " & CONFLICT_ID & ", '0-R9NH', " & _
                        "'N', '" & SqlString(NAME) & "','" & SqlString(LOC) & "', 'Y', 'N', 'Y', 'ENU', 'USD', " & _
                        "'SADMIN', 'Prospect', 'CleanOrganization service', 'N', 'FOB', '', '" & ORG_PHONE & "', " & _
                        "'0-5220', 'N', MAX(X_ACCOUNT_NUM)+1, '" & ADDR_ID & "', '" & ADDR_ID & _
                        "', '" & ADDR_ID & "', '', '','" & ORG_MATCH & "',GETDATE(),'" & INDUSTRY & "' FROM siebeldb.dbo.S_ORG_EXT"
                    End If
                    Try
                        temp = ExecQuery("Create", "Org record", cmd, SqlS, mydebuglog, debug)
                        NEW_ORG = "Y"
                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Insert account error: " & ex.ToString)
                        temp = "Insert Org record error"
                    End Try
                    errmsg = errmsg & temp

                    ' Locate the new ORG_ID and ORG_NUM if necessary
                    If ORG_ID = "" Or ORG_NUM = "" Then
                        SqlS = "SELECT ROW_ID, X_ACCOUNT_NUM " & _
                        "FROM siebeldb.dbo.S_ORG_EXT " & _
                        "WHERE MODIFICATION_NUM=" & ORG_NUM & " AND DEDUP_TOKEN='" & ORG_MATCH & "'"
                        If debug = "Y" Then mydebuglog.Debug("  Retrieving new row_id: " & SqlS)
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                            '2021-07-08: Rebecca: AG replication not fast enough to use the RO instance to check
                            'cmd_ro.CommandText = SqlS
                            'dr = cmd_ro.ExecuteReader()
                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        ORG_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                        ORG_NUM = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                    Catch ex As Exception
                                        errmsg = errmsg & "Error reading S_ORG_EXT: " & ex.ToString & vbCrLf
                                    End Try
                                End While
                            End If
                        Catch ex As Exception
                            errmsg = errmsg & vbCrLf & "Error reading org record. " & ex.ToString
                            results = "Failure"
                            GoTo CloseOut
                        End Try
                        dr.Close()
                        If debug = "Y" Then mydebuglog.Debug("   > ORG_ID generated: " & ORG_ID)
                    End If

                    ' Retry if failure
                    If ORG_ID = "" Then
                        ba_count = ba_count + 1
                        If ba_count < 3 Then
                            GoTo InsertOrg
                        Else
                            errmsg = errmsg & vbCrLf & "Unable to insert organization record. "
                            results = "Failure"
                        End If
                    End If

                    ' Create account position record
                    If ORG_ID <> "" Then
                        ' Create account position record
                        SqlS = "INSERT INTO siebeldb.dbo.S_ACCNT_POSTN " & _
                        "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM," & _
                        "CONFLICT_ID,ACCNT_NAME,OU_EXT_ID,POSITION_ID,ROW_STATUS) " & _
                        "SELECT TOP 1 '" & ORG_ID & "', getdate(), '0-1', getdate(), '0-1', 0, " & _
                        CONFLICT_ID & ", '" & SqlString(NAME) & "', '" & ORG_ID & "', '0-5220', 'N' " & _
                        "FROM siebeldb.dbo.S_ACCNT_POSTN " & _
                        "WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_ACCNT_POSTN WHERE ROW_ID='" & ORG_ID & "')"
                        temp = ExecQuery("Create", "Org position record", cmd, SqlS, mydebuglog, debug)
                        errmsg = errmsg & temp

                        ' If Industry code present, create account industry record
                        If INDUSTRY <> "" And INDUSTRY <> "<Please+select>" Then
                            SqlS = "INSERT siebeldb.dbo.S_ORG_INDUST " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,CONFLICT_ID,INDUST_ID,OU_ID) " & _
                            "SELECT TOP 1 '" & ORG_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0,0,'" & SqlString(INDUSTRY) & "','" & ORG_ID & "' " & _
                            "FROM siebeldb.dbo.S_ORG_INDUST " & _
                            "WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.S_ORG_INDUST WHERE ROW_ID='" & ORG_ID & "')"
                            temp = ExecQuery("Create", "Org industry record", cmd, SqlS, mydebuglog, debug)
                            errmsg = errmsg & temp
                        End If
                    End If
                Else
                    If debug = "Y" Then mydebuglog.Debug("Org record already exists.. did not create a new one")
                    results = "Success"
                End If

                ' ============================================
                ' Check to see if the organization match code needs to be computed
                If ORG_MATCH = "" Then
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "ORG_MATCH IS EMPTY - COMPUTING")

                    ' Call StandardizeOrganization to update the record
                    wp = "<Organizations><Organization>"
                    wp = wp & "<Debug>N</Debug>"
                    wp = wp & "<Database>U</Database>"
                    wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
                    wp = wp & "<Name>" & HttpUtility.UrlEncode(NAME) & "</Name>"
                    wp = wp & "<Loc>" & HttpUtility.UrlEncode(LOC) & "</Loc>"
                    wp = wp & "<FullName>" & HttpUtility.UrlEncode(FULL_NAME) & "</FullName>"
                    wp = wp & "</Organization></Organizations>"
                    Try
                        If debug = "Y" Then mydebuglog.Debug("  sXML: " & wp)
                        rDoc = StandardizeOrganization(wp)
                        If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
                        rNodeList = rDoc.SelectNodes("//Organization")
                        For j = 0 To rNodeList.Count - 1
                            Try
                                If debug = "Y" Then mydebuglog.Debug("  Processing node: " & j.ToString)
                                NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("Name", rNodeList.Item(j))))
                                LOC = HttpUtility.UrlDecode(Trim(GetNodeValue("Loc", rNodeList.Item(j))))
                                FULL_NAME = HttpUtility.UrlDecode(GetNodeValue("FullName", rNodeList.Item(j)))
                                If FULL_NAME = "" Then
                                    If LOC <> "" Then
                                        FULL_NAME = Trim(NAME) & " " & LOC
                                    Else
                                        FULL_NAME = NAME
                                    End If
                                End If
                                ORG_MATCH = HttpUtility.UrlDecode(GetNodeValue("MatchCode", rNodeList.Item(j)))
                                ORG_ID = HttpUtility.UrlDecode(Trim(GetNodeValue("OrgId", rNodeList.Item(j))))
                            Catch ex As Exception
                                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                                results = "Failure"
                                GoTo CloseOut2
                            End Try
                        Next
                        If debug = "Y" Then mydebuglog.Debug("  Standardized: " & results)
                        If results <> "Success" Then GoTo CloseOut

                    Catch ex As Exception
                        If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
                    End Try
                    If debug = "Y" Then
                        mydebuglog.Debug(vbCrLf & "  StandardizeOrganization Results====")
                        mydebuglog.Debug("  Name: " & NAME)
                        mydebuglog.Debug("  Loc: " & LOC)
                        mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
                        mydebuglog.Debug("  FullName: " & FULL_NAME)
                        mydebuglog.Debug("  =========================" & vbCrLf)
                    End If
                End If

                ' ============================================
                ' Clean address if applicable
                If ORG_ID <> "" Then
                    If ADDR <> "" And ADDR_ID = "" Then
                        ' If account known then look for that address
                        Call CallCleanAddress(ADDR_ID, ORG_ID, CON_ID, "O", "N", _
                        ADDR, CITY, STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, CON_MATCH, WORK_PH_NUM, _
                        ORG_MATCH, ORG_PHONE, JURIS_ID, debug, mydebuglog, errmsg, results, database)
                    End If
                End If
            End If

            '-----
            ' Update record
            If database = "U" Then
                If ORG_ID <> "" Then
                    SqlS = "UPDATE siebeldb.dbo.S_ORG_EXT SET LAST_UPD=GETDATE()"
                    If NAME <> "" Then SqlS = SqlS & ",NAME='" & SqlString(NAME) & "'"
                    If LOC <> "" Then SqlS = SqlS & ",LOC='" & SqlString(LOC) & "'"
                    If ORG_MATCH <> "" Then SqlS = SqlS & ",DEDUP_TOKEN='" & SqlString(ORG_MATCH) & "',X_MATCH_DT=GETDATE()"
                    'If NAME <> "" Then SqlS = SqlS & ",NAME='" & SqlString(NAME) & "'"
                    If ORG_PHONE <> "" Then SqlS = SqlS & ",MAIN_PH_NUM='" & SqlString(ORG_PHONE) & "'"
                    If INDUSTRY <> "" Then SqlS = SqlS & ",PR_INDUST_ID='" & SqlString(INDUSTRY) & "'"
                    If ADDR_ID <> "" Then SqlS = SqlS & ",PR_ADDR_ID='" & ADDR_ID & "'"
                    SqlS = SqlS & " WHERE ROW_ID='" & ORG_ID & "'"
                    temp = ExecQuery("Update", "Org record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                Else
                    errmsg = errmsg & vbCrLf & "Unable to update org because id not found. "
                    results = "Failure"
                End If
            End If

            ' ============================================
            ' Check to see if the organization match code needs to be computed
            If ORG_ID <> "" And ORG_MATCH = "" Then
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "ORG_MATCH IS EMPTY - COMPUTING")

                ' Call StandardizeOrganization to update the record
                wp = "<Organizations><Organization>"
                wp = wp & "<Debug>N</Debug>"
                wp = wp & "<Database>U</Database>"
                wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
                wp = wp & "<Name>" & HttpUtility.UrlEncode(NAME) & "</Name>"
                wp = wp & "<Loc>" & HttpUtility.UrlEncode(LOC) & "</Loc>"
                wp = wp & "<FullName>" & HttpUtility.UrlEncode(FULL_NAME) & "</FullName>"
                wp = wp & "</Organization></Organizations>"
                Try
                    If debug = "Y" Then mydebuglog.Debug("  sXML: " & wp)
                    rDoc = StandardizeOrganization(wp)
                    If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
                    rNodeList = rDoc.SelectNodes("//Organization")
                    For j = 0 To rNodeList.Count - 1
                        Try
                            If debug = "Y" Then mydebuglog.Debug("  Processing node: " & j.ToString)
                            NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("Name", rNodeList.Item(j))))
                            LOC = HttpUtility.UrlDecode(Trim(GetNodeValue("Loc", rNodeList.Item(j))))
                            FULL_NAME = HttpUtility.UrlDecode(GetNodeValue("FullName", rNodeList.Item(j)))
                            If FULL_NAME = "" Then
                                If LOC <> "" Then
                                    FULL_NAME = Trim(NAME) & " " & LOC
                                Else
                                    FULL_NAME = NAME
                                End If
                            End If
                            ORG_MATCH = HttpUtility.UrlDecode(GetNodeValue("MatchCode", rNodeList.Item(j)))
                            ORG_ID = HttpUtility.UrlDecode(Trim(GetNodeValue("OrgId", rNodeList.Item(j))))
                        Catch ex As Exception
                            errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                            results = "Failure"
                            GoTo CloseOut2
                        End Try
                    Next
                    If debug = "Y" Then mydebuglog.Debug("  Standardized: " & results)
                    If results <> "Success" Then GoTo CloseOut

                Catch ex As Exception
                    If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
                End Try
                If debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & "  StandardizeOrganization Results====")
                    mydebuglog.Debug("  Name: " & NAME)
                    mydebuglog.Debug("  Loc: " & LOC)
                    mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
                    mydebuglog.Debug("  FullName: " & FULL_NAME)
                    mydebuglog.Debug("  =========================" & vbCrLf)
                End If
            End If

        Next

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Close()
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf

        End Try

CloseOut2:
        ' ============================================
        ' Return the cleaned/deduped information as an XML document:
        '   <Organization>
        '       <OrgId></OrgId>         - The Id of an existing organization, if applicable
        '       <Name></Name>           - Name of organization
        '       <Loc></Loc>             - Location of organization
        '       <FullName></Fullname>   - Full name of organization
        '       <OrgMatch></OrgMatch>   - Organization match code
        '       <OrgPhone></OrgPhone>   - Organization main phone number
        '       <OrgNum></OrgNum>       - Organization number
        '       <AddrMatch></AddrMatch> - Address match code, if applicable
        '       <AddrId></AddrId>       - Address Id of a related address, if applicable
        '	    <Address></Address>		- Street Address
        '	    <City></City>			- City
        '	    <State></State>			- State or province
        '	    <County></County>		- County or region
        '	    <Zipcode></Zipcode>		- Zipcode or postal code
        '	    <Country></Country>		- Country
        '       <WorkPhone></WorkPhone> - Work phone
        '       <Industry></Industry>   - Industry
        '       <NewOrg></NewOrg>       - Organization created flag
        '   </Organization>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("Organization")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            If debug <> "T" And ORG_MATCH <> "" Then
                AddXMLChild(odoc, resultsRoot, "OrgId", IIf(ORG_ID = "", " ", HttpUtility.UrlEncode(ORG_ID)))
                AddXMLChild(odoc, resultsRoot, "Name", IIf(NAME = "", " ", HttpUtility.UrlEncode(NAME)))
                AddXMLChild(odoc, resultsRoot, "Loc", IIf(LOC = "", " ", HttpUtility.UrlEncode(LOC)))
                AddXMLChild(odoc, resultsRoot, "FullName", IIf(FULL_NAME = "", " ", HttpUtility.UrlEncode(FULL_NAME)))
                AddXMLChild(odoc, resultsRoot, "OrgMatch", IIf(ORG_MATCH = "", " ", HttpUtility.UrlEncode(ORG_MATCH)))
                AddXMLChild(odoc, resultsRoot, "OrgPhone", IIf(ORG_PHONE = "", " ", HttpUtility.UrlEncode(ORG_PHONE)))
                AddXMLChild(odoc, resultsRoot, "OrgNum", IIf(ORG_NUM = "", " ", HttpUtility.UrlEncode(ORG_NUM)))
                AddXMLChild(odoc, resultsRoot, "AddrMatch", IIf(ADDR_MATCH = "", " ", HttpUtility.UrlEncode(ADDR_MATCH)))
                AddXMLChild(odoc, resultsRoot, "AddrId", IIf(ADDR_ID = "", " ", HttpUtility.UrlEncode(ADDR_ID)))
                AddXMLChild(odoc, resultsRoot, "WorkPhone", IIf(WORK_PH_NUM = "", " ", HttpUtility.UrlEncode(WORK_PH_NUM)))
                AddXMLChild(odoc, resultsRoot, "Industry", IIf(INDUSTRY = "", " ", HttpUtility.UrlEncode(INDUSTRY)))
                AddXMLChild(odoc, resultsRoot, "Address", IIf(ADDR = "", " ", HttpUtility.UrlEncode(ADDR)))
                AddXMLChild(odoc, resultsRoot, "City", IIf(CITY = "", " ", HttpUtility.UrlEncode(CITY)))
                AddXMLChild(odoc, resultsRoot, "County", IIf(COUNTY = "", " ", HttpUtility.UrlEncode(COUNTY)))
                AddXMLChild(odoc, resultsRoot, "State", IIf(STATE = "", " ", HttpUtility.UrlEncode(STATE)))
                AddXMLChild(odoc, resultsRoot, "Zipcode", IIf(ZIPCODE = "", " ", HttpUtility.UrlEncode(ZIPCODE)))
                AddXMLChild(odoc, resultsRoot, "Country", IIf(COUNTRY = "", " ", HttpUtility.UrlEncode(COUNTRY)))
                AddXMLChild(odoc, resultsRoot, "NewOrg", IIf(NEW_ORG = "", "N", HttpUtility.UrlEncode(NEW_ORG)))
            Else
                If ORG_MATCH <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        If debug = "Y" Then
            Dim sw = New StringWriter
            Dim xw = New XmlTextWriter(sw)
            odoc.WriteTo(xw)
            mydebuglog.Debug(vbCrLf & "Generated XML output: " & sw.ToString() & vbCrLf)
            sw = Nothing
            xw = Nothing
        End If

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("CleanOrganization : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("CleanOrganization : Results: " & results & " for " & NAME & " and ID " & ORG_ID & " generated matchcode " & ORG_MATCH)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("Results: " & results & " for " & NAME & " and ID " & ORG_ID & " generated matchcode " & ORG_MATCH)
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended at " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "101"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc

    End Function

    <WebMethod(Description:="Matches and verifies address records")> _
    Public Function CleanAddress(ByVal sXML As String) As XmlDocument
        ' This function attempts to match the address supplied to other records. If a match is
        ' found it returns the the matching record.  It will also optionally update the matching
        ' record with changes from the supplied address.

        ' The input parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <AddressList>
        '       <AddressRec>
        '           <Debug>                 - A flag to indicate the service is to run in Debug mode or not
        '                                       "Y"  - Yes for debug mode on.. logging on
        '                                       "N"  - No for debug mode off.. logging off
        '                                       "T"  - Test mode on.. logging off
        '           <Database>              - "C" create S_ADDR_ORG record(s), "U" update record, "X" match only
        '           <Confidence>            - The number of factors that need to match for a match
        '           <AddrId>                - The Id of an existing address, if applicable
        '           <Type>                  - "O"rganization or "P"ersonal address
        '           <GeoCode>               - Geocode address flag ("Y" or "N") - optional
        '           <JurisId>               - Jurisdiction Id
        '           <Address>               - Street address
        '           <City>                  - City
        '           <State>                 - State
        '           <County>                - County
        '           <Zipcode>               - Zipcode
        '           <Country>               - Country
        '           <ConId></ConId>         - The Id of an existing contact, if known
        '           <ConMatch></ConMatch>   - Contact match code
        '           <WorkPhone></WorkPhone> - Contact work phone number
        '           <OrgId></OrgId>         - The Id of an existing organization, if known
        '           <OrgMatch></OrgMatch>   - Organization match code
        '           <OrgPhone></OrgPhone>   - Organization main phone number
        '       </AddressRec>
        '   </AddressList>

        ' web.config Parameters used:
        '   hcidb        - connection string to siebeldb database

        ' Generic variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim mypath, debug, ddebug, errmsg, logging, wp As String

        ' Database declarations
        Dim con As SqlConnection, con_ro As SqlConnection
        Dim cmd As SqlCommand, cmd_ro As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String, ConnS_ro As String
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("CADebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service
        Dim BasicService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim ADDR, CITY, STATE, ZIPCODE, COUNTY, COUNTRY, GEOCODE, ADDR_TYPE, ADDR_ID, ADDR_MATCH As String
        Dim temp, database, LAT, LON, ORG_ID, CON_ID, JURIS_ID, ORG_MATCH, ORG_PHONE, CON_MATCH As String
        Dim S_ADDR, S_CITY, S_STATE, S_ZIPCODE, S_COUNTY, S_COUNTRY, S_GEOCODE, S_ADDR_TYPE, S_ADDR_ID, S_ADDR_MATCH As String
        Dim S_LAT, S_LON, S_ORG_ID, S_CON_ID, S_JURIS_ID, S_ORG_MATCH, S_ORG_PHONE, S_CON_MATCH As String
        Dim A_ADDR, A_CITY, A_STATE, A_ZIPCODE, A_COUNTY, A_COUNTRY, A_ADDR_MATCH As String
        Dim A_LAT, A_LON, A_JURIS_ID As String
        Dim O_ADDR As String
        Dim WORK_PH_NUM, DELIVERABLE, temp_match, temp_phone, temp_addr, temp_city, temp_state, temp_zip As String
        Dim match_count, high_count, Confidence, ba_count, MatchNew_count As Integer
        Dim MatchNew As Boolean

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        ADDR_ID = ""
        ADDR_TYPE = "O"
        GEOCODE = "N"
        ADDR = "1101 Wilson Suite 1700"
        O_ADDR = ""
        CITY = "Arlington"
        STATE = "VA"
        COUNTY = ""
        ZIPCODE = ""
        COUNTRY = ""
        ADDR_MATCH = "ZZ0Z$LW4PZI00$$&YWPF~PV&HHH0$$"
        ORG_ID = ""
        CON_ID = ""
        CON_MATCH = ""
        JURIS_ID = ""
        WORK_PH_NUM = ""
        ORG_MATCH = ""
        ORG_PHONE = "800-438-8477"
        DELIVERABLE = ""
        temp_match = ""
        temp_phone = ""
        temp_addr = ""
        temp_city = ""
        temp_state = ""
        temp_zip = ""
        LAT = ""
        LON = ""
        temp = ""
        database = ""
        SqlS = ""
        returnv = 0
        wp = ""
        Confidence = 3
        match_count = 0
        ba_count = 0
        MatchNew = False
        MatchNew_count = 0

        S_ADDR = ""
        S_CITY = ""
        S_STATE = ""
        S_ZIPCODE = ""
        S_COUNTY = ""
        S_COUNTRY = ""
        S_GEOCODE = ""
        S_ADDR_TYPE = ""
        S_ADDR_ID = ""
        S_ADDR_MATCH = ""
        S_LAT = ""
        S_LON = ""
        S_ORG_ID = ""
        S_CON_ID = ""
        S_JURIS_ID = ""
        S_ORG_MATCH = ""
        S_ORG_PHONE = ""
        S_CON_MATCH = ""

        A_ADDR = ""
        A_CITY = ""
        A_STATE = ""
        A_ZIPCODE = ""
        A_COUNTY = ""
        A_COUNTRY = ""
        A_ADDR_MATCH = ""
        A_LAT = ""
        A_LON = ""
        A_JURIS_ID = ""
        high_count = 0
        ConnS = ""
        ConnS_ro = ""

        ' ============================================
        ' Check parameters
        debug = "N"
        ddebug = "N"
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        HttpUtility.UrlDecode(sXML)
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//AddressList/AddressRec")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        debug = UCase(debug)

        ' ============================================
        ' Get system defaults
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb").ConnectionString
            If ConnS = "" Then ConnS = "server="
            ConnS_ro = System.Configuration.ConfigurationManager.ConnectionStrings("hcidb_ro").ConnectionString
            If ConnS_ro = "" Then ConnS_ro = "server="
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("CleanAddress_debug").ToUpper()
            If temp = "Y" And debug <> "T" Then debug = temp
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("CleanAddress_detailed_debug").ToUpper()
            If temp = "Y" Then ddebug = temp
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\clean_address_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\CleanAddress.log"
            Try
                log4net.GlobalContext.Properties("CALogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' ============================================
        ' Open database connection 
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        ' ============================================; 2020-05-18; Ren Hou; Added for read-only per Chris;
        ' Open read-only database connection 
        errmsg = OpenDBConnection(ConnS_ro, con_ro, cmd_ro)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut2
            End Try
        Next
        For i = 0 To oNodeList.Count - 1
            errmsg = ""
            ' ============================================
            ' Extract data from parameter string
            If debug <> "T" Then
                ADDR_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("AddrId", oNodeList.Item(i)))))
                ADDR_ID = Trim(KeySpace(ADDR_ID))
                CON_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("ConId", oNodeList.Item(i)))))
                CON_ID = Trim(KeySpace(CON_ID))
                CON_MATCH = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("ConMatch", oNodeList.Item(i)))))
                WORK_PH_NUM = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("WorkPhone", oNodeList.Item(i)))))
                If WORK_PH_NUM <> "" Then WORK_PH_NUM = StndPhone(WORK_PH_NUM)
                ORG_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("OrgId", oNodeList.Item(i)))))
                ORG_ID = Trim(KeySpace(ORG_ID))
                ORG_MATCH = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("OrgMatch", oNodeList.Item(i)))))
                ORG_PHONE = Trim(HttpUtility.UrlDecode(GetNodeValue("OrgPhone", oNodeList.Item(i))))
                If ORG_PHONE <> "" Then ORG_PHONE = StndPhone(ORG_PHONE)
                ADDR_TYPE = Left(Trim(GetNodeValue("Type", oNodeList.Item(i))), 1)
                If ADDR_TYPE = "" And (ORG_ID <> "" Or ORG_MATCH <> "" Or ORG_PHONE <> "") Then ADDR_TYPE = "O"
                If ADDR_TYPE = "" And CON_ID <> "" Then ADDR_TYPE = "P"
                If ADDR_TYPE = "B" Then ADDR_TYPE = "O"
                GEOCODE = Trim(GetNodeValue("GeoCode", oNodeList.Item(i)))
                If GEOCODE <> "Y" Then GEOCODE = "N"
                JURIS_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("JurisId", oNodeList.Item(i)))))
                JURIS_ID = Trim(KeySpace(JURIS_ID))
                ADDR = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("Address", oNodeList.Item(i)))))
                ADDR = Trim(CleanString(ADDR))
                O_ADDR = Trim(UCase(ADDR))
                CITY = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("City", oNodeList.Item(i)))))
                CITY = CleanString(CITY)
                STATE = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("State", oNodeList.Item(i)))))
                STATE = Trim(CleanString(STATE))
                COUNTY = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("County", oNodeList.Item(i)))))
                COUNTY = Trim(CleanString(COUNTY))
                ZIPCODE = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("Zipcode", oNodeList.Item(i)))))
                ZIPCODE = Trim(CleanString(ZIPCODE))
                COUNTRY = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("Country", oNodeList.Item(i)))))
                COUNTRY = Trim(CleanString(COUNTRY))
                database = Left(GetNodeValue("Database", oNodeList.Item(i)), 1)
                temp = Trim(GetNodeValue("Confidence", oNodeList.Item(i)))
                If temp <> "" And IsNumeric(temp) Then Confidence = Int(temp)
                If Confidence < 5 Then Confidence = 5
            End If
            If debug = "Y" Then
                mydebuglog.Debug("INPUTS------" & vbCrLf & "  ADDR_ID: " & ADDR_ID)
                mydebuglog.Debug("  ORG_ID: " & ORG_ID)
                mydebuglog.Debug("  ORG_MATCH: " & ORG_MATCH)
                mydebuglog.Debug("  ORG_PHONE: " & ORG_PHONE)
                mydebuglog.Debug("  CON_ID: " & CON_ID)
                mydebuglog.Debug("  CON_MATCH: " & CON_MATCH)
                mydebuglog.Debug("  WORK_PH_NUM: " & WORK_PH_NUM)
                mydebuglog.Debug("  ADDR_TYPE: " & ADDR_TYPE)
                mydebuglog.Debug("  GEOCODE: " & GEOCODE)
                mydebuglog.Debug("  JURIS_ID: " & JURIS_ID)
                mydebuglog.Debug("  ADDR: " & ADDR)
                mydebuglog.Debug("  CITY: " & CITY)
                mydebuglog.Debug("  STATE: " & STATE)
                mydebuglog.Debug("  COUNTY: " & COUNTY)
                mydebuglog.Debug("  ZIPCODE: " & ZIPCODE)
                mydebuglog.Debug("  COUNTRY: " & COUNTRY)
                mydebuglog.Debug("  Confidence: " & Confidence)
                mydebuglog.Debug("  database: " & database & vbCrLf & "------------")
            End If

            ' ============================================
            ' Call StandardizeAddress to update the record
            ' Supress (pass "N" instead of "Y") geocoding in CallStandardizeAddress sub due to new Asynch geocoding; Ren Hou; 2018-11-01;
            Call CallStandardizeAddress(ADDR_ID, ORG_ID, CON_ID, ADDR_TYPE, "N", ADDR, CITY, _
              STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, JURIS_ID, DELIVERABLE, debug, mydebuglog, errmsg, results, "X")
            If results = "Failure" Then GoTo CloseOut2
            A_ADDR = ADDR
            A_CITY = CITY
            A_STATE = STATE
            A_ZIPCODE = ZIPCODE
            A_COUNTY = COUNTY
            A_COUNTRY = COUNTRY
            A_ADDR_MATCH = ADDR_MATCH
            A_LAT = LAT
            A_LON = LON
            A_JURIS_ID = JURIS_ID

            ' ============================================
            ' If incomplete address, exit
            If ADDR = "" Or CITY = "" Or STATE = "" Then GoTo ExitMatch

            ' Adjust confidence as applicable
            If ORG_ID <> "" Then Confidence = Confidence + 1

            ' ============================================
            ' Locate match
            If debug = "Y" Then mydebuglog.Debug("  Get Match Candidates: '" & ORG_ID & "' for address type '" & ADDR_TYPE & "' with id of '" & ADDR_ID & "'")
            If ADDR_ID = "" Then
                Select Case ADDR_TYPE
                    Case "P"
                        ' ----
                        ' Personal Address
                        returnv = 0
                        SqlS = "SELECT A.ROW_ID,(SELECT CASE WHEN A.ADDR IS NULL THEN '' ELSE A.ADDR END)," & _
                        "(SELECT CASE WHEN A.CITY IS NULL THEN '' ELSE A.CITY END)," & _
                        "(SELECT CASE WHEN A.STATE IS NULL THEN '' ELSE A.STATE END)," & _
                        "(SELECT CASE WHEN A.ZIPCODE IS NULL THEN '' ELSE A.ZIPCODE END)," & _
                        "(SELECT CASE WHEN A.COUNTRY IS NULL THEN '' ELSE A.COUNTRY END)," & _
                        "(SELECT CASE WHEN A.COUNTY IS NULL THEN '' ELSE A.COUNTY END)," & _
                        "(SELECT CASE WHEN A.X_MATCH_CD IS NULL THEN '' ELSE A.X_MATCH_CD END)," & _
                        "(SELECT CASE WHEN A.PER_ID IS NULL THEN '' ELSE A.PER_ID END)," & _
                        "(SELECT CASE WHEN A.X_LAT IS NULL THEN '0' ELSE CAST(A.X_LAT AS VARCHAR) END)," & _
                        "(SELECT CASE WHEN A.X_LONG IS NULL THEN '0' ELSE CAST(A.X_LONG AS VARCHAR) END),"
                        If CON_ID <> "" Or CON_MATCH <> "" Or WORK_PH_NUM <> "" Then
                            SqlS = SqlS & "(SELECT CASE WHEN C.WORK_PH_NUM IS NULL THEN '' ELSE C.WORK_PH_NUM END) "
                        Else
                            SqlS = SqlS & "'' "
                        End If
                        SqlS = SqlS & "FROM siebeldb.dbo.S_ADDR_PER A "
                        If ADDR_MATCH <> "" And ADDR_MATCH <> "$$$$$$$$$$$$$$$" Then
                            If CON_ID <> "" Or CON_MATCH <> "" Or WORK_PH_NUM <> "" Then
                                SqlS = SqlS & "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=A.PER_ID "
                                SqlS = SqlS & "WHERE (A.X_MATCH_CD='" & ADDR_MATCH & "') AND ("
                                If CON_ID <> "" Then SqlS = SqlS & "C.ROW_ID='" & CON_ID & "' OR "
                                If CON_MATCH <> "" Then SqlS = SqlS & "C.X_MATCH_CD='" & CON_MATCH & "' OR "
                                If WORK_PH_NUM <> "" Then SqlS = SqlS & "C.WORK_PH_NUM='" & WORK_PH_NUM & "'"
                                If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                                SqlS = SqlS & ")"
                                If Right(SqlS, 6) = "AND ()" Then SqlS = Left(SqlS, Len(SqlS) - 6)
                            Else
                                SqlS = SqlS & "WHERE (A.X_MATCH_CD='" & ADDR_MATCH & "') AND ("
                                SqlS = SqlS & "(UPPER(A.ADDR)='" & UCase(ADDR) & "' OR " & _
                                    "UPPER(A.CITY)='" & UCase(CITY) & "' OR " & _
                                    "UPPER(A.STATE)='" & UCase(STATE) & "')"
                            End If
                        Else
                            If CON_ID <> "" Or CON_MATCH <> "" Or WORK_PH_NUM <> "" Then
                                SqlS = SqlS & "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=A.PER_ID " & _
                                    "WHERE "
                                If CON_ID <> "" Then SqlS = SqlS & "C.ROW_ID='" & CON_ID & "' OR "
                                If CON_MATCH <> "" Then SqlS = SqlS & "C.X_MATCH_CD='" & CON_MATCH & "' OR "
                                If WORK_PH_NUM <> "" Then SqlS = SqlS & "C.WORK_PH_NUM='" & WORK_PH_NUM & "'"
                                If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                            Else
                                SqlS = SqlS & "WHERE (UPPER(A.ADDR)='" & UCase(ADDR) & "' AND " & _
                                    "UPPER(A.CITY)='" & UCase(CITY) & "' AND " & _
                                    "UPPER(A.STATE)='" & UCase(STATE) & "')"
                            End If
                        End If
                        If debug = "Y" Then mydebuglog.Debug("  =========================" & vbCrLf & "  Checking matches: " & SqlS)
                        'cmd.CommandText = SqlS
                        'dr = cmd.ExecuteReader()
                        '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                        cmd_ro.CommandText = SqlS
                        dr = cmd_ro.ExecuteReader()

                        If Not dr Is Nothing Then
                            ' -----
                            ' Evaluate all of the records found and locate the best match
                            While dr.Read()
                                Try
                                    returnv = returnv + 1
                                    ' Get variables
                                    match_count = 0
                                    If debug = "Y" Then mydebuglog.Debug("  > Checking address match: " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)))
                                    temp_match = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                    temp_phone = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                    If temp_phone <> "" Then temp_phone = StndPhone(temp_phone)

                                    ' Check sub-matches based on match code, or not depending on match
                                    Try
                                        If ADDR_MATCH <> "" And ADDR_MATCH <> "$$$$$$$$$$$$$$$" Then
                                            If Trim(ADDR_MATCH) = temp_match Then match_count = match_count + 1
                                        Else
                                            If debug = "Y" Then mydebuglog.Debug("  > Address match NOT found: " & temp_match)
                                        End If
                                        If debug = "Y" Then mydebuglog.Debug("  > WORK_PH_NUM: " & WORK_PH_NUM & " -- " & temp_phone)
                                        If Trim(WORK_PH_NUM) <> "" And Trim(WORK_PH_NUM) = temp_phone Then match_count = match_count + 1

                                        If debug = "Y" Then mydebuglog.Debug("  > ADDR_ID: " & ADDR_ID & " -- " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)))
                                        If Trim(ADDR_ID) <> "" And Trim(ADDR_ID) = Trim(CheckDBNull(dr(0), enumObjectType.StrType)) Then match_count = match_count + 1

                                        If debug = "Y" Then mydebuglog.Debug("  > CON_ID: " & CON_ID & " -- " & Trim(CheckDBNull(dr(8), enumObjectType.StrType)))
                                        If Trim(CON_ID) <> "" And Trim(CON_ID) = Trim(CheckDBNull(dr(8), enumObjectType.StrType)) Then match_count = match_count + 1

                                        If debug = "Y" Then mydebuglog.Debug("  > ADDR: " & ADDR & " -- " & Trim(CheckDBNull(dr(1), enumObjectType.StrType)))
                                        If Trim(ADDR) <> "" And Trim(ADDR) = Trim(CheckDBNull(dr(1), enumObjectType.StrType)) Then match_count = match_count + 1
                                        If Trim(ADDR) <> "" And Trim(ADDR) <> Trim(CheckDBNull(dr(1), enumObjectType.StrType)) Then match_count = match_count - 1

                                        If debug = "Y" Then mydebuglog.Debug("  > CITY: " & CITY & " -- " & Trim(CheckDBNull(dr(2), enumObjectType.StrType)))
                                        If Trim(CITY) <> "" And Trim(CITY) = Trim(CheckDBNull(dr(2), enumObjectType.StrType)) Then match_count = match_count + 1
                                        If Trim(CITY) <> "" And Trim(CITY) <> Trim(CheckDBNull(dr(2), enumObjectType.StrType)) Then match_count = match_count - 1

                                        If debug = "Y" Then mydebuglog.Debug("  > STATE: " & STATE & " -- " & Trim(CheckDBNull(dr(3), enumObjectType.StrType)))
                                        If Trim(STATE) <> "" And Trim(STATE) = Trim(CheckDBNull(dr(3), enumObjectType.StrType)) Then match_count = match_count + 1
                                        If Trim(STATE) <> "" And Trim(STATE) <> Trim(CheckDBNull(dr(3), enumObjectType.StrType)) Then match_count = match_count - 1

                                        If debug = "Y" Then mydebuglog.Debug("  > ZIPCODE: " & ZIPCODE & " -- " & Trim(CheckDBNull(dr(4), enumObjectType.StrType)))
                                        If Trim(ZIPCODE) <> "" And Trim(ZIPCODE) = Trim(CheckDBNull(dr(4), enumObjectType.StrType)) Then match_count = match_count + 1
                                        If Trim(ZIPCODE) <> "" And Trim(ZIPCODE) <> Trim(CheckDBNull(dr(4), enumObjectType.StrType)) Then match_count = match_count - 1

                                        If debug = "Y" Then mydebuglog.Debug("  > COUNTRY: " & COUNTRY & " -- " & Trim(CheckDBNull(dr(5), enumObjectType.StrType)))
                                        If Trim(COUNTRY) <> "" And Trim(COUNTRY) = Trim(CheckDBNull(dr(5), enumObjectType.StrType)) Then match_count = match_count + 1

                                        If debug = "Y" Then mydebuglog.Debug("  > COUNTY: " & COUNTY & " -- " & Trim(CheckDBNull(dr(6), enumObjectType.StrType)))
                                        If Trim(COUNTY) <> "" And Trim(COUNTY) = Trim(CheckDBNull(dr(6), enumObjectType.StrType)) Then match_count = match_count + 1

                                        If debug = "Y" Then mydebuglog.Debug("  > LAT: " & LAT & " -- " & Trim(CheckDBNull(dr(9), enumObjectType.StrType)))
                                        If Trim(LAT) <> "" And IsNumeric(LAT) And Trim(LAT) = CheckDBNull(dr(9), enumObjectType.StrType) Then match_count = match_count + 1

                                        If debug = "Y" Then mydebuglog.Debug("  > LON: " & LON & " -- " & Trim(CheckDBNull(dr(10), enumObjectType.StrType)))
                                        If Trim(LON) <> "" And IsNumeric(LON) And Trim(LON) = CheckDBNull(dr(10), enumObjectType.StrType) Then match_count = match_count + 1

                                        ' If not for the same person, can't possibly be a match
                                        If Trim(CON_ID) <> "" And Trim(CON_ID) <> Trim(CheckDBNull(dr(8), enumObjectType.StrType)) Then match_count = high_count - 1
                                    Catch ex As Exception
                                    End Try

                                    ' Save the values from the highest match
                                    If debug = "Y" Then mydebuglog.Debug("   >> Score: " & match_count.ToString & ", High Score Count: " & high_count.ToString)
                                    If match_count > high_count Then
                                        S_ADDR_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                        S_ADDR = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                        S_CITY = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                        S_STATE = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                        S_ZIPCODE = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                        S_COUNTRY = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                        S_COUNTY = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                        S_ADDR_MATCH = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                        S_CON_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                        S_LAT = CheckDBNull(dr(9), enumObjectType.StrType)
                                        S_LON = CheckDBNull(dr(10), enumObjectType.StrType)
                                        high_count = match_count
                                        If debug = "Y" Then mydebuglog.Debug("   >> Saving S_ADDR_ID: " & S_ADDR_ID & " with high_count: " & high_count.ToString)
                                    End If

                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading address: " & ex.ToString & vbCrLf
                                End Try
                            End While

                            ' If the best match is better than the confidence level, then return it
                            If high_count >= Confidence Then GoTo UpdPerAddrMatch
                            If debug = "Y" Then mydebuglog.Debug("  > Exited record without finding a match. Partial matches found: " & match_count.ToString)
                        End If
                        dr.Close()
                        If debug = "Y" Then mydebuglog.Debug("  # Matches found: " & returnv.ToString)
                        'Else
                        'If debug = "Y" Then mydebuglog.Debug("  No matches found")
                        'End If
                        GoTo ExitMatch
UpdPerAddrMatch:
                        Try
                            If debug = "Y" Then mydebuglog.Debug("  > Full match found: " & S_ADDR_ID & ". Score: " & high_count.ToString & vbCrLf & "  =========================" & vbCrLf)
                            ADDR_ID = S_ADDR_ID
                            ADDR = S_ADDR
                            CITY = S_CITY
                            STATE = S_STATE
                            ZIPCODE = S_ZIPCODE
                            COUNTRY = S_COUNTRY
                            COUNTY = S_COUNTY
                            'If ADDR_MATCH = "" Then ADDR_MATCH = S_ADDR_MATCH
                            ADDR_MATCH = S_ADDR_MATCH
                            CON_ID = S_CON_ID
                            If LAT = "0" Or LAT = "" Then LAT = S_LAT
                            If LON = "0" Or LAT = "" Then LON = S_LON
                        Catch ex As Exception
                            errmsg = errmsg & "Error reading record: " & ex.ToString & vbCrLf
                        End Try
                        dr.Close()

                        ' Call StandardizeAddress again
                        ' Supress (pass "N" instead of "Y") geocoding in CallStandardizeAddress sub due to new Asynch geocoding; Ren Hou; 2018-11-01;
                        Call CallStandardizeAddress(ADDR_ID, ORG_ID, CON_ID, ADDR_TYPE, "N", ADDR, CITY, _
                          STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, JURIS_ID, DELIVERABLE, debug, mydebuglog, errmsg, results, "X")
                        If results = "Failure" Then GoTo CloseOut2

                    Case "O"
                        ' ----
                        ' Business Address
                        returnv = 0
                        SqlS = "SELECT A.ROW_ID,(SELECT CASE WHEN A.ADDR IS NULL THEN '' ELSE A.ADDR END)," & _
                        "(SELECT CASE WHEN A.CITY IS NULL THEN '' ELSE A.CITY END)," & _
                        "(SELECT CASE WHEN A.STATE IS NULL THEN '' ELSE A.STATE END)," & _
                        "(SELECT CASE WHEN A.ZIPCODE IS NULL THEN '' ELSE A.ZIPCODE END)," & _
                        "(SELECT CASE WHEN A.COUNTRY IS NULL THEN '' ELSE A.COUNTRY END)," & _
                        "(SELECT CASE WHEN A.COUNTY IS NULL THEN '' ELSE A.COUNTY END)," & _
                        "(SELECT CASE WHEN A.X_MATCH_CD IS NULL THEN '' ELSE A.X_MATCH_CD END)," & _
                        "(SELECT CASE WHEN A.OU_ID IS NULL THEN '' ELSE A.OU_ID END)," & _
                        "(SELECT CASE WHEN A.X_LAT IS NULL THEN '0' ELSE CAST(A.X_LAT AS VARCHAR) END)," & _
                        "(SELECT CASE WHEN A.X_LONG IS NULL THEN '0' ELSE CAST(A.X_LONG AS VARCHAR) END),"
                        If ORG_ID <> "" Or ORG_MATCH <> "" Or ORG_PHONE <> "" Then
                            SqlS = SqlS & "(SELECT CASE WHEN O.MAIN_PH_NUM IS NULL THEN '' ELSE O.MAIN_PH_NUM END), " & _
                            "(SELECT CASE WHEN O.DEDUP_TOKEN IS NULL THEN '' ELSE O.DEDUP_TOKEN END) "
                        Else
                            SqlS = SqlS & "'', '' "
                        End If
                        SqlS = SqlS & "FROM siebeldb.dbo.S_ADDR_ORG A "
                        If ADDR_MATCH <> "" And ADDR_MATCH <> "$$$$$$$$$$$$$$$" Then
                            If ORG_ID <> "" Or ORG_MATCH <> "" Or ORG_PHONE <> "" Then
                                SqlS = SqlS & "WITH (INDEX([S_ADDR_ORG_MC_X])) " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT O ON O.ROW_ID=A.OU_ID "
                                SqlS = SqlS & "WHERE (A.X_MATCH_CD='" & ADDR_MATCH & "') OR ("
                                If ORG_ID <> "" Then
                                    SqlS = SqlS & "(O.ROW_ID='" & ORG_ID & "' AND "
                                    If CITY <> "" Then SqlS = SqlS & "UPPER(A.CITY)='" & UCase(CITY) & "' AND "
                                    If STATE <> "" Then SqlS = SqlS & "A.STATE='" & UCase(STATE) & "'"
                                    If Right(SqlS, 4) = "AND " Then SqlS = Left(SqlS, Len(SqlS) - 4)
                                    SqlS = SqlS & ") OR "
                                End If
                                If ORG_PHONE <> "" Then
                                    SqlS = SqlS & "(O.MAIN_PH_NUM='" & ORG_PHONE & "'"
                                    If CITY <> "" Then SqlS = SqlS & " AND UPPER(A.CITY)='" & UCase(CITY) & "'"
                                    If STATE <> "" Then SqlS = SqlS & " AND A.STATE='" & UCase(STATE) & "'"
                                    SqlS = SqlS & ")"
                                End If
                                If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                                SqlS = SqlS & ")"
                                If Right(SqlS, 6) = "AND ()" Then SqlS = Left(SqlS, Len(SqlS) - 6)
                            Else
                                SqlS = SqlS & "WITH (INDEX([S_ADDR_ORG_U1])) " & _
                                    "WHERE (A.X_MATCH_CD='" & ADDR_MATCH & "') AND "
                                SqlS = SqlS & "(UPPER(A.ADDR)='" & UCase(ADDR) & "' OR " & _
                                    "UPPER(A.CITY)='" & UCase(CITY) & "' OR " & _
                                    "UPPER(A.STATE)='" & UCase(STATE) & "')"
                            End If
                        Else
                            If ORG_ID <> "" Or ORG_MATCH <> "" Or ORG_PHONE <> "" Then
                                SqlS = SqlS & "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT O ON O.ROW_ID=A.OU_ID " & _
                                "WHERE "
                                If ORG_ID <> "" Then SqlS = SqlS & "O.ROW_ID='" & ORG_ID & "' OR "
                                If ORG_MATCH <> "" Then SqlS = SqlS & "O.DEDUP_TOKEN='" & ORG_MATCH & "' OR "
                                If ORG_PHONE <> "" Then SqlS = SqlS & "O.MAIN_PH_NUM='" & ORG_PHONE & "'"
                                If Right(SqlS, 3) = "OR " Then SqlS = Left(SqlS, Len(SqlS) - 3)
                            Else
                                SqlS = SqlS & "WHERE (UPPER(A.ADDR)='" & UCase(ADDR) & "' AND " & _
                                    "UPPER(A.CITY)='" & UCase(CITY) & "' AND " & _
                                    "UPPER(A.STATE)='" & UCase(STATE) & "')"
                            End If
                        End If
                        If debug = "Y" Then mydebuglog.Debug("  =========================" & vbCrLf & "  Checking matches: " & SqlS)
                        Try
                            'cmd.CommandText = SqlS
                            'dr = cmd.ExecuteReader()
                            '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;
                            cmd_ro.CommandText = SqlS
                            dr = cmd_ro.ExecuteReader()

                            If Not dr Is Nothing Then
                                ' -----
                                ' Evaluate all of the records found and locate the best match
                                While dr.Read()
                                    Try
                                        returnv = returnv + 1
                                        ' Get variables
                                        match_count = 0
                                        If debug = "Y" Then mydebuglog.Debug("  > Checking address match: " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)))
                                        Try
                                            ' If the IDs match - no more work to be done
                                            If Trim(ADDR_ID) <> "" And Trim(ADDR_ID) = Trim(CheckDBNull(dr(0), enumObjectType.StrType)) Then
                                                If debug = "Y" Then mydebuglog.Debug("    - ADDR ID MATCH: " & ADDR_ID & " to " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)))
                                                S_ADDR_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                                S_ADDR = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                                S_CITY = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                                S_STATE = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                                S_ZIPCODE = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                                S_COUNTRY = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                                S_COUNTY = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                                S_ADDR_MATCH = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                                S_ORG_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                                S_LAT = Trim(CheckDBNull(dr(9), enumObjectType.StrType))
                                                S_LON = Trim(CheckDBNull(dr(10), enumObjectType.StrType))
                                                GoTo UpdBusAddrMatch
                                            End If

                                            ' If no id match, check sub-matches and score
                                            temp_match = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                            temp_phone = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                            temp_addr = UCase(Trim(CheckDBNull(dr(1), enumObjectType.StrType)))
                                            temp_city = UCase(Trim(CheckDBNull(dr(2), enumObjectType.StrType)))
                                            temp_state = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                            temp_zip = Left(Trim(CheckDBNull(dr(4), enumObjectType.StrType)), 5)
                                            If temp_phone <> "" Then temp_phone = StndPhone(temp_phone)
                                            If ADDR_MATCH <> "" And ADDR_MATCH <> "$$$$$$$$$$$$$$$" Then
                                                If Trim(ADDR_MATCH) = temp_match Then match_count = match_count + 1
                                                If Trim(ADDR_MATCH) <> temp_match Then match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ADDR_MATCH: " & ADDR_MATCH & " to " & temp_match & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If ORG_PHONE <> "" And ORG_PHONE = temp_phone Then
                                                match_count = match_count + 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ORG_PHONE: " & ORG_PHONE & " to " & temp_phone & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If ORG_PHONE <> "" And ORG_PHONE <> temp_phone Then
                                                match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ORG_PHONE: " & ORG_PHONE & " to " & temp_phone & ". unmatch_count so far: " & match_count.ToString)
                                            End If
                                            If ORG_ID <> "" Then
                                                If ORG_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType)) Then match_count = match_count + 2 Else match_count = match_count - 2
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ORG_ID: " & ORG_ID & " to " & Trim(CheckDBNull(dr(8), enumObjectType.StrType)) & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If ORG_MATCH <> "" And ORG_MATCH <> Trim(CheckDBNull(dr(12), enumObjectType.StrType)) Then
                                                match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ORG_MATCH: " & ORG_MATCH & " to " & Trim(CheckDBNull(dr(12), enumObjectType.StrType)) & ". unmatch_count so far: " & match_count.ToString)
                                            End If
                                            If ADDR_ID <> "" And ADDR_ID <> Trim(CheckDBNull(dr(0), enumObjectType.StrType)) Then
                                                match_count = match_count - 2
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ADDR_ID: " & ADDR_ID & " to " & Trim(CheckDBNull(dr(0), enumObjectType.StrType)) & ". unmatch_count so far: " & match_count.ToString)
                                            End If
                                            If ADDR <> "" Then
                                                If ADDR <> "" And (UCase(ADDR) = temp_addr Or O_ADDR = temp_addr) Then match_count = match_count + 1
                                                If ADDR <> "" And (UCase(ADDR) <> temp_addr And O_ADDR <> temp_addr) Then match_count = match_count - 2
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ADDR: '" & UCase(ADDR) & "' or '" & O_ADDR & "' to '" & temp_addr & "'. match_count so far: " & match_count.ToString)
                                            End If
                                            If CITY <> "" Then
                                                If UCase(CITY) = temp_city Then match_count = match_count + 1 Else match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - CITY: " & UCase(CITY) & " to " & temp_city & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If STATE <> "" Then
                                                If STATE = temp_state Then match_count = match_count + 1 Else match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - STATE: " & STATE & " to " & temp_state & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If ZIPCODE <> "" Then
                                                If Left(ZIPCODE, 5) = temp_zip Then match_count = match_count + 1 Else match_count = match_count - 1
                                                If Len(ZIPCODE) = 10 And Right(Trim(ZIPCODE), 4) <> Right(Trim(CheckDBNull(dr(4), enumObjectType.StrType)), 4) Then match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - ZIPCODE: " & Left(Trim(ZIPCODE), 5) & " to " & temp_zip & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If COUNTRY <> "" Then
                                                If Trim(COUNTRY) = Trim(CheckDBNull(dr(5), enumObjectType.StrType)) Then match_count = match_count + 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - COUNTRY: " & COUNTRY & " to " & Trim(CheckDBNull(dr(5), enumObjectType.StrType)) & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If COUNTY <> "" Then
                                                If UCase(COUNTY) = UCase(Trim(CheckDBNull(dr(6), enumObjectType.StrType))) Then match_count = match_count + 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - COUNTY: " & COUNTY & " to " & Trim(CheckDBNull(dr(6), enumObjectType.StrType)) & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If LAT <> "" And IsNumeric(LAT) Then
                                                If Trim(Str(Val(LAT))) = Trim(Str(Val(CheckDBNull(dr(9), enumObjectType.StrType)))) Then match_count = match_count + 1 Else match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - LAT: " & LAT & " to " & Trim(CheckDBNull(dr(9), enumObjectType.StrType)) & ". match_count so far: " & match_count.ToString)
                                            End If
                                            If LON <> "" And IsNumeric(LON) Then
                                                If Trim(Str(Val(LON))) = Trim(Str(Val(CheckDBNull(dr(10), enumObjectType.StrType)))) Then match_count = match_count + 1 Else match_count = match_count - 1
                                                If ddebug = "Y" Then mydebuglog.Debug("    - LON: " & LON & " to " & Trim(CheckDBNull(dr(10), enumObjectType.StrType)) & ". match_count so far: " & match_count.ToString)
                                            End If

                                            ' If the org id doesn't match, can't possibly be a match
                                            If Trim(ORG_ID) <> "" And Trim(ORG_ID) <> Trim(CheckDBNull(dr(8), enumObjectType.StrType)) Then
                                                If match_count >= Confidence And match_count > high_count Then
                                                    ' The address supplied duplicates another address that is linked to a different account.  Handle this as a new address
                                                    MatchNew = True
                                                    MatchNew_count = match_count
                                                End If
                                                match_count = match_count - 1
                                                'match_count = high_count - 1
                                            Else
                                                If match_count >= Confidence And match_count >= high_count And match_count >= MatchNew_count Then
                                                    MatchNew = False
                                                End If
                                            End If
                                            If ddebug = "Y" Then mydebuglog.Debug("    - MatchNew: " & MatchNew & " .. " & MatchNew_count.ToString)
                                        Catch ex As Exception
                                        End Try

                                        ' Save the values from the highest match
                                        If debug = "Y" Then mydebuglog.Debug("    - Score: " & match_count.ToString & ", High Score Count: " & high_count.ToString & " for '" & S_ADDR_ID & "'")
                                        If match_count > high_count Then
                                            If MatchNew Then
                                                S_ADDR_ID = ""
                                            Else
                                                S_ADDR_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                            End If
                                            S_ADDR = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                            S_CITY = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                            S_STATE = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                            S_ZIPCODE = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                            S_COUNTRY = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                            S_COUNTY = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                            S_ADDR_MATCH = Trim(CheckDBNull(dr(7), enumObjectType.StrType))
                                            If MatchNew Then
                                                S_ORG_ID = ""
                                            Else
                                                S_ORG_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                            End If
                                            S_LAT = CheckDBNull(dr(9), enumObjectType.StrType)
                                            S_LON = CheckDBNull(dr(10), enumObjectType.StrType)
                                            high_count = match_count
                                            If debug = "Y" Then mydebuglog.Debug("   >> Saving S_ADDR_ID: " & S_ADDR_ID & " with high_count: " & high_count.ToString)
                                        End If

                                    Catch ex As Exception
                                        errmsg = errmsg & "Error reading address: " & ex.ToString & vbCrLf & vbCrLf & "Query: " & SqlS
                                    End Try
                                End While

                                ' If the best match is better than the confidence level, then return it
                                If high_count >= Confidence Then GoTo UpdBusAddrMatch
                                If debug = "Y" Then mydebuglog.Debug("  > Exited record without finding a match. Partial matches found: " & match_count.ToString)
                            End If
                        Catch ex As Exception
                            errmsg = errmsg & "Error finding addresses: " & ex.ToString & vbCrLf & "Query: " & SqlS
                        End Try

                        Try
                            dr.Close()
                        Catch ex As Exception
                        End Try
                        If debug = "Y" Then mydebuglog.Debug("  # Matches found: " & returnv.ToString)
                        GoTo ExitMatch
UpdBusAddrMatch:
                        Try
                            If debug = "Y" Then mydebuglog.Debug("  > Full match found: " & S_ADDR_ID & ". Score: " & high_count.ToString & vbCrLf & "  =========================" & vbCrLf)
                            ADDR_ID = S_ADDR_ID
                            ADDR = S_ADDR
                            CITY = S_CITY
                            STATE = S_STATE
                            ZIPCODE = S_ZIPCODE
                            COUNTRY = S_COUNTRY
                            COUNTY = S_COUNTY
                            ADDR_MATCH = S_ADDR_MATCH
                            If Not MatchNew Then
                                ORG_ID = S_ORG_ID
                                If ADDR_MATCH = "" Then ADDR_MATCH = S_ADDR_MATCH
                            End If
                            If LAT = "0" Or LAT = "" Then LAT = S_LAT
                            If LON = "0" Or LAT = "" Then LON = S_LON
                        Catch ex As Exception
                            errmsg = errmsg & "Error reading record: " & ex.ToString & vbCrLf
                        End Try
                        dr.Close()

                        ' Call StandardizeAddress again if ADDR_ID is not provided
                        If ADDR_ID = "" Or ADDR_MATCH = "" Then
                            ' Supress (pass "N" instead of "Y") geocoding in CallStandardizeAddress sub due to new Asynch geocoding; Ren Hou; 2018-11-01;
                            Call CallStandardizeAddress(ADDR_ID, ORG_ID, CON_ID, ADDR_TYPE, "N", ADDR, CITY, _
                              STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, JURIS_ID, DELIVERABLE, debug, mydebuglog, errmsg, results, "X")
                            If results = "Failure" Then GoTo CloseOut2
                        End If

                End Select
            End If

ExitMatch:
            ' ============================================
            ' Database operations
            Try
                dr.Close()
            Catch ex As Exception
            End Try
            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "====================" & vbCrLf & "Database Operation: " & database & " for type " & ADDR_TYPE & " - existing ID: " & ADDR_ID)
                If ORG_ID <> "" Then mydebuglog.Debug(" and ORG_ID: " & ORG_ID)
            End If

            ' Fix lat/lon if necessary
            If LAT = "" Then LAT = "0"
            If LON = "" Then LON = "0"

            ' Create record
            If database = "C" Then
                If ADDR_ID = "" And ADDR <> "" Then
                    If ADDR_TYPE = "P" And CON_ID = "" Then GoTo UpdateAddr ' Skip if personal and no contact id
                    If ADDR_TYPE = "O" And ORG_ID = "" Then GoTo UpdateAddr ' Skip if organizational and no organization id

InsertAddress:
                    ' Generate random address id for id lookup
                    Select Case ADDR_TYPE
                        Case "P"
                            ADDR_ID = BasicService.GenerateRecordId("S_ADDR_PER", "N", debug)
                        Case "O"
                            ADDR_ID = BasicService.GenerateRecordId("S_ADDR_ORG", "N", debug)
                    End Select
                    If debug = "Y" Then mydebuglog.Debug("    > Address id generated: " & ADDR_ID)

                    ' Create address record 
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "INSERT INTO siebeldb.dbo.S_ADDR_PER " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM," & _
                            "INTEGRATION_ID,MODIFICATION_NUM,CONFLICT_ID,DISA_CLEANSE_FLG,PER_ID,ADDR,CITY,COMMENTS," & _
                            "COUNTY,COUNTRY,STATE,ZIPCODE,X_MATCH_CD," & _
                            "X_MATCH_DT, X_LAT, X_LONG, X_JURIS_ID,X_CASS_CHECKED,X_CASS_CODE) " & _
                            "VALUES ('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0," & _
                            "'" & ADDR_ID & "',0,0,'N','" & CON_ID & "','" & SqlString(ADDR) & "','" & SqlString(CITY) & "','From CleanAddress', '" & _
                            SqlString(COUNTY) & "','" & COUNTRY & "','" & STATE & "', '" & ZIPCODE & "','" & ADDR_MATCH & _
                            "',GETDATE(),'" & LAT & "','" & LON & "','" & JURIS_ID & "',GETDATE(),'" & DELIVERABLE & "')"
                        Case "O"
                            SqlS = "INSERT INTO siebeldb.dbo.S_ADDR_ORG " & _
                            "(ROW_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,DCKING_NUM," & _
                            "INTEGRATION_ID,MODIFICATION_NUM,CONFLICT_ID,DISA_CLEANSE_FLG,OU_ID,ADDR,CITY,COMMENTS," & _
                            "COUNTY,COUNTRY,STATE,ZIPCODE,X_MATCH_CD," & _
                            "X_MATCH_DT, X_LAT, X_LONG, X_JURIS_ID,X_CASS_CHECKED,X_CASS_CODE) " & _
                            "VALUES ('" & ADDR_ID & "',GETDATE(),'0-1',GETDATE(),'0-1',0," & _
                            "'" & ADDR_ID & "',0,0,'N','" & ORG_ID & "','" & SqlString(ADDR) & "','" & SqlString(CITY) & "','From CleanAddress', '" & _
                            SqlString(COUNTY) & "','" & COUNTRY & "','" & STATE & "', '" & ZIPCODE & "','" & ADDR_MATCH & _
                            "',GETDATE(),'" & LAT & "','" & LON & "','" & JURIS_ID & "',GETDATE(),'" & DELIVERABLE & "')"
                    End Select
                    temp = ExecQuery("Create", "Address record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp

                    ' Get address id of record created
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "SELECT ROW_ID " & _
                            "FROM siebeldb.dbo.S_ADDR_PER " & _
                            "WHERE INTEGRATION_ID='" & ADDR_ID & "' AND X_MATCH_CD='" & ADDR_MATCH & "'"
                        Case "O"
                            SqlS = "SELECT ROW_ID " & _
                            "FROM siebeldb.dbo.S_ADDR_ORG " & _
                            "WHERE INTEGRATION_ID='" & ADDR_ID & "' AND X_MATCH_CD='" & ADDR_MATCH & "'"
                    End Select
                    ADDR_ID = ""
                    If debug = "Y" Then mydebuglog.Debug("  Verifying address created: " & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        '2020-05-18; Ren Hou; Chnaged to Read-only per Chris;  
                        '2021-07-08: Rebecca; AG replication not fast enough to use the RO member to validate a record was written
                        'cmd_ro.CommandText = SqlS
                        'dr = cmd_ro.ExecuteReader()

                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    ADDR_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                Catch ex As Exception
                                    errmsg = errmsg & "Error looking up address id: " & ex.ToString & vbCrLf
                                End Try
                            End While
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & vbCrLf & "Error looking up address id. " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                    dr.Close()
                    If debug = "Y" Then mydebuglog.Debug("    > ADDR_ID created: " & ADDR_ID)

                    If ADDR_ID = "" Then
                        ba_count = ba_count + 1
                        If ba_count < 3 Then
                            GoTo InsertAddress
                        Else
                            errmsg = errmsg & vbCrLf & "Unable to insert address record. "
                            results = "Failure"
                        End If
                    Else
                        SqlS = ""
                        Select Case ADDR_TYPE
                            Case "P"
                                SqlS = "UPDATE siebeldb.dbo.S_CONTACT " & _
                                "SET PR_PER_ADDR_ID='" & ADDR_ID & "' " & _
                                "WHERE ROW_ID='" & CON_ID & "' AND (PR_PER_ADDR_ID IS NULL OR PR_PER_ADDR_ID='')"
                                temp = ExecQuery("Update", "Contact record", cmd, SqlS, mydebuglog, debug)
                                errmsg = errmsg & temp
                            Case "O"
                                SqlS = "UPDATE siebeldb.dbo.S_ORG_EXT " & _
                                "SET PR_ADDR_ID='" & ADDR_ID & "' " & _
                                "WHERE ROW_ID='" & ORG_ID & "' AND (PR_ADDR_ID IS NULL OR PR_ADDR_ID='')"
                                temp = ExecQuery("Update", "Account record", cmd, SqlS, mydebuglog, debug)
                                errmsg = errmsg & temp
                        End Select
                    End If
                Else
                    If debug = "Y" Then mydebuglog.Debug("Record already exists - no need to recreate - update instead as applicable from the standardized code. ")
                    'errmsg = errmsg & vbCrLf & "Record already exists - no need to recreate. "
                    If ADDR_ID <> "" And ADDR <> "" Then
                        Select Case ADDR_TYPE
                            Case "P"
                                SqlS = "UPDATE siebeldb.dbo.S_ADDR_PER SET LAST_UPD=GETDATE()"
                            Case "O"
                                SqlS = "UPDATE siebeldb.dbo.S_ADDR_ORG SET LAST_UPD=GETDATE()"
                        End Select
                        If A_COUNTY <> "" Then SqlS = SqlS & ",COUNTY='" & SqlString(A_COUNTY) & "'"
                        If A_ADDR_MATCH <> "" Then SqlS = SqlS & ",X_MATCH_CD='" & A_ADDR_MATCH & "',X_MATCH_DT=GETDATE()"
                        If DELIVERABLE <> "" Then SqlS = SqlS & ",X_CASS_CHECKED=GETDATE(),X_CASS_CODE='" & DELIVERABLE & "'"
                        If A_LAT <> "" Then SqlS = SqlS & ",X_LAT='" & A_LAT & "'"
                        If A_LON <> "" Then SqlS = SqlS & ",X_LONG='" & A_LON & "'"
                        If A_COUNTRY <> "" Then SqlS = SqlS & ",COUNTRY='" & SqlString(A_COUNTRY) & "'"
                        If A_ADDR <> "" Then SqlS = SqlS & ",ADDR='" & SqlString(A_ADDR) & "'"
                        If A_STATE <> "" Then SqlS = SqlS & ",STATE='" & SqlString(A_STATE) & "'"
                        If A_CITY <> "" Then SqlS = SqlS & ",CITY='" & SqlString(A_CITY) & "'"
                        If A_JURIS_ID <> "" Then SqlS = SqlS & ",X_JURIS_ID='" & A_JURIS_ID & "'"
                        SqlS = SqlS & " WHERE ROW_ID='" & ADDR_ID & "'"
                        temp = ExecQuery("Update", "Address record", cmd, SqlS, mydebuglog, debug)
                        errmsg = errmsg & temp
                        results = "Success"
                    End If
                End If

            End If

            '-----
            ' Update record
UpdateAddr:
            If database = "U" Then
                If ADDR_ID <> "" And ADDR <> "" Then
                    Select Case ADDR_TYPE
                        Case "P"
                            SqlS = "UPDATE siebeldb.dbo.S_ADDR_PER SET LAST_UPD=GETDATE()," & _
                            "ADDR='" & SqlString(ADDR) & "',CITY='" & SqlString(CITY) & "',STATE='" & SqlString(STATE) & "',COUNTRY='" & COUNTRY & "'," & _
                            "ZIPCODE='" & ZIPCODE & "',X_MATCH_CD='" & ADDR_MATCH & "',X_MATCH_DT=GETDATE(),X_CASS_CHECKED=GETDATE(),X_CASS_CODE='" & DELIVERABLE & "'"
                            If COUNTY <> "" Then SqlS = SqlS & ",COUNTY='" & SqlString(COUNTY) & "'"
                            If LAT <> "" Then SqlS = SqlS & ",X_LAT='" & LAT & "'"
                            If LON <> "" Then SqlS = SqlS & ",X_LONG='" & LON & "'"
                            If JURIS_ID <> "" Then SqlS = SqlS & ",X_JURIS_ID='" & JURIS_ID & "'"
                            SqlS = SqlS & " WHERE ROW_ID='" & ADDR_ID & "'"
                        Case "O"
                            SqlS = "UPDATE siebeldb.dbo.S_ADDR_ORG SET LAST_UPD=GETDATE()," & _
                            "ADDR='" & SqlString(ADDR) & "',CITY='" & SqlString(CITY) & "',STATE='" & SqlString(STATE) & "',COUNTRY='" & COUNTRY & "'," & _
                            "ZIPCODE='" & ZIPCODE & "',X_MATCH_CD='" & ADDR_MATCH & "',X_MATCH_DT=GETDATE(),X_CASS_CHECKED=GETDATE(),X_CASS_CODE='" & DELIVERABLE & "'"
                            If COUNTY <> "" Then SqlS = SqlS & ",COUNTY='" & SqlString(COUNTY) & "'"
                            If LAT <> "" Then SqlS = SqlS & ",X_LAT='" & LAT & "'"
                            If LON <> "" Then SqlS = SqlS & ",X_LONG='" & LON & "'"
                            If JURIS_ID <> "" Then SqlS = SqlS & ",X_JURIS_ID='" & JURIS_ID & "'"
                            SqlS = SqlS & " WHERE ROW_ID='" & ADDR_ID & "'"
                    End Select
                    temp = ExecQuery("Update", "Address record", cmd, SqlS, mydebuglog, debug)
                    errmsg = errmsg & temp
                Else
                    errmsg = errmsg & vbCrLf & "Error updating address record. "
                    results = "Failure"
                End If
            End If

            ' Update LAT LON using Asynch geocoding.
            If ADDR_ID <> "" And ADDR <> "" Then
                Call CallStandardizeAddress(ADDR_ID, ORG_ID, CON_ID, ADDR_TYPE, "Y", ADDR, CITY, _
                  STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, JURIS_ID, DELIVERABLE, debug, mydebuglog, errmsg, results, "V") 'pass "V" for database parameter to supress address validation
                'If results = "Failure" Then GoTo CloseOut2
            End If
        Next

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            dr = Nothing
            con.Close()
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
            con_ro.Close()
            con_ro.Dispose()
            con_ro = Nothing
            cmd_ro.Dispose()
            cmd_ro = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Unable to close the database connection. " & vbCrLf

        End Try

CloseOut2:
        ' ============================================
        ' Return the standardized information as an XML document:
        '   <AddressRec>
        '       <AddrId>   
        '       <ConId>   
        '       <OrgId>   
        '       <JurisId>   
        '       <MatchCode>
        '       <Type>
        '       <Street>        
        '       <City>           
        '       <State>          
        '       <County>         
        '       <Zipcode>         
        '       <Country>         
        '       <Lat>         
        '       <Long>
        '   </AddressRec>

        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("AddressRec")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            If debug <> "T" Then
                AddXMLChild(odoc, resultsRoot, "AddrId", HttpUtility.UrlEncode(IIf(ADDR_ID = "", " ", ADDR_ID)))
                AddXMLChild(odoc, resultsRoot, "ConId", HttpUtility.UrlEncode(IIf(CON_ID = "", " ", CON_ID)))
                AddXMLChild(odoc, resultsRoot, "OrgId", HttpUtility.UrlEncode(IIf(ORG_ID = "", " ", ORG_ID)))
                AddXMLChild(odoc, resultsRoot, "JurisId", HttpUtility.UrlEncode(IIf(JURIS_ID = "", " ", JURIS_ID)))
                AddXMLChild(odoc, resultsRoot, "MatchCode", HttpUtility.UrlEncode(IIf(ADDR_MATCH = "", " ", ADDR_MATCH)))
                AddXMLChild(odoc, resultsRoot, "Type", IIf(ADDR_TYPE = "", " ", ADDR_TYPE))
                AddXMLChild(odoc, resultsRoot, "Address", HttpUtility.UrlEncode(IIf(ADDR = "", " ", ADDR)))
                AddXMLChild(odoc, resultsRoot, "City", HttpUtility.UrlEncode(IIf(CITY = "", " ", CITY)))
                AddXMLChild(odoc, resultsRoot, "State", HttpUtility.UrlEncode(IIf(STATE = "", " ", STATE)))
                AddXMLChild(odoc, resultsRoot, "County", HttpUtility.UrlEncode(IIf(COUNTY = "", " ", COUNTY)))
                AddXMLChild(odoc, resultsRoot, "Zipcode", HttpUtility.UrlEncode(IIf(ZIPCODE = "", " ", ZIPCODE)))
                AddXMLChild(odoc, resultsRoot, "Country", HttpUtility.UrlEncode(IIf(COUNTRY = "", " ", COUNTRY)))
                AddXMLChild(odoc, resultsRoot, "Lat", HttpUtility.UrlEncode(IIf(LAT = "", " ", LAT)))
                AddXMLChild(odoc, resultsRoot, "Long", HttpUtility.UrlEncode(IIf(LON = "", " ", LON)))
            Else
                If ADDR_MATCH <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        ' ============================================
        ' Close the log file if any
        If results = "Failure" Then myeventlog.Error("CleanAddress : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("CleanAddress : Results: " & results & " for " & ADDR & " generated matchcode " & ADDR_MATCH)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for " & ADDR & " generated matchcode " & ADDR_MATCH)
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "100"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function

    <WebMethod(Description:="Cleans and/or matches a provided record with existing ones")> _
    Public Function CleanRecord(ByVal sXML As String) As XmlDocument
        ' This service is designed to "clean" the entire record provided.  It does this by invoking
        ' functions that do the following:
        '   Standardize the information in the record and generate match code
        '   Use the information provided to look up matches
        '   Update the record (if applicable)
        ' If invoked with a database value of "L" though, this function is designed to locate a record
        ' and return values found if a match is located

        '	<Records>
        '	<Record>
        '	<Debug></Debug>			- Debug mode "Y"/"N"/"T"
        '	<Database></Database>	- Database "U"pdate, "C"reate, "L"ocate or "X" No change
        '	<Confidence></Confidence>-Confidence level of matching
        '	<AddrId></AddrId>       - Address Id of a related address, if applicable
        '	<AddrType></AddrType>	- "P"ersonal or "O"rganization address
        '	<AddrMatch></AddrMatch> - Address match code
        '	<GeoCode></GeoCode>		- Geocode "Y"/"N"
        '	<JurisId></JurisId>		- Jurisdiction Id
        '	<Address></Address>		- Street Address
        '	<City></City>			- City
        '	<State></State>			- State or province
        '	<County></County>		- County or region
        '	<Zipcode></Zipcode>		- Zipcode or postal code
        '	<Country></Country>		- Country
        '	<ConId></ConId>			- Contact Id
        '	<ConMatch></ConMatch>	- Contact match code
        '	<PartId></PartId>		- Participant Id
        '	<FirstName></FirstName> - First name of contact
        '	<MidName></MidName>     - Middle name of contact
        '	<LastName></LastName>   - Last name of contact
        '	<Gender></Gender>       - Gender of contact
        '	<DOB></DOB>             - Date of Birth
        '	<WorkPhone></WorkPhone> - Work phone of a contact at the organization
        '	<SSN></SSN>             - Social Security Number
        '	<EmailAddr></EmailAddr> - Email address
        '	<RegNum></RegNum>       - Web registration id
        '	<SubConId></SubConId>   - Subscription contact record
        '	<HomePhone></HomePhone> - Home phone number
        '	<TrainerNo></TrainerNo> - Trainer number 
        '	<Name></Name>           - Name of organization
        '	<Loc></Loc>             - Location of organization
        '	<OrgId></OrgId>			- Organization Id 
        '	<OrgMatch></OrgMatch>   - Organization match code
        '	<OrgPhone></OrgPhone>   - Organization main phone number
        '   <JobTitle></JobTitle>   - Contact Job Title
        '   <PerTitle></PerTitle>   - Contact Personal Title
        '   <Source></Source>       - Contact Source
        '   <Industry></Industry>   - Organization industry
        '	</Record></Records>

        ' Process:
        '   1. Standardize address and organization
        '   2. Clean contact 
        '   3. Clean organization
        '   4. Clean address
        '   5. Update/insert records as necessary

        ' web.config Parameters used:
        '   hcidb        - connection string to siebeldb database

        ' Generic variables
        Dim results As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim mypath, debug, errmsg, logging, wp As String

        ' Database declarations
        Dim returnv As Integer

        ' Logging declarations
        Dim fs As FileStream
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("CRDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Data declarations
        Dim ADDR, CITY, STATE, ZIPCODE, COUNTY, COUNTRY, GEOCODE, ADDR_TYPE, ADDR_ID, ADDR_MATCH As String
        Dim temp, database, LAT, LON, ORG_ID, CON_ID, JURIS_ID, ORG_MATCH, ORG_PHONE, CON_MATCH As String
        Dim FST_NAME, MID_NAME, LAST_NAME, GENDER, FULL_NAME As String
        Dim DOB, SSN, EMAIL_ADDR, NAME, LOC, PER_TITLE As String
        Dim REG_NUM, SUB_CON_ID, TRAINER_NO, HOME_PHONE, PART_ID, JOB_TITLE, SOURCE, INDUSTRY As String
        Dim WORK_PH_NUM, DELIVERABLE, temp_match, temp_phone As String
        Dim match_count, Confidence As Integer
        Dim Personal As Boolean

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        logging = "Y"
        errmsg = ""
        results = "Success"
        ADDR_ID = ""
        ADDR_TYPE = "O"
        GEOCODE = "N"
        ADDR = "1101 Wilson Suite 1700"
        CITY = "Arlington"
        STATE = "VA"
        COUNTY = ""
        ZIPCODE = ""
        COUNTRY = ""
        ADDR_MATCH = "ZZ0Z$LW4PZI00$$&YWPF~PV&HHH0$$"
        JURIS_ID = ""
        LAT = ""
        LON = ""
        CON_ID = ""
        CON_MATCH = ""
        FST_NAME = "Christopher"
        MID_NAME = "L"
        LAST_NAME = "Bobbitt"
        GENDER = ""
        FULL_NAME = ""
        PART_ID = ""
        DOB = ""
        WORK_PH_NUM = "800-438-8477"
        SSN = ""
        EMAIL_ADDR = "bobbittc@gettips.com"
        ADDR_MATCH = ""
        REG_NUM = ""
        SUB_CON_ID = ""
        TRAINER_NO = ""
        HOME_PHONE = ""
        NAME = "Health Communications, Inc."
        LOC = ""
        ORG_ID = ""
        ORG_MATCH = ""
        ORG_PHONE = "800-438-8477"
        JOB_TITLE = ""
        SOURCE = ""
        INDUSTRY = ""
        PER_TITLE = ""
        DELIVERABLE = ""
        Personal = False

        temp_match = ""
        temp_phone = ""
        temp = ""
        database = ""
        returnv = 0
        wp = ""
        Confidence = 5
        match_count = 0

        ' ============================================
        ' Check parameters
        debug = "N"
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut
        End If
        HttpUtility.UrlDecode(sXML)
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//Records/Record")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                results = "Failure"
                GoTo CloseOut
            End Try
        Next
        debug = UCase(debug)

        ' Write XML query to file if debug is set
        If debug = "Y" Then
            logfile = "C:\Logs\clean_record_XML.log"
            Try
                If File.Exists(logfile) Then
                    fs = New FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Write)
                Else
                    fs = New FileStream(logfile, FileMode.CreateNew, FileAccess.Write, FileShare.Write)
                End If
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut
            End Try
            writeoutputfs(fs, Now.ToString & " : " & sXML)
            fs.Close()
        End If

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\CleanRecord.log"
            Try
                log4net.GlobalContext.Properties("CRLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                results = "Failure"
                GoTo CloseOut
            End Try

            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  debug: " & debug)
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' ============================================
        ' Process data
        For i = 0 To oNodeList.Count - 1
            errmsg = ""

            ' ============================================
            ' Extract data from parameter string
            If debug <> "T" Then
                database = Left(GetNodeValue("Database", oNodeList.Item(i)), 1)
                temp = Trim(GetNodeValue("Confidence", oNodeList.Item(i)))
                If temp <> "" And IsNumeric(temp) Then Confidence = Int(temp)
                ADDR_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("AddrId", oNodeList.Item(i)))))
                ADDR_ID = KeySpace(ADDR_ID)
                ADDR_TYPE = Trim(GetNodeValue("AddrType", oNodeList.Item(i)))
                If ADDR_TYPE = "" And ORG_ID <> "" Then ADDR_TYPE = "O"
                If ADDR_TYPE = "" And CON_ID <> "" Then ADDR_TYPE = "P"
                ADDR_MATCH = HttpUtility.UrlDecode(Trim(GetNodeValue("AddrMatch", oNodeList.Item(i))))
                GEOCODE = Left(Trim(GetNodeValue("GeoCode", oNodeList.Item(i))), 1)
                If GEOCODE <> "Y" Then GEOCODE = "N"
                If database <> "L" Then GEOCODE = "Y"
                JURIS_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("JurisId", oNodeList.Item(i)))))
                JURIS_ID = KeySpace(JURIS_ID)
                ADDR = HttpUtility.UrlDecode(Trim(GetNodeValue("Address", oNodeList.Item(i))))
                CITY = HttpUtility.UrlDecode(Trim(GetNodeValue("City", oNodeList.Item(i))))
                STATE = HttpUtility.UrlDecode(Trim(GetNodeValue("State", oNodeList.Item(i))))
                COUNTY = HttpUtility.UrlDecode(Trim(GetNodeValue("County", oNodeList.Item(i))))
                ZIPCODE = HttpUtility.UrlDecode(Trim(GetNodeValue("Zipcode", oNodeList.Item(i))))
                COUNTRY = HttpUtility.UrlDecode(Trim(GetNodeValue("Country", oNodeList.Item(i))))
                CON_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("ConId", oNodeList.Item(i)))))
                CON_ID = KeySpace(CON_ID)
                CON_MATCH = HttpUtility.UrlDecode(Trim(GetNodeValue("ConMatch", oNodeList.Item(i))))
                FST_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("FirstName", oNodeList.Item(i))))
                MID_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("MidName", oNodeList.Item(i))))
                LAST_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("LastName", oNodeList.Item(i))))
                GENDER = HttpUtility.UrlDecode(GetNodeValue("Gender", oNodeList.Item(i)))
                DOB = HttpUtility.UrlDecode(GetNodeValue("DOB", oNodeList.Item(i)))
                WORK_PH_NUM = HttpUtility.UrlDecode(GetNodeValue("WorkPhone", oNodeList.Item(i)))
                If WORK_PH_NUM <> "" Then WORK_PH_NUM = StndPhone(WORK_PH_NUM)
                SSN = HttpUtility.UrlDecode(GetNodeValue("SSN", oNodeList.Item(i)))
                EMAIL_ADDR = HttpUtility.UrlDecode(GetNodeValue("EmailAddr", oNodeList.Item(i)))
                REG_NUM = HttpUtility.UrlDecode(GetNodeValue("RegNum", oNodeList.Item(i)))
                SUB_CON_ID = Trim(HttpUtility.UrlDecode(GetNodeValue("SubConId", oNodeList.Item(i))))
                SUB_CON_ID = KeySpace(SUB_CON_ID)
                HOME_PHONE = HttpUtility.UrlDecode(GetNodeValue("HomePhone", oNodeList.Item(i)))
                If HOME_PHONE <> "" Then HOME_PHONE = StndPhone(HOME_PHONE)
                TRAINER_NO = Trim(HttpUtility.UrlDecode(GetNodeValue("TrainerNo", oNodeList.Item(i)))) '6/27/21; Ren Hou; Added to trim Teainer_No;
                PART_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("PartId", oNodeList.Item(i)))))
                PART_ID = KeySpace(PART_ID)
                NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("Name", oNodeList.Item(i))))
                LOC = HttpUtility.UrlDecode(Trim(GetNodeValue("Loc", oNodeList.Item(i))))
                ORG_ID = Trim(HttpUtility.UrlDecode(Trim(GetNodeValue("OrgId", oNodeList.Item(i)))))
                ORG_ID = KeySpace(ORG_ID)
                ORG_MATCH = HttpUtility.UrlDecode(Trim(GetNodeValue("OrgMatch", oNodeList.Item(i))))
                ORG_PHONE = HttpUtility.UrlDecode(GetNodeValue("OrgPhone", oNodeList.Item(i)))
                If ORG_PHONE <> "" Then ORG_PHONE = StndPhone(ORG_PHONE)
                JOB_TITLE = HttpUtility.UrlDecode(Trim(GetNodeValue("JobTitle", oNodeList.Item(i))))
                PER_TITLE = HttpUtility.UrlDecode(Trim(GetNodeValue("PerTitle", oNodeList.Item(i))))
                SOURCE = HttpUtility.UrlDecode(Trim(GetNodeValue("Source", oNodeList.Item(i))))
                INDUSTRY = HttpUtility.UrlDecode(Trim(GetNodeValue("Industry", oNodeList.Item(i))))
            End If
            If debug = "Y" Then
                mydebuglog.Debug("INPUTS------" & vbCrLf & "  ADDR_ID: " & ADDR_ID)
                mydebuglog.Debug("  AddrMatch: " & ADDR_MATCH)
                mydebuglog.Debug("  AddrType: " & ADDR_TYPE)
                mydebuglog.Debug("  Confidence: " & Confidence)
                mydebuglog.Debug("  database: " & database & vbCrLf & "------------")
                mydebuglog.Debug("  Geocode: " & GEOCODE)
                mydebuglog.Debug("  JurisId: " & JURIS_ID)
                mydebuglog.Debug("  Address: " & ADDR)
                mydebuglog.Debug("  City: " & CITY)
                mydebuglog.Debug("  State: " & STATE)
                mydebuglog.Debug("  County: " & COUNTY)
                mydebuglog.Debug("  Zipcode: " & ZIPCODE)
                mydebuglog.Debug("  Country: " & COUNTRY)
                mydebuglog.Debug("  ConId: " & CON_ID)
                mydebuglog.Debug("  ConMatch: " & CON_MATCH)
                mydebuglog.Debug("  FirstName: " & FST_NAME)
                mydebuglog.Debug("  MidName: " & MID_NAME)
                mydebuglog.Debug("  LastName: " & LAST_NAME)
                mydebuglog.Debug("  PartId: " & PART_ID)
                mydebuglog.Debug("  Gender: " & GENDER)
                mydebuglog.Debug("  DOB: " & DOB)
                mydebuglog.Debug("  WorkPhone: " & WORK_PH_NUM)
                mydebuglog.Debug("  SSN: " & SSN)
                mydebuglog.Debug("  JobTitle: " & JOB_TITLE)
                mydebuglog.Debug("  PerTitle: " & PER_TITLE)
                mydebuglog.Debug("  EmailAddr: " & EMAIL_ADDR)
                mydebuglog.Debug("  RegNum: " & REG_NUM)
                mydebuglog.Debug("  SubConId: " & SUB_CON_ID)
                mydebuglog.Debug("  HomePhone: " & HOME_PHONE)
                mydebuglog.Debug("  TrainerNo: " & TRAINER_NO)
                mydebuglog.Debug("  Name: " & NAME)
                mydebuglog.Debug("  Loc: " & LOC)
                mydebuglog.Debug("  OrgId: " & ORG_ID)
                mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
                mydebuglog.Debug("  OrgPhone: " & ORG_PHONE)
                mydebuglog.Debug("  Source: " & SOURCE)
                mydebuglog.Debug("  Industry: " & INDUSTRY)
            End If

            ' ============================================
            ' Call Services to Standardize Address and Organization
            If ADDR_ID = "" And ADDR <> "" Then
                Call CallStandardizeAddress(ADDR_ID, ORG_ID, CON_ID, ADDR_TYPE, GEOCODE, ADDR, CITY, _
                  STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, JURIS_ID, DELIVERABLE, debug, mydebuglog, errmsg, results, "X")
                If results = "Failure" Then GoTo CloseOut
            End If

            If NAME = "" And LOC = "" And ORG_ID = "" Then Personal = True
            If Not Personal Then
                If ORG_ID = "" And NAME <> "" Then
                    Call CallStandardizeOrganization(ORG_ID, NAME, LOC, ORG_MATCH, debug, mydebuglog, errmsg, results, "X")
                    If results = "Failure" Then GoTo CloseOut
                End If
            End If

            ' ============================================
            ' Call Service to clean Contact
            Call CallCleanContact(CON_ID, PART_ID, FST_NAME, LAST_NAME, MID_NAME, _
              GENDER, DOB, WORK_PH_NUM, SSN, EMAIL_ADDR, ADDR_MATCH, _
              ADDR_TYPE, ORG_MATCH, ORG_PHONE, REG_NUM, SUB_CON_ID, TRAINER_NO, _
              HOME_PHONE, FULL_NAME, Confidence, CON_MATCH, ORG_ID, ADDR_ID, _
              JOB_TITLE, SOURCE, PER_TITLE, debug, mydebuglog, errmsg, results, database)
            If results = "Failure" Then GoTo CloseOut
            If NAME = "" And LOC = "" And ORG_ID = "" Then Personal = True
            If database = "L" Then
                GoTo CloseOut
            End If

            ' ============================================
            ' Call Service to clean Organization
            If Not Personal Then
                Call CallCleanOrganization(ORG_ID, Trim(NAME), Trim(LOC), ORG_MATCH, ORG_PHONE, _
                  ADDR_MATCH, ADDR_ID, WORK_PH_NUM, INDUSTRY, debug, mydebuglog, errmsg, results, database)
                If results = "Failure" Then GoTo CloseOut
                If debug = "Y" Then mydebuglog.Debug("  ORG_MATCH: " & ORG_MATCH)
            End If

            ' ============================================
            ' Call Service to clean Address
            Call CallCleanAddress(ADDR_ID, ORG_ID, CON_ID, ADDR_TYPE, GEOCODE, _
              ADDR, CITY, STATE, COUNTY, ZIPCODE, COUNTRY, ADDR_MATCH, LAT, LON, CON_MATCH, WORK_PH_NUM, _
              ORG_MATCH, ORG_PHONE, JURIS_ID, debug, mydebuglog, errmsg, results, database)
            If results = "Failure" Then GoTo CloseOut

            ' ============================================
            ' Call Service to update Contact
            If database = "C" Or database = "U" Then
                Call CallCleanContact(CON_ID, PART_ID, FST_NAME, LAST_NAME, MID_NAME, _
                  GENDER, DOB, WORK_PH_NUM, SSN, EMAIL_ADDR, ADDR_MATCH, _
                  ADDR_TYPE, ORG_MATCH, ORG_PHONE, REG_NUM, SUB_CON_ID, TRAINER_NO, _
                  HOME_PHONE, FULL_NAME, Confidence, CON_MATCH, ORG_ID, ADDR_ID, _
                  JOB_TITLE, SOURCE, PER_TITLE, debug, mydebuglog, errmsg, results, "U")
                If results = "Failure" Then GoTo CloseOut
            End If
        Next

CloseOut:
        ' ============================================
        ' Return the standardized information as an XML document:
        '	<Record>
        '	<AddrId></AddrId>       - Address Id of a related address, if applicable
        '	<AddrType></AddrType>	- "P"ersonal or "O"rganization address
        '	<AddrMatch></AddrMatch> - Address match code
        '	<JurisId></JurisId>		- Jurisdiction Id
        '	<Address></Address>		- Street Address
        '	<City></City>			- City
        '	<State></State>			- State or province
        '	<County></County>		- County or region
        '	<Zipcode></Zipcode>		- Zipcode or postal code
        '	<Country></Country>		- Country
        '   <Lat></Lat>             - Latitude
        '   <Long></Long>           - Longitude
        '	<ConId></ConId>			- Contact Id
        '	<ConMatch></ConMatch>	- Contact match code
        '	<PartId></PartId>		- Participant Id
        '	<FirstName></FirstName> - First name of contact
        '	<MidName></MidName>     - Middle name of contact
        '	<LastName></LastName>   - Last name of contact
        '	<Gender></Gender>       - Gender of contact
        '	<DOB></DOB>             - Date of Birth
        '	<WorkPhone></WorkPhone> - Work phone of a contact at the organization
        '	<SSN></SSN>             - Social Security Number
        '	<EmailAddr></EmailAddr> - Email address
        '	<RegNum></RegNum>       - Web registration id
        '	<SubConId></SubConId>   - Subscription contact record
        '	<HomePhone></HomePhone> - Home phone number
        '	<TrainerNo></TrainerNo> - Trainer number 
        '	<Name></Name>           - Name of organization
        '	<Loc></Loc>             - Location of organization
        '	<OrgId></OrgId>			- Organization Id 
        '	<OrgMatch></OrgMatch>   - Organization match code
        '	<OrgPhone></OrgPhone>   - Organization main phone number
        '   <JobTitle></JobTitle>   - Contact Job Title
        '   <PerTitle></PerTitle>   - Contact Personal Title
        '   <Source></Source>       - Contact Source
        '   <Industry></Industry>   - Organization industry
        '	</Record>

        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("Record")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            If debug <> "T" Then
                AddXMLChild(odoc, resultsRoot, "AddrId", IIf(ADDR_ID.Trim = "", " ", HttpUtility.UrlEncode(ADDR_ID)))
                AddXMLChild(odoc, resultsRoot, "AddrType", IIf(ADDR_TYPE.Trim = "", " ", ADDR_TYPE))
                AddXMLChild(odoc, resultsRoot, "AddrMatch", IIf(ADDR_MATCH.Trim = "", " ", HttpUtility.UrlEncode(ADDR_MATCH)))
                AddXMLChild(odoc, resultsRoot, "JurisId", IIf(JURIS_ID.Trim = "", " ", HttpUtility.UrlEncode(JURIS_ID)))
                AddXMLChild(odoc, resultsRoot, "Address", IIf(ADDR.Trim = "", " ", HttpUtility.UrlEncode(ADDR)))
                AddXMLChild(odoc, resultsRoot, "City", IIf(CITY.Trim = "", " ", HttpUtility.UrlEncode(CITY)))
                AddXMLChild(odoc, resultsRoot, "State", IIf(STATE.Trim = "", " ", HttpUtility.UrlEncode(STATE)))
                AddXMLChild(odoc, resultsRoot, "County", IIf(COUNTY.Trim = "", " ", HttpUtility.UrlEncode(COUNTY)))
                AddXMLChild(odoc, resultsRoot, "Zipcode", IIf(ZIPCODE.Trim = "", " ", HttpUtility.UrlEncode(ZIPCODE)))
                AddXMLChild(odoc, resultsRoot, "Country", IIf(COUNTRY.Trim = "", " ", HttpUtility.UrlEncode(COUNTRY)))
                AddXMLChild(odoc, resultsRoot, "Lat", IIf(LAT = "", " ", HttpUtility.UrlEncode(LAT)))
                AddXMLChild(odoc, resultsRoot, "Long", IIf(LON = "", " ", HttpUtility.UrlEncode(LON)))
                AddXMLChild(odoc, resultsRoot, "ConId", IIf(CON_ID.Trim = "", " ", HttpUtility.UrlEncode(CON_ID)))
                AddXMLChild(odoc, resultsRoot, "ConMatch", IIf(CON_MATCH.Trim = "", " ", HttpUtility.UrlEncode(CON_MATCH)))
                AddXMLChild(odoc, resultsRoot, "PartId", IIf(PART_ID.Trim = "", " ", HttpUtility.UrlEncode(PART_ID)))
                AddXMLChild(odoc, resultsRoot, "FirstName", IIf(FST_NAME.Trim = "", " ", HttpUtility.UrlEncode(FST_NAME)))
                AddXMLChild(odoc, resultsRoot, "MidName", IIf(MID_NAME.Trim = "", " ", HttpUtility.UrlEncode(MID_NAME)))
                AddXMLChild(odoc, resultsRoot, "LastName", IIf(LAST_NAME.Trim = "", " ", HttpUtility.UrlEncode(LAST_NAME)))
                AddXMLChild(odoc, resultsRoot, "Gender", IIf(GENDER.Trim = "", " ", HttpUtility.UrlEncode(GENDER)))
                AddXMLChild(odoc, resultsRoot, "DOB", IIf(DOB.Trim = "", " ", HttpUtility.UrlEncode(DOB)))
                AddXMLChild(odoc, resultsRoot, "WorkPhone", IIf(WORK_PH_NUM.Trim = "", " ", HttpUtility.UrlEncode(WORK_PH_NUM)))
                AddXMLChild(odoc, resultsRoot, "SSN", IIf(SSN.Trim = "", " ", HttpUtility.UrlEncode(SSN)))
                AddXMLChild(odoc, resultsRoot, "EmailAddr", IIf(EMAIL_ADDR.Trim = "", " ", HttpUtility.UrlEncode(EMAIL_ADDR)))
                AddXMLChild(odoc, resultsRoot, "RegNum", IIf(REG_NUM.Trim = "", " ", HttpUtility.UrlEncode(REG_NUM)))
                AddXMLChild(odoc, resultsRoot, "SubConId", IIf(SUB_CON_ID.Trim = "", " ", HttpUtility.UrlEncode(SUB_CON_ID)))
                AddXMLChild(odoc, resultsRoot, "HomePhone", IIf(HOME_PHONE.Trim = "", " ", HttpUtility.UrlEncode(HOME_PHONE)))
                AddXMLChild(odoc, resultsRoot, "TrainerNo", IIf(TRAINER_NO.Trim = "", " ", HttpUtility.UrlEncode(TRAINER_NO)))
                AddXMLChild(odoc, resultsRoot, "Name", IIf(NAME.Trim = "", " ", HttpUtility.UrlEncode(NAME)))
                AddXMLChild(odoc, resultsRoot, "Loc", IIf(LOC.Trim = "", " ", HttpUtility.UrlEncode(LOC)))
                AddXMLChild(odoc, resultsRoot, "OrgId", IIf(ORG_ID.Trim = "", " ", HttpUtility.UrlEncode(ORG_ID)))
                AddXMLChild(odoc, resultsRoot, "OrgMatch", IIf(ORG_MATCH.Trim = "", " ", HttpUtility.UrlEncode(ORG_MATCH)))
                AddXMLChild(odoc, resultsRoot, "OrgPhone", IIf(ORG_PHONE.Trim = "", " ", HttpUtility.UrlEncode(ORG_PHONE)))
                AddXMLChild(odoc, resultsRoot, "JobTitle", IIf(JOB_TITLE.Trim = "", " ", HttpUtility.UrlEncode(JOB_TITLE)))
                AddXMLChild(odoc, resultsRoot, "PerTitle", IIf(PER_TITLE.Trim = "", " ", HttpUtility.UrlEncode(PER_TITLE)))
                AddXMLChild(odoc, resultsRoot, "Source", IIf(SOURCE = "", " ", HttpUtility.UrlEncode(SOURCE)))
                AddXMLChild(odoc, resultsRoot, "Industry", IIf(INDUSTRY = "", " ", HttpUtility.UrlEncode(INDUSTRY)))
            Else
                If ADDR_MATCH <> "" Then
                    results = "Success"
                Else
                    results = "Failure"
                End If
                AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("CleanRecord : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("CleanRecord : Results: " & results & " for " & ADDR_ID & "/" & CON_ID & "/" & ORG_ID)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                mydebuglog.Debug("  Results: " & results & " for " & ADDR_ID & "/" & CON_ID & "/" & ORG_ID & "   Started: " & LogStartTime.ToString & "  Ended: " & Now.ToString)
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' Close logging
        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' Log Performance Data
        Dim VersionNum As String = "100"
        If debug <> "T" Then
            ' ============================================
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' Close other objects
        Try
            iDoc = Nothing
            resultsDeclare = Nothing
            resultsRoot = Nothing
            LoggingService = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Return results
        Return odoc
    End Function

    ' =================================================
    ' WEB SERVICE SUPPORT
    Public Function CallService(ByVal address As Uri, ByVal Debug As String) As String

        Dim request As Net.HttpWebRequest
        Dim response As Net.HttpWebResponse = Nothing
        Dim reader As StreamReader
        Dim sbSource As StringBuilder
        Dim results As String

        results = ""

        If address Is Nothing Then Throw New ArgumentNullException("address")
        Try
            ' Create and initialize the web request  
            request = DirectCast(Net.WebRequest.Create(address), Net.HttpWebRequest)
            request.UserAgent = ".NET"
            request.KeepAlive = False
            request.Timeout = 15 * 1000

            ' Get response  
            response = DirectCast(request.GetResponse(), Net.HttpWebResponse)

            If request.HaveResponse = True AndAlso Not (response Is Nothing) Then

                ' Get the response stream  
                reader = New StreamReader(response.GetResponseStream())

                ' Read it into a StringBuilder  
                sbSource = New StringBuilder(reader.ReadToEnd())

                ' Format response  
                results = sbSource.ToString()
            End If
        Catch wex As Net.WebException
            ' This exception will be raised if the server didn't return 200 - OK  
            ' Try to retrieve more information about the network error  
            If Not wex.Response Is Nothing Then
                Dim errorResponse As Net.HttpWebResponse = Nothing
                Try
                    errorResponse = DirectCast(wex.Response, Net.HttpWebResponse)
                Finally
                    If Not errorResponse Is Nothing Then errorResponse.Close()
                End Try
            End If
        Finally
            If Not response Is Nothing Then response.Close()
        End Try
        Return results
    End Function

    ' =================================================
    ' DATABASE
    Public Function ExecQuery(ByVal QType As String, ByVal QRec As String, ByVal cmd As SqlCommand, _
      ByVal SqlS As String, ByVal mydebuglog As ILog, ByVal Debug As String) As String
        Dim returnv As Integer
        Dim errmsg As String
        errmsg = ""
        If Debug = "Y" Then mydebuglog.Debug("  " & QType & " " & QRec & " record: " & SqlS)
        Try
            cmd.CommandText = SqlS
            returnv = cmd.ExecuteNonQuery()
            If returnv = 0 Then
                errmsg = errmsg & "The " & QRec & " record was not " & QType & vbCrLf
            End If
        Catch ex As Exception
            errmsg = errmsg & "Error " & QType & " record. " & ex.ToString & vbCrLf & "For query: " & SqlS & vbCrLf
        End Try
        Return errmsg
    End Function

    ' =================================================
    ' =================================================
    ' Additional Logging 
    Public Function LogTrainerDataChanges(ByVal con_ro As SqlConnection, ByVal TrainerNo As String, ByVal NewSqlS As String, ByVal OldSqlS As String, ByVal LogFileName As String, ByVal mydebuglog As ILog) As String
        'Get the old record values
        Dim dt As New Data.DataTable
        Using da As New SqlDataAdapter(OldSqlS, con_ro)
            da.Fill(dt)
        End Using
        Using writer As New StreamWriter(LogFileName, True)
            writer.WriteLine("-------------------------------------------")
            writer.WriteLine(Now.ToString & "; Trainer: " & TrainerNo)
            writer.WriteLine("---- Old Data Records -----------")
            For Each r As Data.DataRow In dt.Rows
                For Each col As Data.DataColumn In dt.Columns
                    writer.WriteLine("[" & col.ColumnName & "]: " & r(col).ToString())
                Next
                writer.WriteLine("------")
            Next
            writer.WriteLine("---------------------------------")
            writer.WriteLine("----- New Update Query-------------------")
            writer.WriteLine(NewSqlS)
            writer.WriteLine("-------------------------------------------")
        End Using
    End Function
    ' =================================================

    ' CALL SERVICE - STANDARDIZE ADDRESS
    Public Sub CallStandardizeAddress(ByRef ADDR_ID As String, ByRef ORG_ID As String, _
            ByRef CON_ID As String, ByRef ADDR_TYPE As String, ByRef GEOCODE As String, _
            ByRef ADDR As String, ByRef CITY As String, ByRef STATE As String, _
            ByRef COUNTY As String, ByRef ZIPCODE As String, ByRef COUNTRY As String, _
            ByRef ADDR_MATCH As String, ByRef LAT As String, ByRef LON As String, ByRef JURIS_ID As String, _
            ByRef DELIVERABLE As String, ByVal debug As String, ByRef mydebuglog As ILog, ByRef errmsg As String, _
            ByRef results As String, ByVal database As String)

        ' Declarations
        Dim wp, lresults As String
        Dim rDoc As XmlDocument
        Dim rNodeList As XmlNodeList
        Dim j As Integer
        lresults = ""

        ' This function calls the StandardizeAddress function and extracts the answers
        wp = "<AddressList><AddressRec>"
        wp = wp & "<Debug>" & debug & "</Debug>"
        wp = wp & "<Database>" & database & "</Database>"
        wp = wp & "<AddrId>" & HttpUtility.UrlEncode(ADDR_ID) & "</AddrId>"
        wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
        wp = wp & "<ConId>" & HttpUtility.UrlEncode(CON_ID) & "</ConId>"
        wp = wp & "<Type>" & HttpUtility.UrlEncode(ADDR_TYPE) & "</Type>"
        wp = wp & "<GeoCode>" & HttpUtility.UrlEncode(GEOCODE) & "</GeoCode>"
        wp = wp & "<Address>" & HttpUtility.UrlEncode(ADDR) & "</Address>"
        wp = wp & "<City>" & HttpUtility.UrlEncode(CITY) & "</City>"
        wp = wp & "<State>" & HttpUtility.UrlEncode(STATE) & "</State>"
        wp = wp & "<County>" & HttpUtility.UrlEncode(COUNTY) & "</County>"
        wp = wp & "<Zipcode>" & HttpUtility.UrlEncode(ZIPCODE) & "</Zipcode>"
        wp = wp & "<Country>" & HttpUtility.UrlEncode(COUNTRY) & "</Country>"
        wp = wp & "</AddressRec></AddressList>"
        Try
            If debug = "Y" Then mydebuglog.Debug("  Function CallStandardizeAddress===========" & vbCrLf & "  sXML: " & wp)
            rDoc = StandardizeAddress(wp)
            If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
            rNodeList = rDoc.SelectNodes("/AddressRec")
            For j = 0 To rNodeList.Count - 1
                Try
                    lresults = HttpUtility.UrlDecode(Trim(GetNodeValue("results", rNodeList.Item(j))))
                    If lresults <> "Failure" Then
                        If debug = "Y" Then mydebuglog.Debug("  Found node: " & j.ToString & "    " & Trim(GetNodeValue("Address", rNodeList.Item(j))))
                        ADDR = HttpUtility.UrlDecode(Trim(GetNodeValue("Address", rNodeList.Item(j))))
                        CITY = HttpUtility.UrlDecode(Trim(GetNodeValue("City", rNodeList.Item(j))))
                        STATE = HttpUtility.UrlDecode(Trim(GetNodeValue("State", rNodeList.Item(j))))
                        COUNTY = HttpUtility.UrlDecode(Trim(GetNodeValue("County", rNodeList.Item(j))))
                        ZIPCODE = HttpUtility.UrlDecode(Trim(GetNodeValue("Zipcode", rNodeList.Item(j))))
                        COUNTRY = HttpUtility.UrlDecode(Trim(GetNodeValue("Country", rNodeList.Item(j))))
                        ADDR_MATCH = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("MatchCode", rNodeList.Item(j)))))
                        LAT = HttpUtility.UrlDecode(Trim(GetNodeValue("Lat", rNodeList.Item(j))))
                        LON = HttpUtility.UrlDecode(Trim(GetNodeValue("Long", rNodeList.Item(j))))
                        JURIS_ID = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("JurisId", rNodeList.Item(j)))))
                        DELIVERABLE = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("Deliverable", rNodeList.Item(j)))))
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                    results = "Failure"
                End Try
            Next
            If debug = "Y" Then mydebuglog.Debug("  Standardized: " & results)
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
        End Try
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "  AFTER STANDARDIZATION====")
            mydebuglog.Debug("  ADDR: " & ADDR)
            mydebuglog.Debug("  CITY: " & CITY)
            mydebuglog.Debug("  STATE: " & STATE)
            mydebuglog.Debug("  COUNTY: " & COUNTY)
            mydebuglog.Debug("  ZIPCODE: " & ZIPCODE)
            mydebuglog.Debug("  COUNTRY: " & COUNTRY)
            mydebuglog.Debug("  ADDR_MATCH: " & ADDR_MATCH)
            mydebuglog.Debug("  LAT: " & LAT)
            mydebuglog.Debug("  LON: " & LON)
            mydebuglog.Debug("  JURIS_ID: " & JURIS_ID)
            mydebuglog.Debug("  DELIVERABLE: " & DELIVERABLE)
            mydebuglog.Debug("  =========================" & vbCrLf)
        End If
    End Sub

    ' =================================================
    ' CALL SERVICE - STANDARDIZE ORGANIZATION
    Public Sub CallStandardizeOrganization(ByRef ORG_ID As String, ByRef NAME As String, _
               ByRef LOC As String, ByRef ORG_MATCH As String, _
               ByVal debug As String, ByRef mydebuglog As ILog, ByRef errmsg As String, _
               ByRef results As String, ByVal database As String)

        ' Declarations
        Dim wp, lresults As String
        Dim rDoc As XmlDocument
        Dim oNodeList As XmlNodeList
        Dim i As Integer
        lresults = ""

        ' This function calls the StandardizeOrganization function and extracts the answers
        wp = "<Organizations><Organization>"
        wp = wp & "<Debug>" & debug & "</Debug>"
        wp = wp & "<Database>" & database & "</Database>"
        wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
        wp = wp & "<Name>" & HttpUtility.UrlEncode(NAME) & "</Name>"
        wp = wp & "<Loc>" & HttpUtility.UrlEncode(LOC) & "</Loc>"
        wp = wp & "<FullName> </FullName>"
        wp = wp & "</Organization></Organizations>"
        Try
            If debug = "Y" Then mydebuglog.Debug("  Function StandardizeOrganization===========" & vbCrLf & "  sXML: " & wp)
            rDoc = StandardizeOrganization(wp)
            If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
            oNodeList = rDoc.SelectNodes("/Organization")
            For i = 0 To oNodeList.Count - 1
                Try
                    lresults = HttpUtility.UrlDecode(GetNodeValue("results", oNodeList.Item(i)))
                    If lresults <> "Failure" Then
                        If debug = "Y" Then mydebuglog.Debug("  Found node: " & i.ToString & "    " & Trim(GetNodeValue("Name", oNodeList.Item(i))))
                        NAME = HttpUtility.UrlDecode(GetNodeValue("Name", oNodeList.Item(i)))
                        LOC = HttpUtility.UrlDecode(GetNodeValue("Loc", oNodeList.Item(i)))
                        ORG_ID = HttpUtility.UrlDecode(Trim(GetNodeValue("OrgId", oNodeList.Item(i))))
                        ORG_MATCH = HttpUtility.UrlDecode(Trim(GetNodeValue("MatchCode", oNodeList.Item(i))))
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                    results = "Failure"
                End Try
            Next
            If debug = "Y" Then mydebuglog.Debug("  Standardized: " & results)
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("  Unable to standardize: " & ex.Message)
        End Try
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "  AFTER STANDARDIZATION====")
            mydebuglog.Debug("  Name: " & NAME)
            mydebuglog.Debug("  Loc: " & LOC)
            mydebuglog.Debug("  OrgId: " & ORG_ID)
            mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
            mydebuglog.Debug("  =========================" & vbCrLf)
        End If
    End Sub

    ' =================================================
    ' CALL SERVICE - CLEAN ORGANIZATION
    Public Sub CallCleanOrganization(ByRef ORG_ID As String, ByRef NAME As String, _
            ByRef LOC As String, ByRef ORG_MATCH As String, ByRef ORG_PHONE As String, _
            ByRef ADDR_MATCH As String, ByVal ADDR_ID As String, ByVal WORK_PH_NUM As String, ByRef INDUSTRY As String, _
            ByVal debug As String, ByRef mydebuglog As ILog, ByRef errmsg As String, _
            ByRef results As String, ByVal database As String)

        ' Declarations
        Dim wp As String
        Dim rDoc As XmlDocument
        Dim oNodeList As XmlNodeList
        Dim i As Integer

        ' This function calls the CleanOrganization function and extracts the answers
        wp = "<Organizations><Organization>"
        wp = wp & "<Debug>" & debug & "</Debug>"
        wp = wp & "<Database>" & database & "</Database>"
        wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
        wp = wp & "<Name>" & HttpUtility.UrlEncode(NAME) & "</Name>"
        wp = wp & "<Loc>" & HttpUtility.UrlEncode(LOC) & "</Loc>"
        wp = wp & "<FullName>+</FullName>"
        wp = wp & "<OrgMatch><![CDATA[" & ORG_MATCH & "]]></OrgMatch>"
        wp = wp & "<OrgPhone>" & HttpUtility.UrlEncode(ORG_PHONE) & "</OrgPhone>"
        wp = wp & "<AddrMatch><![CDATA[" & ADDR_MATCH & "]]></AddrMatch>"
        wp = wp & "<AddrId>" & HttpUtility.UrlEncode(ADDR_ID) & "</AddrId>"
        wp = wp & "<WorkPhone>" & HttpUtility.UrlEncode(WORK_PH_NUM) & "</WorkPhone>"
        wp = wp & "<Industry>" & HttpUtility.UrlEncode(INDUSTRY) & "</Industry>"
        wp = wp & "<Confidence>3</Confidence>"  'Changed from 2 to 3 per Chris's request; Ren Hou; 04/07/17
        wp = wp & "</Organization></Organizations>"
        Try
            If debug = "Y" Then mydebuglog.Debug("  Function CleanOrganization===========" & vbCrLf & "  sXML: " & wp)
            rDoc = CleanOrganization(wp)
            If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
            oNodeList = rDoc.SelectNodes("/Organization")
            For i = 0 To oNodeList.Count - 1
                Try
                    '   <Organization>
                    '       <OrgId></OrgId>         - The Id of an existing organization, if applicable
                    '       <Name></Name>           - Name of organization
                    '       <Loc></Loc>             - Location of organization
                    '       <FullName></Fullname>   - Full name of organization
                    '       <OrgMatch></OrgMatch>   - Organization match code
                    '       <OrgPhone></OrgPhone>   - Organization main phone number
                    '       <AddrMatch></AddrMatch> - Address match code, if applicable
                    '       <AddrId></AddrId>       - Address Id of a related address, if applicable
                    '       <WorkPhone></WorkPhone> - Work phone
                    '       <Industry></Industry>   - Industry
                    '   </Organization>
                    If debug = "Y" Then mydebuglog.Debug("  Found node: " & i.ToString & "    " & Trim(GetNodeValue("Name", oNodeList.Item(i))))
                    ORG_ID = HttpUtility.UrlDecode(GetNodeValue("OrgId", oNodeList.Item(i)))
                    NAME = HttpUtility.UrlDecode(GetNodeValue("Name", oNodeList.Item(i)))
                    LOC = HttpUtility.UrlDecode(GetNodeValue("Loc", oNodeList.Item(i)))
                    ORG_MATCH = HttpUtility.UrlDecode(GetNodeValue("OrgMatch", oNodeList.Item(i)))
                    ORG_PHONE = HttpUtility.UrlDecode(GetNodeValue("OrgPhone", oNodeList.Item(i)))
                    If ORG_PHONE <> "" Then ORG_PHONE = StndPhone(ORG_PHONE)
                    ADDR_MATCH = HttpUtility.UrlDecode(GetNodeValue("AddrMatch", oNodeList.Item(i)))
                    ADDR_ID = HttpUtility.UrlDecode(GetNodeValue("AddrId", oNodeList.Item(i)))
                    WORK_PH_NUM = HttpUtility.UrlDecode(GetNodeValue("WorkPhone", oNodeList.Item(i)))
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                    results = "Failure"
                End Try
            Next
            If debug = "Y" Then mydebuglog.Debug("  Cleaned: " & results)
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("  Unable to clean: " & ex.Message)
        End Try
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "  AFTER STANDARDIZATION====")
            mydebuglog.Debug("  OrgId: " & ORG_ID)
            mydebuglog.Debug("  Name: " & NAME)
            mydebuglog.Debug("  Loc: " & LOC)
            mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
            mydebuglog.Debug("  OrgPhone: " & ORG_PHONE)
            mydebuglog.Debug("  AddrMatch: " & ADDR_MATCH)
            mydebuglog.Debug("  AddrId: " & ADDR_ID)
            mydebuglog.Debug("  WorkPhone: " & WORK_PH_NUM)
            mydebuglog.Debug("  =========================" & vbCrLf)
        End If
    End Sub

    ' =================================================
    ' CALL SERVICE - CLEAN CONTACT
    Public Sub CallCleanContact(ByRef CON_ID As String, ByRef PART_ID As String, _
         ByRef FST_NAME As String, ByRef LAST_NAME As String, ByRef MID_NAME As String, _
         ByRef GENDER As String, ByRef DOB As String, ByRef WORK_PH_NUM As String, _
         ByRef SSN As String, ByRef EMAIL_ADDR As String, ByRef ADDR_MATCH As String, _
         ByRef ADDR_TYPE As String, ByRef ORG_MATCH As String, ByRef ORG_PHONE As String, _
         ByRef REG_NUM As String, ByRef SUB_CON_ID As String, ByRef TRAINER_NO As String, _
         ByRef HOME_PHONE As String, ByRef FULL_NAME As String, ByRef Confidence As Integer, _
         ByRef CON_MATCH As String, ByRef ORG_ID As String, ByRef ADDR_ID As String, _
         ByRef JOB_TITLE As String, ByRef SOURCE As String, ByRef PER_TITLE As String, _
         ByVal debug As String, ByRef mydebuglog As ILog, ByRef errmsg As String, _
         ByRef results As String, ByVal database As String)

        ' Declarations
        Dim wp, lresults As String
        Dim rDoc As XmlDocument
        Dim oNodeList As XmlNodeList
        Dim i As Integer
        lresults = ""

        ' This function calls the CleanContact function and extracts the answers
        wp = "<Contacts><Contact>"
        wp = wp & "<Debug>" & debug & "</Debug>"
        wp = wp & "<Database>" & database & "</Database>"
        wp = wp & "<ConId>" & HttpUtility.UrlEncode(CON_ID) & "</ConId>"
        wp = wp & "<PartId>" & HttpUtility.UrlEncode(PART_ID) & "</PartId>"
        wp = wp & "<FirstName>" & HttpUtility.UrlEncode(FST_NAME) & "</FirstName>"
        wp = wp & "<MidName>" & HttpUtility.UrlEncode(MID_NAME) & "</MidName>"
        wp = wp & "<LastName>" & HttpUtility.UrlEncode(LAST_NAME) & "</LastName>"
        wp = wp & "<Gender>" & HttpUtility.UrlEncode(GENDER) & "</Gender>"
        wp = wp & "<FullName>" & HttpUtility.UrlEncode(FULL_NAME) & "</FullName>"
        wp = wp & "<DOB>" & HttpUtility.UrlEncode(DOB) & "</DOB>"
        wp = wp & "<WorkPhone>" & HttpUtility.UrlEncode(WORK_PH_NUM) & "</WorkPhone>"
        wp = wp & "<SSN>" & HttpUtility.UrlEncode(SSN) & "</SSN>"
        wp = wp & "<EmailAddr>" & HttpUtility.UrlEncode(EMAIL_ADDR) & "</EmailAddr>"
        wp = wp & "<AddrMatch><![CDATA[" & ADDR_MATCH & "]]></AddrMatch>"
        wp = wp & "<AddrType>" & ADDR_TYPE & "</AddrType>"
        wp = wp & "<OrgMatch><![CDATA[" & ORG_MATCH & "]]></OrgMatch>"
        wp = wp & "<OrgPhone>" & HttpUtility.UrlEncode(ORG_PHONE) & "</OrgPhone>"
        wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
        wp = wp & "<AddrId>" & HttpUtility.UrlEncode(ADDR_ID) & "</AddrId>"
        wp = wp & "<RegNum>" & HttpUtility.UrlEncode(REG_NUM) & "</RegNum>"
        wp = wp & "<SubConId>" & HttpUtility.UrlEncode(SUB_CON_ID) & "</SubConId>"
        wp = wp & "<TrainerNo>" & HttpUtility.UrlEncode(TRAINER_NO) & "</TrainerNo>"
        wp = wp & "<HomePhone>" & HttpUtility.UrlEncode(HOME_PHONE) & "</HomePhone>"
        wp = wp & "<JobTitle>" & HttpUtility.UrlEncode(JOB_TITLE) & "</JobTitle>"
        wp = wp & "<PerTitle>" & HttpUtility.UrlEncode(PER_TITLE) & "</PerTitle>"
        wp = wp & "<Source>" & HttpUtility.UrlEncode(SOURCE) & "</Source>"
        wp = wp & "<Confidence>" & Confidence.ToString & "</Confidence>"
        wp = wp & "</Contact></Contacts>"
        Try
            If debug = "Y" Then mydebuglog.Debug("  Function CallCleanContact===========" & vbCrLf & "  sXML: " & wp)
            rDoc = CleanContact(wp)
            If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
            oNodeList = rDoc.SelectNodes("/Contact")
            For i = 0 To oNodeList.Count - 1
                Try
                    lresults = HttpUtility.UrlDecode(Trim(GetNodeValue("results", oNodeList.Item(i))))
                    If lresults <> "Failure" Then
                        '   <Contact>
                        '       <FirstName>   
                        '       <MidName>
                        '       <LastName>
                        '       <FullName>
                        '       <Gender>
                        '       <MatchCode>
                        '       <ConId>
                        '       <PartId>
                        '       <DOB></DOB>             
                        '       <WorkPhone></WorkPhone> 
                        '       <SSN></SSN>             
                        '       <EmailAddr></EmailAddr> 
                        '       <RegNum>
                        '       <SubId>
                        '       <TrainerNo>
                        '       <HomePhone>
                        '       <JobTitle>
                        '       <PerTitle>
                        '       <Source>
                        '   </Contact>
                        FST_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("FirstName", oNodeList.Item(i))))
                        MID_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("MidName", oNodeList.Item(i))))
                        LAST_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("LastName", oNodeList.Item(i))))
                        FULL_NAME = HttpUtility.UrlDecode(Trim(GetNodeValue("FullName", oNodeList.Item(i))))
                        GENDER = HttpUtility.UrlDecode(GetNodeValue("Gender", oNodeList.Item(i)))
                        CON_MATCH = HttpUtility.UrlDecode(Trim(GetNodeValue("MatchCode", oNodeList.Item(i))))
                        CON_ID = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("ConId", oNodeList.Item(i)))))
                        PART_ID = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("PartId", oNodeList.Item(i)))))
                        DOB = HttpUtility.UrlDecode(GetNodeValue("DOB", oNodeList.Item(i)))
                        WORK_PH_NUM = HttpUtility.UrlDecode(GetNodeValue("WorkPhone", oNodeList.Item(i)))
                        If WORK_PH_NUM <> "" Then WORK_PH_NUM = StndPhone(WORK_PH_NUM)
                        SSN = HttpUtility.UrlDecode(GetNodeValue("SSN", oNodeList.Item(i)))
                        EMAIL_ADDR = HttpUtility.UrlDecode(GetNodeValue("EmailAddr", oNodeList.Item(i)))
                        REG_NUM = HttpUtility.UrlDecode(GetNodeValue("RegNum", oNodeList.Item(i)))
                        SUB_CON_ID = HttpUtility.UrlDecode(Trim(GetNodeValue("SubId", oNodeList.Item(i))))
                        TRAINER_NO = Trim(HttpUtility.UrlDecode(GetNodeValue("TrainerNo", oNodeList.Item(i)))) '6/27/21; Ren Hou; Added to trim Teainer_No;
                        HOME_PHONE = HttpUtility.UrlDecode(GetNodeValue("HomePhone", oNodeList.Item(i)))
                        If HOME_PHONE <> "" Then HOME_PHONE = StndPhone(HOME_PHONE)
                        JOB_TITLE = HttpUtility.UrlDecode(GetNodeValue("JobTitle", oNodeList.Item(i)))
                        PER_TITLE = HttpUtility.UrlDecode(GetNodeValue("PerTitle", oNodeList.Item(i)))
                        SOURCE = HttpUtility.UrlDecode(GetNodeValue("Source", oNodeList.Item(i)))
                    End If
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                    results = "Failure"
                End Try
            Next
            If debug = "Y" Then mydebuglog.Debug("  Cleaned: " & results & " - " & lresults)
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("  Unable to clean: " & ex.Message)
        End Try
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "  AFTER STANDARDIZATION====")
            mydebuglog.Debug("  FirstName: " & FST_NAME)
            mydebuglog.Debug("  MidName: " & MID_NAME)
            mydebuglog.Debug("  LastName: " & LAST_NAME)
            mydebuglog.Debug("  FullName: " & FULL_NAME)
            mydebuglog.Debug("  Gender: " & GENDER)
            mydebuglog.Debug("  ConMatch: " & CON_MATCH)
            mydebuglog.Debug("  ConId: " & CON_ID)
            mydebuglog.Debug("  AddrId: " & ADDR_ID)
            mydebuglog.Debug("  OrgId: " & ORG_ID)
            mydebuglog.Debug("  PartId: " & PART_ID)
            mydebuglog.Debug("  DOB: " & DOB)
            mydebuglog.Debug("  WorkPhone: " & WORK_PH_NUM)
            mydebuglog.Debug("  SSN: " & SSN)
            mydebuglog.Debug("  EmailAddr: " & EMAIL_ADDR)
            mydebuglog.Debug("  RegNum: " & REG_NUM)
            mydebuglog.Debug("  SubConId: " & SUB_CON_ID)
            mydebuglog.Debug("  TrainerNo: " & TRAINER_NO)
            mydebuglog.Debug("  HomePhone: " & HOME_PHONE)
            mydebuglog.Debug("  JobTitle: " & JOB_TITLE)
            mydebuglog.Debug("  PerTitle: " & PER_TITLE)
            mydebuglog.Debug("  Source: " & SOURCE)
            mydebuglog.Debug("  =========================" & vbCrLf)
        End If
    End Sub

    ' =================================================
    ' CALL SERVICE - CLEAN ADDRESS
    Public Sub CallCleanAddress(ByRef ADDR_ID As String, ByRef ORG_ID As String, _
            ByRef CON_ID As String, ByRef ADDR_TYPE As String, ByVal GEOCODE As String, _
            ByRef ADDR As String, ByRef CITY As String, ByRef STATE As String, _
            ByRef COUNTY As String, ByRef ZIPCODE As String, ByRef COUNTRY As String, _
            ByRef ADDR_MATCH As String, ByRef LAT As String, ByRef LON As String, _
            ByRef CON_MATCH As String, ByRef WORK_PH_NUM As String, _
            ByRef ORG_MATCH As String, ByRef ORG_PHONE As String, ByRef JURIS_ID As String, _
            ByVal debug As String, ByRef mydebuglog As ILog, ByRef errmsg As String, _
            ByRef results As String, ByVal database As String)

        ' Declarations
        Dim wp As String
        Dim rDoc As XmlDocument
        Dim rNodeList As XmlNodeList
        Dim j As Integer

        ' This function calls the CleanAddress function and extracts the answers
        wp = "<AddressList><AddressRec>"
        wp = wp & "<Debug>" & debug & "</Debug>"
        wp = wp & "<Database>" & database & "</Database>"
        wp = wp & "<Confidence>3</Confidence>"
        wp = wp & "<AddrId>" & HttpUtility.UrlEncode(ADDR_ID) & "</AddrId>"
        wp = wp & "<Type>" & ADDR_TYPE & "</Type>"
        wp = wp & "<GeoCode>" & GEOCODE & "</GeoCode>"
        wp = wp & "<JurisId>" & HttpUtility.UrlEncode(JURIS_ID) & "</JurisId>"
        wp = wp & "<Address>" & HttpUtility.UrlEncode(ADDR) & "</Address>"
        wp = wp & "<City>" & HttpUtility.UrlEncode(CITY) & "</City>"
        wp = wp & "<State>" & HttpUtility.UrlEncode(STATE) & "</State>"
        wp = wp & "<County>" & HttpUtility.UrlEncode(COUNTY) & "</County>"
        wp = wp & "<Zipcode>" & HttpUtility.UrlEncode(ZIPCODE) & "</Zipcode>"
        wp = wp & "<Country>" & HttpUtility.UrlEncode(COUNTRY) & "</Country>"
        wp = wp & "<ConId>" & HttpUtility.UrlEncode(CON_ID) & "</ConId>"
        wp = wp & "<ConMatch><![CDATA[" & CON_MATCH & "]]></ConMatch>"
        wp = wp & "<WorkPhone>" & HttpUtility.UrlEncode(WORK_PH_NUM) & "</WorkPhone>"
        wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
        wp = wp & "<OrgMatch><![CDATA[" & ORG_MATCH & "]]></OrgMatch>"
        wp = wp & "<OrgPhone>" & HttpUtility.UrlEncode(ORG_PHONE) & "</OrgPhone>"
        wp = wp & "</AddressRec></AddressList>"
        Try
            If debug = "Y" Then mydebuglog.Debug("  Function CleanAddress===========" & vbCrLf & "  sXML: " & wp)
            rDoc = CleanAddress(wp)
            If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
            rNodeList = rDoc.SelectNodes("/AddressRec")
            For j = 0 To rNodeList.Count - 1
                Try
                    '   <AddressRec>
                    '       <AddrId>   
                    '       <ConId>   
                    '       <OrgId>   
                    '       <JurisId>   
                    '       <MatchCode>
                    '       <Type>
                    '       <Street>        
                    '       <City>           
                    '       <State>          
                    '       <County>         
                    '       <Zipcode>         
                    '       <Country>         
                    '       <Lat>         
                    '       <Long>
                    '   </AddressRec>
                    If debug = "Y" Then mydebuglog.Debug("  Found node: " & j.ToString & "    " & Trim(GetNodeValue("Address", rNodeList.Item(j))))
                    ADDR_ID = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("AddrId", rNodeList.Item(j)))))
                    CON_ID = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("ConId", rNodeList.Item(j)))))
                    ORG_ID = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("OrgId", rNodeList.Item(j)))))
                    JURIS_ID = HttpUtility.UrlDecode(Trim(Trim(GetNodeValue("JurisId", rNodeList.Item(j)))))
                    ADDR_MATCH = HttpUtility.UrlDecode(Trim(GetNodeValue("MatchCode", rNodeList.Item(j))))
                    ADDR_TYPE = Trim(GetNodeValue("Type", rNodeList.Item(j)))
                    ADDR = HttpUtility.UrlDecode(Trim(GetNodeValue("Address", rNodeList.Item(j))))
                    CITY = HttpUtility.UrlDecode(Trim(GetNodeValue("City", rNodeList.Item(j))))
                    STATE = HttpUtility.UrlDecode(Trim(GetNodeValue("State", rNodeList.Item(j))))
                    COUNTY = HttpUtility.UrlDecode(Trim(GetNodeValue("County", rNodeList.Item(j))))
                    ZIPCODE = HttpUtility.UrlDecode(Trim(GetNodeValue("Zipcode", rNodeList.Item(j))))
                    COUNTRY = HttpUtility.UrlDecode(Trim(GetNodeValue("Country", rNodeList.Item(j))))
                    LAT = HttpUtility.UrlDecode(Trim(GetNodeValue("Lat", rNodeList.Item(j))))
                    LON = HttpUtility.UrlDecode(Trim(GetNodeValue("Long", rNodeList.Item(j))))
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                    results = "Failure"
                End Try
            Next
            If debug = "Y" Then mydebuglog.Debug("  Cleaned: " & results)
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("  Unable to clean: " & ex.Message)
        End Try
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "  CallCleanAddress AFTER STANDARDIZATION====")
            mydebuglog.Debug("  AddrId: " & ADDR_ID)
            mydebuglog.Debug("  ConId: " & CON_ID)
            mydebuglog.Debug("  OrgId: " & ORG_ID)
            mydebuglog.Debug("  JurisId: " & JURIS_ID)
            mydebuglog.Debug("  MatchCode: " & ADDR_MATCH)
            mydebuglog.Debug("  Type: " & ADDR_TYPE)
            mydebuglog.Debug("  Address: " & ADDR)
            mydebuglog.Debug("  City: " & CITY)
            mydebuglog.Debug("  State: " & STATE)
            mydebuglog.Debug("  County: " & COUNTY)
            mydebuglog.Debug("  Zipcode: " & ZIPCODE)
            mydebuglog.Debug("  Country: " & COUNTRY)
            mydebuglog.Debug("  Lat: " & LAT)
            mydebuglog.Debug("  Long: " & LON)
            mydebuglog.Debug("  =========================" & vbCrLf)
        End If

    End Sub

    ' =================================================
    ' CALL SERVICE - MATCH RECORD
    <WebMethod(Description:="Look for a matching record to the one provided")> _
    Public Function MatchRecord(ByVal CON_ID As String, ByVal PART_ID As String, _
            ByVal FST_NAME As String, ByVal LAST_NAME As String, ByVal MID_NAME As String, _
            ByVal GENDER As String, ByVal DOB As String, ByVal WORK_PH_NUM As String, _
            ByVal SSN As String, ByVal EMAIL_ADDR As String, ByVal ADDR_MATCH As String, _
            ByVal ADDR_TYPE As String, ByVal ORG_MATCH As String, ByVal ORG_PHONE As String, _
            ByVal REG_NUM As String, ByVal SUB_CON_ID As String, ByVal TRAINER_NO As String, _
            ByVal HOME_PHONE As String, ByVal Confidence As String, _
            ByVal GEOCODE As String, ByVal CON_MATCH As String, ByVal LON As String, _
            ByVal ADDR As String, ByVal CITY As String, ByVal STATE As String, _
            ByVal COUNTY As String, ByVal ZIPCODE As String, ByVal COUNTRY As String, _
            ByVal JURIS_ID As String, ByVal LOC As String, ByVal LAT As String, _
            ByVal ADDR_ID As String, ByVal ORG_ID As String, ByVal NAME As String, _
            ByVal JOB_TITLE As String, ByVal SOURCE As String, ByVal INDUSTRY As String, _
            ByVal PER_TITLE As String, _
            ByVal debug As String, ByVal database As String, ByVal Encode As String) As XmlDocument

        ' This function is a wrapper service for calling the CleanRecord service

        ' Declarations
        Dim wp, mypath, errmsg As String
        Dim rDoc As XmlDocument
        Dim oNodeList As XmlNodeList
        Dim i As Integer
        mypath = HttpRuntime.AppDomainAppPath

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("MRDebugLog")
        Dim logfile As String
        Dim LogStartTime As String = Now.ToString

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' Process parameters
        If Not IsNumeric(Confidence) Or Confidence = "" Then Confidence = "3"
        If Trim(database) = "" Or (database <> "U" And database <> "C") Then database = "X"
        If Trim(Encode = "") Or (Encode <> "Y" And Encode <> "N") Then Encode = "Y"
        errmsg = ""
        debug = "Y"

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Then
            logfile = "C:\Logs\MatchRecord.log"
            Try
                log4net.GlobalContext.Properties("MRLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                GoTo CloseOut
            End Try

            If debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters:")
                mydebuglog.Debug("  database: " & database)
                mydebuglog.Debug("  Encode: " & Encode)
                mydebuglog.Debug("  Confidence: " & Confidence)
            End If
        End If

        ' ============================================
        ' Call the CleanRecord function and extracts the answers
        wp = "<Records><Record>"
        wp = wp & "<Debug>" & debug & "</Debug>"
        wp = wp & "<Database>" & database & "</Database>"
        wp = wp & "<Confidence>" & Confidence & "</Confidence>"
        wp = wp & "<AddrId>" & HttpUtility.UrlEncode(ADDR_ID) & "</AddrId>"
        wp = wp & "<AddrType>" & ADDR_TYPE & "</AddrType>"
        wp = wp & "<AddrMatch><![CDATA[" & ADDR_MATCH & "]]></AddrMatch>"
        wp = wp & "<GeoCode>" & GEOCODE & "</GeoCode>"
        wp = wp & "<JurisId>" & HttpUtility.UrlEncode(JURIS_ID) & "</JurisId>"
        wp = wp & "<Address>" & HttpUtility.UrlEncode(ADDR) & "</Address>"
        wp = wp & "<City>" & HttpUtility.UrlEncode(CITY) & "</City>"
        wp = wp & "<State>" & HttpUtility.UrlEncode(STATE) & "</State>"
        wp = wp & "<County>" & HttpUtility.UrlEncode(COUNTY) & "</County>"
        wp = wp & "<Zipcode>" & HttpUtility.UrlEncode(ZIPCODE) & "</Zipcode>"
        wp = wp & "<Country>" & HttpUtility.UrlEncode(COUNTRY) & "</Country>"
        wp = wp & "<ConId>" & HttpUtility.UrlEncode(CON_ID) & "</ConId>"
        wp = wp & "<ConMatch><![CDATA[" & CON_MATCH & "]]></ConMatch>"
        wp = wp & "<PartId>" & HttpUtility.UrlEncode(PART_ID) & "</PartId>"
        wp = wp & "<FirstName>" & HttpUtility.UrlEncode(FST_NAME) & "</FirstName>"
        wp = wp & "<MidName>" & HttpUtility.UrlEncode(MID_NAME) & "</MidName>"
        wp = wp & "<LastName>" & HttpUtility.UrlEncode(LAST_NAME) & "</LastName>"
        wp = wp & "<Gender>" & HttpUtility.UrlEncode(GENDER) & "</Gender>"
        wp = wp & "<DOB>" & HttpUtility.UrlEncode(DOB) & "</DOB>"
        wp = wp & "<WorkPhone>" & HttpUtility.UrlEncode(WORK_PH_NUM) & "</WorkPhone>"
        wp = wp & "<SSN>" & HttpUtility.UrlEncode(SSN) & "</SSN>"
        wp = wp & "<EmailAddr>" & HttpUtility.UrlEncode(EMAIL_ADDR) & "</EmailAddr>"
        wp = wp & "<RegNum>" & HttpUtility.UrlEncode(REG_NUM) & "</RegNum>"
        wp = wp & "<SubConId>" & HttpUtility.UrlEncode(SUB_CON_ID) & "</SubConId>"
        wp = wp & "<HomePhone>" & HttpUtility.UrlEncode(HOME_PHONE) & "</HomePhone>"
        wp = wp & "<TrainerNo>" & HttpUtility.UrlEncode(TRAINER_NO) & "</TrainerNo>"
        wp = wp & "<Name>" & HttpUtility.UrlEncode(NAME) & "</Name>"
        wp = wp & "<Loc>" & HttpUtility.UrlEncode(LOC) & "</Loc>"
        wp = wp & "<OrgId>" & HttpUtility.UrlEncode(ORG_ID) & "</OrgId>"
        wp = wp & "<OrgMatch><![CDATA[" & ORG_MATCH & "]]></OrgMatch>"
        wp = wp & "<OrgPhone>" & HttpUtility.UrlEncode(ORG_PHONE) & "</OrgPhone>"
        wp = wp & "<JobTitle>" & HttpUtility.UrlEncode(JOB_TITLE) & "</JobTitle>"
        wp = wp & "<PerTitle>" & HttpUtility.UrlEncode(PER_TITLE) & "</PerTitle>"
        wp = wp & "<Source>" & HttpUtility.UrlEncode(SOURCE) & "</Source>"
        wp = wp & "<Industry>" & HttpUtility.UrlEncode(INDUSTRY) & "</Industry>"
        wp = wp & "</Record></Records>"
        Try
            If debug = "Y" Then mydebuglog.Debug("  Calling CleanRecord===========" & vbCrLf & vbCrLf & "wp: " & wp & vbCrLf)
            Try
                rDoc = CleanRecord(wp)
            Catch ex As Exception
                If debug = "Y" Then mydebuglog.Debug("  Error reported: " & ex.Message & vbCrLf)
            End Try
            If debug = "Y" Then mydebuglog.Debug("  Results: " & rDoc.InnerText)
            oNodeList = rDoc.SelectNodes("/Record")
            For i = 0 To oNodeList.Count - 1
                Try
                    If debug = "Y" Then mydebuglog.Debug("  Found node: " & i.ToString & "    " & Trim(GetNodeValue("LastName", oNodeList.Item(i))))
                    ADDR_ID = Trim(GetNodeValue("AddrId", oNodeList.Item(i)))
                    ADDR_TYPE = Trim(GetNodeValue("AddrType", oNodeList.Item(i)))
                    ADDR_MATCH = Trim(GetNodeValue("AddrMatch", oNodeList.Item(i)))
                    JURIS_ID = Trim(GetNodeValue("JurisId", oNodeList.Item(i)))
                    ADDR = Trim(GetNodeValue("Address", oNodeList.Item(i)))
                    CITY = Trim(GetNodeValue("City", oNodeList.Item(i)))
                    STATE = Trim(GetNodeValue("State", oNodeList.Item(i)))
                    COUNTY = Trim(GetNodeValue("County", oNodeList.Item(i)))
                    ZIPCODE = Trim(GetNodeValue("Zipcode", oNodeList.Item(i)))
                    COUNTRY = Trim(GetNodeValue("Country", oNodeList.Item(i)))
                    LAT = Trim(GetNodeValue("Lat", oNodeList.Item(i)))
                    LON = Trim(GetNodeValue("Long", oNodeList.Item(i)))
                    CON_ID = Trim(GetNodeValue("ConId", oNodeList.Item(i)))
                    CON_MATCH = Trim(GetNodeValue("ConMatch", oNodeList.Item(i)))
                    PART_ID = Trim(GetNodeValue("PartId", oNodeList.Item(i)))
                    FST_NAME = Trim(GetNodeValue("FirstName", oNodeList.Item(i)))
                    MID_NAME = Trim(GetNodeValue("MidName", oNodeList.Item(i)))
                    LAST_NAME = Trim(GetNodeValue("LastName", oNodeList.Item(i)))
                    GENDER = GetNodeValue("Gender", oNodeList.Item(i))
                    DOB = GetNodeValue("DOB", oNodeList.Item(i))
                    WORK_PH_NUM = GetNodeValue("WorkPhone", oNodeList.Item(i))
                    If Len(WORK_PH_NUM) > 0 Then WORK_PH_NUM = StndPhone(WORK_PH_NUM)
                    SSN = GetNodeValue("SSN", oNodeList.Item(i))
                    EMAIL_ADDR = GetNodeValue("EmailAddr", oNodeList.Item(i))
                    REG_NUM = GetNodeValue("RegNum", oNodeList.Item(i))
                    SUB_CON_ID = GetNodeValue("SubConId", oNodeList.Item(i))
                    HOME_PHONE = GetNodeValue("HomePhone", oNodeList.Item(i))
                    If Len(HOME_PHONE) > 0 Then HOME_PHONE = StndPhone(HOME_PHONE)
                    TRAINER_NO = GetNodeValue("TrainerNo", oNodeList.Item(i))
                    NAME = GetNodeValue("Name", oNodeList.Item(i))
                    LOC = GetNodeValue("Loc", oNodeList.Item(i))
                    ORG_ID = GetNodeValue("OrgId", oNodeList.Item(i))
                    ORG_MATCH = GetNodeValue("OrgMatch", oNodeList.Item(i))
                    ORG_PHONE = GetNodeValue("OrgPhone", oNodeList.Item(i))
                    If Len(ORG_PHONE) > 0 Then ORG_PHONE = StndPhone(ORG_PHONE)
                    JOB_TITLE = GetNodeValue("JobTitle", oNodeList.Item(i))
                    PER_TITLE = GetNodeValue("PerTitle", oNodeList.Item(i))
                    SOURCE = GetNodeValue("Source", oNodeList.Item(i))
                    INDUSTRY = GetNodeValue("Industry", oNodeList.Item(i))
                Catch ex As Exception
                    errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                End Try
            Next
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug("  Unable to clean: " & ex.Message)
        End Try

        ' Debug output
        If debug = "Y" Then
            mydebuglog.Debug(vbCrLf & "  AFTER CLEANRECORD====")
            mydebuglog.Debug("  AddrId: " & ADDR_ID)
            mydebuglog.Debug("  AddrMatch: " & ADDR_MATCH)
            mydebuglog.Debug("  AddrType: " & ADDR_TYPE)
            mydebuglog.Debug("  JurisId: " & JURIS_ID)
            mydebuglog.Debug("  Address: " & ADDR)
            mydebuglog.Debug("  City: " & CITY)
            mydebuglog.Debug("  State: " & STATE)
            mydebuglog.Debug("  County: " & COUNTY)
            mydebuglog.Debug("  Zipcode: " & ZIPCODE)
            mydebuglog.Debug("  Country: " & COUNTRY)
            mydebuglog.Debug("  Lat: " & LAT)
            mydebuglog.Debug("  Long: " & LON)
            mydebuglog.Debug("  ConId: " & CON_ID)
            mydebuglog.Debug("  ConMatch: " & CON_MATCH)
            mydebuglog.Debug("  PartId: " & PART_ID)
            mydebuglog.Debug("  FirstName: " & FST_NAME)
            mydebuglog.Debug("  MidName: " & MID_NAME)
            mydebuglog.Debug("  LastName: " & LAST_NAME)
            mydebuglog.Debug("  Gender: " & GENDER)
            mydebuglog.Debug("  DOB: " & DOB)
            mydebuglog.Debug("  WorkPhone: " & WORK_PH_NUM)
            mydebuglog.Debug("  SSN: " & SSN)
            mydebuglog.Debug("  EmailAddr: " & EMAIL_ADDR)
            mydebuglog.Debug("  RegNum: " & REG_NUM)
            mydebuglog.Debug("  SubConId: " & SUB_CON_ID)
            mydebuglog.Debug("  HomePhone: " & HOME_PHONE)
            mydebuglog.Debug("  TrainerNo: " & TRAINER_NO)
            mydebuglog.Debug("  JobTitle: " & JOB_TITLE)
            mydebuglog.Debug("  PerTitle: " & PER_TITLE)
            mydebuglog.Debug("  OrgId: " & ORG_ID)
            mydebuglog.Debug("  Name: " & NAME)
            mydebuglog.Debug("  Loc: " & LOC)
            mydebuglog.Debug("  OrgMatch: " & ORG_MATCH)
            mydebuglog.Debug("  OrgPhone: " & ORG_PHONE)
            mydebuglog.Debug("  Source: " & SOURCE)
            mydebuglog.Debug("  Industry: " & INDUSTRY)
            mydebuglog.Debug("  =========================" & vbCrLf)
        End If

CloseOut:
        ' ============================================
        ' Return the standardized information as an XML document:
        '	<Record>
        '	<AddrId></AddrId>       - Address Id of a related address, if applicable
        '	<AddrType></AddrType>	- "P"ersonal or "O"rganization address
        '	<AddrMatch></AddrMatch> - Address match code
        '	<JurisId></JurisId>		- Jurisdiction Id
        '	<Address></Address>		- Street Address
        '	<City></City>			- City
        '	<State></State>			- State or province
        '	<County></County>		- County or region
        '	<Zipcode></Zipcode>		- Zipcode or postal code
        '	<Country></Country>		- Country
        '   <Lat></Lat>             - Latitude
        '   <Long></Long>           - Longitude
        '	<ConId></ConId>			- Contact Id
        '	<ConMatch></ConMatch>	- Contact match code
        '	<PartId></PartId>		- Participant Id
        '	<FirstName></FirstName> - First name of contact
        '	<MidName></MidName>     - Middle name of contact
        '	<LastName></LastName>   - Last name of contact
        '	<Gender></Gender>       - Gender of contact
        '	<DOB></DOB>             - Date of Birth
        '	<WorkPhone></WorkPhone> - Work phone of a contact at the organization
        '	<SSN></SSN>             - Social Security Number
        '	<EmailAddr></EmailAddr> - Email address
        '	<RegNum></RegNum>       - Web registration id
        '	<SubConId></SubConId>   - Subscription contact record
        '	<HomePhone></HomePhone> - Home phone number
        '	<TrainerNo></TrainerNo> - Trainer number 
        '	<Name></Name>           - Name of organization
        '	<Loc></Loc>             - Location of organization
        '	<OrgId></OrgId>			- Organization Id 
        '	<OrgMatch></OrgMatch>   - Organization match code
        '	<OrgPhone></OrgPhone>   - Organization main phone number
        '   <JobTitle></JobTitle>   - Job Title
        '   <PerTitle></PerTitle>   - Personal Title
        '   <Source></Source>       - Interest source
        '   <Industry></Industry>   - Organization industry
        '	</Record>

        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)
        resultsRoot = odoc.CreateElement("Record")
        odoc.InsertAfter(resultsRoot, resultsDeclare)

        Try
            ' Add result items - send what was submitted for debugging purposes 
            If debug <> "T" Then
                If Encode = "Y" Then
                    AddXMLChild(odoc, resultsRoot, "AddrId", IIf(ADDR_ID = "", " ", ADDR_ID))
                    AddXMLChild(odoc, resultsRoot, "AddrType", IIf(ADDR_TYPE = "", " ", ADDR_TYPE))
                    AddXMLChild(odoc, resultsRoot, "AddrMatch", IIf(ADDR_MATCH = "", " ", ADDR_MATCH))
                    AddXMLChild(odoc, resultsRoot, "JurisId", IIf(JURIS_ID = "", " ", JURIS_ID))
                    AddXMLChild(odoc, resultsRoot, "Address", IIf(ADDR = "", " ", ADDR))
                    AddXMLChild(odoc, resultsRoot, "City", IIf(CITY = "", " ", CITY))
                    AddXMLChild(odoc, resultsRoot, "State", IIf(STATE = "", " ", STATE))
                    AddXMLChild(odoc, resultsRoot, "County", IIf(COUNTY = "", " ", COUNTY))
                    AddXMLChild(odoc, resultsRoot, "Zipcode", IIf(ZIPCODE = "", " ", ZIPCODE))
                    AddXMLChild(odoc, resultsRoot, "Country", IIf(COUNTRY = "", " ", COUNTRY))
                    AddXMLChild(odoc, resultsRoot, "Lat", IIf(LAT = "", " ", LAT))
                    AddXMLChild(odoc, resultsRoot, "Long", IIf(LON = "", " ", LON))
                    AddXMLChild(odoc, resultsRoot, "ConId", IIf(CON_ID = "", " ", CON_ID))
                    AddXMLChild(odoc, resultsRoot, "ConMatch", IIf(CON_MATCH = "", " ", CON_MATCH))
                    AddXMLChild(odoc, resultsRoot, "PartId", IIf(PART_ID = "", " ", PART_ID))
                    AddXMLChild(odoc, resultsRoot, "FirstName", IIf(FST_NAME = "", " ", FST_NAME))
                    AddXMLChild(odoc, resultsRoot, "MidName", IIf(MID_NAME = "", " ", MID_NAME))
                    AddXMLChild(odoc, resultsRoot, "LastName", IIf(LAST_NAME = "", " ", LAST_NAME))
                    AddXMLChild(odoc, resultsRoot, "Gender", IIf(GENDER = "", " ", GENDER))
                    AddXMLChild(odoc, resultsRoot, "DOB", IIf(DOB = "", " ", DOB))
                    AddXMLChild(odoc, resultsRoot, "WorkPhone", IIf(WORK_PH_NUM = "", " ", WORK_PH_NUM))
                    AddXMLChild(odoc, resultsRoot, "SSN", IIf(SSN = "", " ", SSN))
                    AddXMLChild(odoc, resultsRoot, "EmailAddr", IIf(EMAIL_ADDR = "", " ", EMAIL_ADDR))
                    AddXMLChild(odoc, resultsRoot, "RegNum", IIf(REG_NUM = "", " ", REG_NUM))
                    AddXMLChild(odoc, resultsRoot, "SubConId", IIf(SUB_CON_ID = "", " ", SUB_CON_ID))
                    AddXMLChild(odoc, resultsRoot, "HomePhone", IIf(HOME_PHONE = "", " ", HOME_PHONE))
                    AddXMLChild(odoc, resultsRoot, "TrainerNo", IIf(TRAINER_NO = "", " ", TRAINER_NO))
                    AddXMLChild(odoc, resultsRoot, "Name", IIf(NAME = "", " ", NAME))
                    AddXMLChild(odoc, resultsRoot, "Loc", IIf(LOC = "", " ", LOC))
                    AddXMLChild(odoc, resultsRoot, "OrgId", IIf(ORG_ID = "", " ", ORG_ID))
                    AddXMLChild(odoc, resultsRoot, "OrgMatch", IIf(ORG_MATCH = "", " ", ORG_MATCH))
                    AddXMLChild(odoc, resultsRoot, "OrgPhone", IIf(ORG_PHONE = "", " ", ORG_PHONE))
                    AddXMLChild(odoc, resultsRoot, "JobTitle", IIf(JOB_TITLE = "", " ", JOB_TITLE))
                    AddXMLChild(odoc, resultsRoot, "PerTitle", IIf(PER_TITLE = "", " ", PER_TITLE))
                    AddXMLChild(odoc, resultsRoot, "Source", IIf(SOURCE = "", " ", SOURCE))
                    AddXMLChild(odoc, resultsRoot, "Industry", IIf(INDUSTRY = "", " ", INDUSTRY))
                Else
                    AddXMLChild(odoc, resultsRoot, "AddrId", IIf(ADDR_ID = "", " ", Trim(HttpUtility.UrlDecode(ADDR_ID))))
                    AddXMLChild(odoc, resultsRoot, "AddrType", IIf(ADDR_TYPE = "", " ", Trim(HttpUtility.UrlDecode(ADDR_TYPE))))
                    AddXMLChild(odoc, resultsRoot, "AddrMatch", IIf(ADDR_MATCH = "", " ", Trim(HttpUtility.UrlDecode(ADDR_MATCH))))
                    AddXMLChild(odoc, resultsRoot, "JurisId", IIf(JURIS_ID = "", " ", Trim(HttpUtility.UrlDecode(JURIS_ID))))
                    AddXMLChild(odoc, resultsRoot, "Address", IIf(ADDR = "", " ", Trim(HttpUtility.UrlDecode(ADDR))))
                    AddXMLChild(odoc, resultsRoot, "City", IIf(CITY = "", " ", Trim(HttpUtility.UrlDecode(CITY))))
                    AddXMLChild(odoc, resultsRoot, "State", IIf(STATE = "", " ", Trim(HttpUtility.UrlDecode(STATE))))
                    AddXMLChild(odoc, resultsRoot, "County", IIf(COUNTY = "", " ", Trim(HttpUtility.UrlDecode(COUNTY))))
                    AddXMLChild(odoc, resultsRoot, "Zipcode", IIf(ZIPCODE = "", " ", Trim(HttpUtility.UrlDecode(ZIPCODE))))
                    AddXMLChild(odoc, resultsRoot, "Country", IIf(COUNTRY = "", " ", Trim(HttpUtility.UrlDecode(COUNTRY))))
                    AddXMLChild(odoc, resultsRoot, "Lat", IIf(LAT = "", " ", Trim(HttpUtility.UrlDecode(LAT))))
                    AddXMLChild(odoc, resultsRoot, "Long", IIf(LON = "", " ", Trim(HttpUtility.UrlDecode(LON))))
                    AddXMLChild(odoc, resultsRoot, "ConId", IIf(CON_ID = "", " ", Trim(HttpUtility.UrlDecode(CON_ID))))
                    AddXMLChild(odoc, resultsRoot, "ConMatch", IIf(CON_MATCH = "", " ", Trim(HttpUtility.UrlDecode(CON_MATCH))))
                    AddXMLChild(odoc, resultsRoot, "PartId", IIf(PART_ID = "", " ", Trim(HttpUtility.UrlDecode(PART_ID))))
                    AddXMLChild(odoc, resultsRoot, "FirstName", IIf(FST_NAME = "", " ", Trim(HttpUtility.UrlDecode(FST_NAME))))
                    AddXMLChild(odoc, resultsRoot, "MidName", IIf(MID_NAME = "", " ", Trim(HttpUtility.UrlDecode(MID_NAME))))
                    AddXMLChild(odoc, resultsRoot, "LastName", IIf(LAST_NAME = "", " ", Trim(HttpUtility.UrlDecode(LAST_NAME))))
                    AddXMLChild(odoc, resultsRoot, "Gender", IIf(GENDER = "", " ", Trim(HttpUtility.UrlDecode(GENDER))))
                    AddXMLChild(odoc, resultsRoot, "DOB", IIf(DOB = "", " ", Trim(HttpUtility.UrlDecode(DOB))))
                    AddXMLChild(odoc, resultsRoot, "WorkPhone", IIf(WORK_PH_NUM = "", " ", Trim(HttpUtility.UrlDecode(WORK_PH_NUM))))
                    AddXMLChild(odoc, resultsRoot, "SSN", IIf(SSN = "", " ", Trim(HttpUtility.UrlDecode(SSN))))
                    AddXMLChild(odoc, resultsRoot, "EmailAddr", IIf(EMAIL_ADDR = "", " ", Trim(HttpUtility.UrlDecode(EMAIL_ADDR))))
                    AddXMLChild(odoc, resultsRoot, "RegNum", IIf(REG_NUM = "", " ", Trim(HttpUtility.UrlDecode(REG_NUM))))
                    AddXMLChild(odoc, resultsRoot, "SubConId", IIf(SUB_CON_ID = "", " ", Trim(HttpUtility.UrlDecode(SUB_CON_ID))))
                    AddXMLChild(odoc, resultsRoot, "HomePhone", IIf(HOME_PHONE = "", " ", Trim(HttpUtility.UrlDecode(HOME_PHONE))))
                    AddXMLChild(odoc, resultsRoot, "TrainerNo", IIf(TRAINER_NO = "", " ", Trim(HttpUtility.UrlDecode(TRAINER_NO))))
                    AddXMLChild(odoc, resultsRoot, "Name", IIf(NAME = "", " ", Trim(HttpUtility.UrlDecode(NAME))))
                    AddXMLChild(odoc, resultsRoot, "Loc", IIf(LOC = "", " ", Trim(HttpUtility.UrlDecode(LOC))))
                    AddXMLChild(odoc, resultsRoot, "OrgId", IIf(ORG_ID = "", " ", Trim(HttpUtility.UrlDecode(ORG_ID))))
                    AddXMLChild(odoc, resultsRoot, "OrgMatch", IIf(ORG_MATCH = "", " ", Trim(HttpUtility.UrlDecode(ORG_MATCH))))
                    AddXMLChild(odoc, resultsRoot, "OrgPhone", IIf(ORG_PHONE = "", " ", Trim(HttpUtility.UrlDecode(ORG_PHONE))))
                    AddXMLChild(odoc, resultsRoot, "JobTitle", IIf(JOB_TITLE = "", " ", Trim(HttpUtility.UrlDecode(JOB_TITLE))))
                    AddXMLChild(odoc, resultsRoot, "PerTitle", IIf(PER_TITLE = "", " ", Trim(HttpUtility.UrlDecode(PER_TITLE))))
                    AddXMLChild(odoc, resultsRoot, "Source", IIf(SOURCE = "", " ", Trim(HttpUtility.UrlDecode(SOURCE))))
                    AddXMLChild(odoc, resultsRoot, "Industry", IIf(INDUSTRY = "", " ", Trim(HttpUtility.UrlDecode(INDUSTRY))))
                End If
            End If
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))

        Catch ex As Exception
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")

        End Try

        ' ============================================
        ' Close the log file if any
        If Trim(errmsg) <> "" Then myeventlog.Error("MacthRecord : Error: " & Trim(errmsg))
        If debug <> "T" Then myeventlog.Info("MatchRecord : Attempted for ADDR_ID=" & ADDR_ID & ", CON_ID=" & CON_ID & ", ORG_ID=" & ORG_ID)
        If debug = "Y" Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "  Error: " & Trim(errmsg))
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Log Performance Data
        Dim VersionNum As String = "100"
        If debug <> "T" Then
            ' Send the web request
            Try
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        Return odoc
    End Function

    ' =================================================
    ' CALL SERVICE - SEND EMAIL
    Public Function PrepareMail(ByVal FromEmail As String, ByVal ToEmail As String, ByVal Subject As String, _
        ByVal Body As String, ByVal Debug As String, ByRef mydebuglog As ILog) As Boolean
        ' This function wraps message info into the XML necessary to call the SendMail web service function.
        ' This is used by other services executing from this application.
        ' Assumptions:  Create a record in MESSAGES and IDs are unknown 
        Dim wp As String

        ' Web service declarations
        Dim EmailService As New com.certegrity.cloudsvc.Service

        wp = "<EMailMessageList><EMailMessage>"
        wp = wp & "<debug>" & Debug & "</debug>"
        wp = wp & "<database>C</database>"
        wp = wp & "<Id> </Id>"
        wp = wp & "<SourceId></SourceId>"
        wp = wp & "<From>" & FromEmail & "</From>"
        wp = wp & "<FromId></FromId>"
        wp = wp & "<FromName></FromName>"
        wp = wp & "<To>" & ToEmail & "</To>"
        wp = wp & "<ToId></ToId>"
        wp = wp & "<Cc></Cc>"
        wp = wp & "<Bcc></Bcc>"
        wp = wp & "<ReplyTo></ReplyTo>"
        wp = wp & "<Subject>" & Subject & "</Subject>"
        wp = wp & "<Body>" & Body & "</Body>"
        wp = wp & "<Format></Format>"
        wp = wp & "</EMailMessage></EMailMessageList>"
        If Debug = "Y" Then mydebuglog.Debug("Email XML: " & wp)
        PrepareMail = EmailService.SendMail(wp)

    End Function

    ' =================================================
    ' NUMERIC
    Public Function Round(ByVal nValue As Double, ByVal nDigits As Integer) As Double
        Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
    End Function

    ' =================================================
    ' XML DOCUMENT MANAGEMENT
    Private Sub AddXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
        root.AppendChild(resultsItem)
    End Sub

    Private Function GetNodeValue(ByVal sNodeName As String, ByVal oParentNode As XmlNode) As String
        ' Generic function to return the value of a node in an XML document
        Dim oNode As XmlNode = oParentNode.SelectSingleNode(".//" + sNodeName)
        If oNode Is Nothing Then
            Return String.Empty
        Else
            Return oNode.InnerText
        End If
    End Function
    Private Function IsAddressVerifiedBefore(ByRef cmd As SqlCommand, addr_type As String, matchcode As String, num_day_before As String) As Boolean
        'Check siebeldb
        Dim sql As String
        Dim cnt As String
        sql = "SELECT count(1) from siebeldb.dbo." + If(addr_type = "O", "S_ADDR_ORG ", "S_ADDR_PER ") + " WITH (NOLOCK) WHERE X_MATCH_CD='" + matchcode + "' AND NULLIF(X_CASS_CODE, '') IS NOT NULL AND X_CASS_CHECKED + " + num_day_before + " >= getdate()"
        Try
            cmd.CommandText = sql
            cnt = cmd.ExecuteScalar().ToString()
            If cnt = "0" Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return True
        End Try
    End Function
    Private Function IsAddressVerifiedBefore(ByRef cmd As SqlCommand, ADDR As String, LastLine As String, num_day_before As Double) As Boolean
        'SET TABLE
        Dim sql As String, errMessage As String
        Dim cnt As String
        sql = "IF OBJECT_ID('MD_DuplAddrCheckInput') IS NULL Create Table MD_DuplAddrCheckInput (ID int IDENTITY(1,1) NOT NULL, ADDR varchar(500), LastLine varchar(500), InputDate datetime)"
        Try
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            errMessage = ex.Message + vbNewLine
        End Try

        'Check input data
        sql = "SELECT InputDate from siebeldb.dbo.MD_DuplAddrCheckInput" + " WHERE ADDR='" + ADDR + "' AND LastLine = " + LastLine + " InputDate + " + num_day_before.ToString() + " >= getdate()"
        Try
            cmd.CommandText = sql
            cmd.ExecuteReader()
        Catch ex As Exception
            errMessage = ex.Message + vbNewLine
        End Try

        'Log input address to datatable
        If Not Dupl_Addr_Check.g_dt Is Nothing And Dupl_Addr_Check.g_dt.Columns.Count = 0 Then
            Dupl_Addr_Check.g_dt.Columns.Add("ADDR", System.Type.GetType("System.String"))
            Dupl_Addr_Check.g_dt.Columns.Add("LastLine", System.Type.GetType("System.String"))
            Dupl_Addr_Check.g_dt.Columns.Add("InputDate", System.Type.GetType("System.DateTime"))
        End If
        Dim retRow As Data.DataRow()
        retRow = Dupl_Addr_Check.g_dt.Select("ADDR = '" + ADDR + "' and LastLine = '" + LastLine + "' ")
        If retRow.Length > 0 Then
            If retRow(0)("InputDate") - Convert.ToDouble(num_day_before) > DateTime.Now Then
                'Duplicate
                Return True
            Else
                retRow(0).Delete()
                'Dupl_Addr_Check.g_dt.Rows.Remove(retRow(0).)
                Return False
            End If
        Else
            Dim newRow As Data.DataRow = Dupl_Addr_Check.g_dt.NewRow()
            newRow("ADDR") = ADDR
            newRow("LastLine") = LastLine
            newRow("InputDate") = DateTime.Now
            Dupl_Addr_Check.g_dt.Rows.Add(newRow)
            Return False
        End If
    End Function

    Private Function IsAddressVerifiedBefore(ADDR As String, LastLine As String, num_day_before As String) As Boolean
        'Log input address to datatable
        If Not Dupl_Addr_Check.g_dt Is Nothing And Dupl_Addr_Check.g_dt.Columns.Count = 0 Then
            Dupl_Addr_Check.g_dt.Columns.Add("ADDR", System.Type.GetType("System.String"))
            Dupl_Addr_Check.g_dt.Columns.Add("LastLine", System.Type.GetType("System.String"))
            Dupl_Addr_Check.g_dt.Columns.Add("InputDate", System.Type.GetType("System.DateTime"))
        End If
        Dim retRow As Data.DataRow()
        retRow = Dupl_Addr_Check.g_dt.Select("ADDR = '" + ADDR + "' and LastLine = '" + LastLine + "' ")
        If retRow.Length > 0 Then
            If retRow(0)("InputDate") - Convert.ToDouble(num_day_before) > DateTime.Now Then
                'Duplicate
                Return True
            Else
                retRow(0).Delete()
                'Dupl_Addr_Check.g_dt.Rows.Remove(retRow(0).)
                Return False
            End If
        Else
            Dim newRow As Data.DataRow = Dupl_Addr_Check.g_dt.NewRow()
            newRow("ADDR") = ADDR
            newRow("LastLine") = lastline
            newRow("InputDate") = DateTime.Now
            Dupl_Addr_Check.g_dt.Rows.Add(newRow)
            Return False
        End If
    End Function

    Private Function GetMDResultDesc(ByVal ResultsString As String) As String
        Dim OutString As String = ""
        If (InStr(1, ResultsString, "AS")) Then
            ' address was verified
            If (InStr(1, ResultsString, "AS01")) Then
                OutString = OutString + "AS01: Full Address Matched to Postal Database and is deliverable " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS02")) Then
                OutString = OutString + "AS02: Address matched to USPS database but a suite was missing Or invalid" + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS03")) Then
                OutString = OutString + "AS03: Valid Physical Address, not Serviced by the USPS " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS09")) Then
                OutString = OutString + "AS09: Foreign Postal Code Detected " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS10")) Then
                OutString = OutString + "AS10: Address Matched to CMRA" + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS13")) Then
                OutString = OutString + "AS13: Address has been Updated by LACSLink " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS14")) Then
                OutString = OutString + "AS14: Suite Appended by SuiteLink " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS15")) Then
                OutString = OutString + "AS15: Suite Appended by SuiteFinder " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS16")) Then
                OutString = OutString + "AS16: Address is vacant." + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS17")) Then
                OutString = OutString + "AS17: Alternate delivery." + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS18")) Then
                OutString = OutString + "AS18: Artificially created adresses detected,DPV processing terminated at this point" + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS20")) Then
                OutString = OutString + "AS20: Address Deliverable by USPS only " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS21")) Then
                OutString = OutString + "AS21: Alternate Address Suggestion Found" + vbCrLf
            End If
            If (InStr(1, ResultsString, "AS22")) Then
                OutString = OutString + "AS22: No Alternate Address Suggestion Found + vbCrLf"
            End If
            If (InStr(1, ResultsString, "AS23")) Then
                OutString = OutString + "AS23: Extraneous information found " + vbCrLf
            End If
        End If

        If (InStr(1, ResultsString, "AE")) Then
            ' there was an error in verifying the address
            If (InStr(1, ResultsString, "AE01")) Then
                OutString = OutString + "AE01: Zip Code Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE02")) Then
                OutString = OutString + "AE02: Unknown Street Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE03")) Then
                OutString = OutString + "AE03: Component Mismatch Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE04")) Then
                OutString = OutString + "AE04: Non-Deliverable Address Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE05")) Then
                OutString = OutString + "AE05: Multiple Match Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE06")) Then
                OutString = OutString + "AE06: Early Warning System Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE07")) Then
                OutString = OutString + "AE07: Missing Minimum Address Input " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE08")) Then
                OutString = OutString + "AE08: Suite Range Invalid Error" + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE09")) Then
                OutString = OutString + "AE09: Suite Range Missing Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE10")) Then
                OutString = OutString + "AE10: Primary Range Invalid Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE11")) Then
                OutString = OutString + "AE11: Primary Range Missing Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE12")) Then
                OutString = OutString + "AE12: PO, HC, or RR Box Number Invalid " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE13")) Then
                OutString = OutString + "AE13: PO, HC, or RR Box Number Missing " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE14")) Then
                OutString = OutString + "AE14: CMRA Secondary Missing Error" + vbCrLf
            End If

            ' program can not attempt address lookup
            If (InStr(1, ResultsString, "AE15")) Then
                OutString = OutString + "AE15: Demo Mode " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE16")) Then
                OutString = OutString + "AE16: Expired Database" + vbCrLf
            End If

            If (InStr(1, ResultsString, "AE17")) Then
                OutString = OutString + "AE17: Unnecessary Suite Error " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE19")) Then
                OutString = OutString + "AE19: Max time for FindSuggestion exceeded " + vbCrLf
            End If
            If (InStr(1, ResultsString, "AE20")) Then
                OutString = OutString + "AE20: FindSuggestion cannot be used" + vbCrLf
            End If

        End If

        ' a change was made to the address
        If (InStr(1, ResultsString, "AC01")) Then
            OutString = OutString + "AC01: ZIP Code Change " + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC02")) Then
            OutString = OutString + "AC02: State Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC03")) Then
            OutString = OutString + "AC03: City Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC04")) Then
            OutString = OutString + "AC04: Base/Alternate Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC05")) Then
            OutString = OutString + "AC05: Alias Name Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC06")) Then
            OutString = OutString + "AC06: Address1/Address2 Swap" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC07")) Then
            OutString = OutString + "AC07: Address1/Company Swap" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC08")) Then
            OutString = OutString + "AC08: Plus4 Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC09")) Then
            OutString = OutString + "AC09: Urbanization Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC10")) Then
            OutString = OutString + "AC10: Street Name Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC11")) Then
            OutString = OutString + "AC11: Street Suffix Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC12")) Then
            OutString = OutString + "AC12: Street Directional Change" + vbCrLf
        End If
        If (InStr(1, ResultsString, "AC13")) Then
            OutString = OutString + "AC13: Suite Name Change" + vbCrLf
        End If

        Return OutString

    End Function


    ' =================================================
    ' COLLECTIONS 
    ' This class implements a simple dictionary using an array of DictionaryEntry objects (key/value pairs).
    Public Class SimpleDictionary
        Implements IDictionary

        ' The array of items
        Dim items() As DictionaryEntry
        Dim ItemsInUse As Integer = 0

        ' Construct the SimpleDictionary with the desired number of items.
        ' The number of items cannot change for the life time of this SimpleDictionary.
        Public Sub New(ByVal numItems As Integer)
            items = New DictionaryEntry(numItems - 1) {}
        End Sub

        ' IDictionary Members
        Public ReadOnly Property IsReadOnly() As Boolean Implements IDictionary.IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Function Contains(ByVal key As Object) As Boolean Implements IDictionary.Contains
            Dim index As Integer
            Return TryGetIndexOfKey(key, index)
        End Function

        Public ReadOnly Property IsFixedSize() As Boolean Implements IDictionary.IsFixedSize
            Get
                Return False
            End Get
        End Property

        Public Sub Remove(ByVal key As Object) Implements IDictionary.Remove
            If key = Nothing Then
                Throw New ArgumentNullException("key")
            End If
            ' Try to find the key in the DictionaryEntry array
            Dim index As Integer
            If TryGetIndexOfKey(key, index) Then

                ' If the key is found, slide all the items up.
                Array.Copy(items, index + 1, items, index, (ItemsInUse - index) - 1)
                ItemsInUse = ItemsInUse - 1
            Else

                ' If the key is not in the dictionary, just return. 
            End If
        End Sub

        Public Sub Clear() Implements IDictionary.Clear
            ItemsInUse = 0
        End Sub

        Public Sub Add(ByVal key As Object, ByVal value As Object) Implements IDictionary.Add

            ' Add the new key/value pair even if this key already exists in the dictionary.
            If ItemsInUse = items.Length Then
                Throw New InvalidOperationException("The dictionary cannot hold any more items.")
            End If
            items(ItemsInUse) = New DictionaryEntry(key, value)
            ItemsInUse = ItemsInUse + 1
        End Sub

        Public ReadOnly Property Keys() As ICollection Implements IDictionary.Keys
            Get

                ' Return an array where each item is a key.
                ' Note: Declaring keyArray() to have a size of ItemsInUse - 1
                '       ensures that the array is properly sized, in VB.NET
                '       declaring an array of size N creates an array with
                '       0 through N elements, including N, as opposed to N - 1
                '       which is the default behavior in C# and C++.
                Dim keyArray() As Object = New Object(ItemsInUse - 1) {}
                Dim n As Integer
                For n = 0 To ItemsInUse - 1
                    keyArray(n) = items(n).Key
                Next n

                Return keyArray
            End Get
        End Property

        Public ReadOnly Property Values() As ICollection Implements IDictionary.Values
            Get
                ' Return an array where each item is a value.
                Dim valueArray() As Object = New Object(ItemsInUse - 1) {}
                Dim n As Integer
                For n = 0 To ItemsInUse - 1
                    valueArray(n) = items(n).Value
                Next n

                Return valueArray
            End Get
        End Property

        Default Public Property Item(ByVal key As Object) As Object Implements IDictionary.Item
            Get

                ' If this key is in the dictionary, return its value.
                Dim index As Integer
                If TryGetIndexOfKey(key, index) Then

                    ' The key was found return its value.
                    Return items(index).Value
                Else

                    ' The key was not found return null.
                    Return Nothing
                End If
            End Get

            Set(ByVal value As Object)
                ' If this key is in the dictionary, change its value. 
                Dim index As Integer
                If TryGetIndexOfKey(key, index) Then

                    ' The key was found change its value.
                    items(index).Value = value
                Else

                    ' This key is not in the dictionary add this key/value pair.
                    Add(key, value)
                End If
            End Set
        End Property

        Private Function TryGetIndexOfKey(ByVal key As Object, ByRef index As Integer) As Boolean
            For index = 0 To ItemsInUse - 1
                ' If the key is found, return true (the index is also returned).
                If items(index).Key.Equals(key) Then
                    Return True
                End If
            Next index

            ' Key not found, return false (index should be ignored by the caller).
            Return False
        End Function

        Private Class SimpleDictionaryEnumerator
            Implements IDictionaryEnumerator

            ' A copy of the SimpleDictionary object's key/value pairs.
            Dim items() As DictionaryEntry
            Dim index As Integer = -1

            Public Sub New(ByVal sd As SimpleDictionary)
                ' Make a copy of the dictionary entries currently in the SimpleDictionary object.
                items = New DictionaryEntry(sd.Count - 1) {}
                Array.Copy(sd.items, 0, items, 0, sd.Count)
            End Sub

            ' Return the current item.
            Public ReadOnly Property Current() As Object Implements IDictionaryEnumerator.Current
                Get
                    ValidateIndex()
                    Return items(index)
                End Get
            End Property

            ' Return the current dictionary entry.
            Public ReadOnly Property Entry() As DictionaryEntry Implements IDictionaryEnumerator.Entry
                Get
                    Return Current
                End Get
            End Property

            ' Return the key of the current item.
            Public ReadOnly Property Key() As Object Implements IDictionaryEnumerator.Key
                Get
                    ValidateIndex()
                    Return items(index).Key
                End Get
            End Property

            ' Return the value of the current item.
            Public ReadOnly Property Value() As Object Implements IDictionaryEnumerator.Value
                Get
                    ValidateIndex()
                    Return items(index).Value
                End Get
            End Property

            ' Advance to the next item.
            Public Function MoveNext() As Boolean Implements IDictionaryEnumerator.MoveNext
                If index < items.Length - 1 Then
                    index = index + 1
                    Return True
                End If

                Return False
            End Function

            ' Validate the enumeration index and throw an exception if the index is out of range.
            Private Sub ValidateIndex()
                If index < 0 Or index >= items.Length Then
                    Throw New InvalidOperationException("Enumerator is before or after the collection.")
                End If
            End Sub

            ' Reset the index to restart the enumeration.
            Public Sub Reset() Implements IDictionaryEnumerator.Reset
                index = -1
            End Sub

        End Class

        Public Function GetEnumerator() As IDictionaryEnumerator Implements IDictionary.GetEnumerator

            'Construct and return an enumerator.
            Return New SimpleDictionaryEnumerator(Me)
        End Function


        ' ICollection Members
        Public ReadOnly Property IsSynchronized() As Boolean Implements IDictionary.IsSynchronized
            Get
                Return False
            End Get
        End Property

        Public ReadOnly Property SyncRoot() As Object Implements IDictionary.SyncRoot
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public ReadOnly Property Count() As Integer Implements IDictionary.Count
            Get
                Return ItemsInUse
            End Get
        End Property

        Public Sub CopyTo(ByVal array As Array, ByVal index As Integer) Implements IDictionary.CopyTo
            Throw New NotImplementedException()
        End Sub

        ' IEnumerable Members
        Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator

            ' Construct and return an enumerator.
            Return Me.GetEnumerator()
        End Function
    End Class

    ' =================================================
    ' STRING FUNCTIONS
    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        ' Validate email address

        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If

    End Function

    Function SqlString(ByVal Instring As String) As String
        ' Make a string safe for use in a SQL query - filter out all but standard ascii
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            SqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp)
            If Mid(temp, i, 1) = "'" Then
                outstring = outstring & "''"
            Else
                If Asc(Mid(temp, i, 1)) > 0 And Asc(Mid(temp, i, 1)) < 127 Then
                    outstring = outstring & Mid(temp, i, 1)
                End If
            End If
        Next
        SqlString = outstring
    End Function

    Function KeySpace(ByVal Instring As String) As String
        ' Replaces spaces with "+" signs in key fields
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            KeySpace = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp)
            If Mid(temp, i, 1) = " " Then
                outstring = outstring & "+"
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        KeySpace = outstring
    End Function

    Function RemoveSymbols(ByVal Instring As String) As String
        ' Removes all symbols except dashes, converts to uppercase
        Dim temp As String
        Dim outstring, tocheck As String
        Dim i As Integer

        temp = Instring.ToString.ToUpper()
        outstring = ""
        For i = 1 To Len(temp)
            tocheck = Mid(temp, i, 1)
            If Asc(tocheck) = 45 Then
                outstring = outstring & tocheck
            End If
            If Asc(tocheck) > 47 And Asc(tocheck) < 58 Then
                outstring = outstring & tocheck
            End If
            If Asc(tocheck) > 64 And Asc(tocheck) < 91 Then
                outstring = outstring & tocheck
            End If
        Next
        RemoveSymbols = outstring
    End Function

    Function CleanString(ByVal Instring As String) As String
        ' Replaces spaces with "+" signs in key fields
        Dim temp As String
        Dim outstring, tocheck As String
        Dim i As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            CleanString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp)
            tocheck = Mid(temp, i, 1)
            If Asc(tocheck) > 31 And Asc(tocheck) < 127 Then
                Select Case Asc(tocheck)
                    Case 34     ' "
                    Case 38     ' &
                    Case Else
                        outstring = outstring & tocheck
                End Select
            End If
        Next
        CleanString = outstring
    End Function

    Function CheckNull(ByVal Instring As String) As String
        ' Check to see if a string is null
        If Instring Is Nothing Then
            CheckNull = ""
        Else
            CheckNull = Instring
        End If
    End Function

    Public Function CheckDBNull(ByVal obj As Object, _
    Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
        ' Checks an object to determine if its null, and if so sets it to a not-null empty value
        Dim objReturn As Object
        objReturn = obj
        If ObjectType = enumObjectType.StrType And IsDBNull(obj) Then
            objReturn = ""
        ElseIf ObjectType = enumObjectType.IntType And IsDBNull(obj) Then
            objReturn = 0
        ElseIf ObjectType = enumObjectType.DblType And IsDBNull(obj) Then
            objReturn = 0.0
        End If
        Return objReturn
    End Function

    Public Function NumString(ByVal strString As String) As String
        ' Remove everything but numbers from a string
        Dim bln As Boolean
        Dim i As Integer
        Dim iv As String
        NumString = ""

        'Can array element be evaluated as a number?
        For i = 1 To Len(strString)
            iv = Mid(strString, i, 1)
            bln = IsNumeric(iv)
            If bln Then NumString = NumString & iv
        Next

    End Function

    Public Function ToBase64(ByVal data() As Byte) As String
        ' Encode a Base64 string
        If data Is Nothing Then Throw New ArgumentNullException("data")
        Return Convert.ToBase64String(data)
    End Function

    Public Function FromBase64(ByVal base64 As String) As Byte()
        ' Decode a Base64 string
        If base64 Is Nothing Then Throw New ArgumentNullException("base64")
        Return Convert.FromBase64String(base64)
    End Function

    Function DeSqlString(ByVal Instring As String) As String
        ' Convert a string from SQL query encoded to non-encoded
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        CheckDBNull(Instring, enumObjectType.StrType)
        If Len(Instring) = 0 Then
            DeSqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 2) = "''" Then
                outstring = outstring & "'"
                i = i + 1
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        DeSqlString = outstring
    End Function

    Function RemovePluses(ByVal Instring As String) As String
        ' Replace "+" signs in a string with spaces
        Dim temp As String
        Dim outstring As String
        Dim i As Integer

        CheckDBNull(Instring, enumObjectType.StrType)
        If Len(Instring) = 0 Then
            RemovePluses = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 1) = "+" Then
                outstring = outstring & " "
                i = i + 1
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        RemovePluses = outstring

    End Function

    Public Function StringToBytes(ByVal str As String) As Byte()
        ' Convert a random string to a byte array
        ' e.g. "abcdefg" to {a,b,c,d,e,f,g}
        Dim s As Char()
        Dim t As Char
        s = str.ToCharArray
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            If Asc(s(i)) < 128 And Asc(s(i)) > 0 Then
                Try
                    b(i) = Convert.ToByte(s(i))
                Catch ex As Exception
                    b(i) = Convert.ToByte(Chr(32))
                End Try
            Else
                ' Filter out extended ASCII - convert common symbols when possible
                t = Chr(32)
                Try
                    Select Case Asc(s(i))
                        Case 147
                            t = Chr(34)
                        Case 148
                            t = Chr(34)
                        Case 145
                            t = Chr(39)
                        Case 146
                            t = Chr(39)
                        Case 150
                            t = Chr(45)
                        Case 151
                            t = Chr(45)
                        Case Else
                            t = Chr(32)
                    End Select
                Catch ex As Exception
                End Try
                b(i) = Convert.ToByte(t)
            End If
        Next
        Return b
    End Function

    Public Function NumStringToBytes(ByVal str As String) As Byte()
        ' Convert a string containing numbers to a byte array
        ' e.g. "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16" to 
        '  {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16}
        Dim s As String()
        s = str.Split(" ")
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        For i = 0 To s.Length - 1
            b(i) = Convert.ToByte(s(i))
        Next
        Return b
    End Function

    Public Function BytesToString(ByVal b() As Byte) As String
        ' Convert a byte array to a string
        Dim i As Integer
        Dim s As New System.Text.StringBuilder()
        For i = 0 To b.Length - 1
            Console.WriteLine(b(i))
            If i <> b.Length - 1 Then
                s.Append(b(i) & " ")
            Else
                s.Append(b(i))
            End If
        Next
        Return s.ToString
    End Function

    Function StndPhone(ByVal InString) As String
        Dim temp, outstring, tocheck As String
        Dim i As Integer

        temp = Trim(InString)
        If temp = "" Or temp = "--" Or Len(temp) = 0 Then
            StndPhone = ""
            Exit Function
        End If
        outstring = ""

        ' Remove non-numeric characters
        For i = 1 To Len(temp)
            tocheck = Mid(temp, i, 1)
            If IsNumeric(tocheck) Then
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next

        ' Add dashes
        If Mid(outstring, 4, 1) <> "-" Then
            If (Len(outstring) - 3) > 0 Then
                outstring = Left(outstring, 3) & "-" & Mid(outstring, 4, Len(outstring) - 3)
            End If
        End If
        If Mid(outstring, 8, 1) <> "-" Then
            If (Len(outstring) - 7) > 0 Then
                outstring = Left(outstring, 7) & "-" & Mid(outstring, 8, Len(outstring) - 7)
            End If
        End If

        ' Check length
        outstring = Left(outstring, 12)
        StndPhone = outstring
    End Function

    Function StndDate(ByVal InString As String) As String
        Dim tDay, tMonth, tYear As String
        If Not InString Is Nothing And Format(InString, "MM/DD/YYYY") <> "12/30/1899" And Format(InString, "MM/DD/YYYY") <> "01/01/1900" And InString <> "" And InString <> "12:00:00 AM" Then
            StndDate = String.Format(InString, "d")

            Dim tDate As Date = Date.Parse(StndDate)
            tDay = tDate.Day.ToString
            If tDay.Length = 1 Then tDay = "0" & tDay
            tMonth = tDate.Month.ToString
            If tMonth.Length = 1 Then tMonth = "0" & tMonth
            tYear = tDate.Year.ToString
            If tYear.Length = 2 Then tYear = "20" & tYear
            StndDate = tMonth & "/" & tDay & "/" & tYear
        Else
            StndDate = ""
        End If

    End Function

    Function StndSSN(ByVal InString As String) As String
        Dim temp, outstring, tocheck As String
        Dim i As Integer
        temp = InString
        If temp = "" Or InString Is Nothing Then
            StndSSN = ""
            Exit Function
        End If
        outstring = ""

        If temp = "--" Or temp = "" Then
            outstring = ""
            StndSSN = outstring
            Exit Function
        End If

        ' Remove extraneous characters
        For i = 1 To Len(temp)
            tocheck = Mid(temp, i, 1)
            If IsNumeric(tocheck) Then
                outstring = outstring & Mid(temp, i, 1)
            Else
                outstring = outstring & ""
            End If
        Next

        ' Add dashes
        If Len(outstring) > 8 Then
            If Mid(outstring, 4, 1) <> "-" Then
                outstring = Left(outstring, 3) & "-" & Mid(outstring, 4, Len(outstring) - 3)
            End If
            If Mid(outstring, 7, 1) <> "-" Then
                outstring = Left(outstring, 6) & "-" & Mid(outstring, 7, Len(outstring) - 6)
            End If
        End If

        ' Check length
        outstring = Left(outstring, 11)
        StndSSN = outstring
    End Function

    ' =================================================
    ' DATABASE FUNCTIONS
    Public Function OpenDBConnection(ByVal ConnS As String, ByRef con As SqlConnection, ByRef cmd As SqlCommand) As String
        ' Function to open a database connection with extreme error-handling
        ' Returns an error message if unable to open the connection
        Dim SqlS As String
        SqlS = ""
        OpenDBConnection = ""

        Try
            con = New SqlConnection(ConnS)
            con.Open()
            If Not con Is Nothing Then
                Try
                    cmd = New SqlCommand(SqlS, con)
                    cmd.CommandTimeout = 300
                Catch ex2 As Exception
                    OpenDBConnection = "Error opening the command string: " & ex2.ToString
                End Try
            End If
        Catch ex As Exception
            If con.State <> Data.ConnectionState.Closed Then con.Dispose()
            ConnS = ConnS & ";Pooling=false"
            Try
                con = New SqlConnection(ConnS)
                con.Open()
                If Not con Is Nothing Then
                    Try
                        cmd = New SqlCommand(SqlS, con)
                        cmd.CommandTimeout = 300
                    Catch ex2 As Exception
                        OpenDBConnection = "Error opening the command string: " & ex2.ToString
                    End Try
                End If
            Catch ex2 As Exception
                OpenDBConnection = "Unable to open database connection for connection string: " & ConnS & vbCrLf & "Windows error: " & vbCrLf & ex2.ToString & vbCrLf
            End Try
        End Try

    End Function

    ' =================================================
    ' DEBUG FUNCTIONS
    Public Sub writeoutput(ByVal fs As StreamWriter, ByVal instring As String)
        ' This function writes a line to a previously opened streamwriter, and then flushes it
        ' promptly.  This assists in debugging services
        fs.WriteLine(instring)
        fs.Flush()
    End Sub

    Public Sub writeoutputfs(ByVal fs As FileStream, ByVal instring As String)
        ' This function writes a line to a previously opened filestream, and then flushes it
        ' promptly.  This assists in debugging services
        fs.Write(StringToBytes(instring), 0, Len(instring))
        fs.Write(StringToBytes(vbCrLf), 0, 2)
        fs.Flush()
    End Sub

    ' =================================================
    ' HTTP PROXY CLASS
    Class simplehttp
        Public Function geturl(ByVal url As String, ByVal proxyip As String, ByVal port As Integer, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 30000
            req.CookieContainer = New CookieContainer()
            req.Referer = ""
            req.Headers.[Set]("Accept-Language", "en,en-us")
            Dim stream_in As StreamReader

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            'req.Proxy = proxy

            Dim response As String = ""
            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                stream_in = New StreamReader(resp.GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
                'Print(ex.Message)
                Return ex.Message
            End Try
            Return response
        End Function

        Public Function getposturl(ByVal url As String, ByVal postdata As String, ByVal proxyip As String, ByVal port As Short, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Method = "POST"
            req.ContentType = "application/x-www-form-urlencoded"
            req.ContentLength = postdata.Length
            req.Referer = ""

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim stream_out As New StreamWriter(req.GetRequestStream(), System.Text.Encoding.ASCII)
            stream_out.Write(postdata)
            stream_out.Close()
            Dim response As String = ""

            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                Dim resStream As Stream = resp.GetResponseStream()
                Dim stream_in As New StreamReader(req.GetResponse().GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function

    End Class
End Class
