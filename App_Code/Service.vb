Option Explicit On
Imports System.Net
Imports System.Threading
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Collections
Imports System.Configuration
Imports System.Net.Mail
Imports System.Math
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports log4net
'added for SAP RAS coding
Imports CrystalDecisions.ReportAppServer
Imports CrystalDecisions.ReportAppServer.Controllers
Imports CrystalDecisions.ReportAppServer.DataDefModel
Imports CrystalDecisions.ReportAppServer.DataSetConversion
Imports CrystalDecisions.Enterprise
Imports CrystalDecisions.Enterprise.InfoObject
Imports System.Data.OleDb
Imports System.Security

<WebService(Namespace:="http://scormsvc.certegrity.com/CertSvc/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService

    ' PDF Editor Tools declarations
    Private Declare Sub VerySetLicenseKey Lib "verywrite.dll" _
        (ByVal szLicenseKey As String)

    Private Declare Function VeryOpen Lib "verywrite.dll" _
        (ByVal inFileName As String) As Integer

    Private Declare Function VeryClose Lib "verywrite.dll" _
        (ByVal id As Integer) As Integer

    Private Declare Function VeryEncryptPDF Lib "verywrite.dll" _
        (ByVal inFileName As String, ByVal outFileName As String, ByVal EnctyptLen As Integer, _
        ByVal permission As Integer, ByVal OwnerPassword As String, ByVal UserPassword As String) As Integer

    Private Declare Function VeryAddInfo Lib "verywrite.dll" _
        (ByVal id As Integer, ByVal Title As String, ByVal Subject As String, _
        ByVal Author As String, ByVal Keywords As String, ByVal Creator As String) As Integer

    Private Declare Function VerySplitMergePDF Lib "verywrite.dll" _
        (ByVal szCommand As String) As Integer

    Private Declare Function VeryIsPDFEncrypted Lib "verywrite.dll" _
        (ByVal inFileName As String) As Integer

    ' PDF Conversion declarations
    Private Declare Function PDFToImageConverter Lib "pdf2image.dll" _
        (ByVal szPDFFileName As String, ByVal szOutputName As String, _
        ByVal szUserPassword As String, ByVal szOwnPassword As String, _
        ByVal xresolution As Integer, ByVal yresolution As Integer, ByVal bitcount As Integer, _
        ByVal compression As Integer, ByVal quality As Integer, ByVal grayscale As Integer, _
        ByVal multipage As Integer, ByVal firstPage As Integer, ByVal lastPage As Integer) As Integer

    Private Declare Function PDFToImageConverterEx Lib "pdf2image.dll" _
        (ByVal szPDFFileName As String, ByVal szOutputName As String, _
        ByVal szUserPassword As String, ByVal szOwnPassword As String, ByVal pagesizemode As Integer, _
        ByVal xresolution As Integer, ByVal yresolution As Integer, ByVal bitcount As Integer, _
        ByVal compression As Integer, ByVal quality As Integer, ByVal grayscale As Integer, _
        ByVal multipage As Integer, ByVal firstPage As Integer, ByVal lastPage As Integer) As Integer

    Private Declare Function PDFToImageGetPageWidth Lib "pdf2image.dll" _
        (ByVal szPDFFileName As String, ByVal intPage As Integer) As Integer

    Private Declare Sub PDFToImageSetCode Lib "pdf2image.dll" (ByVal szRegcode As String)

    ' Kernel functions
    Public Declare Auto Function LoadLibrary Lib "kernel32.dll" (ByVal lpFileName As String) As IntPtr
    Public Declare Auto Function GetModuleHandle Lib "kernel32.dll" (ByVal lpModuleName As String) As IntPtr
    Public Declare Auto Function FreeLibrary Lib "kernel32.dll" (ByVal hModule As IntPtr) As Boolean

    Private RandNum As New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer))

    Enum CompressionType
        COMPRESSION_NONE = 1              '/* dump mode */
        COMPRESSION_CCITTRLE = 2          '/* CCITT modified Huffman RLE */
        COMPRESSION_CCITTFAX3 = 3         '/* CCITT Group 3 fax encoding */
        COMPRESSION_CCITTFAX4 = 4         '/* CCITT Group 4 fax encoding */
        COMPRESSION_LZW = 5               '/* Lempel-Ziv  & Welch */
        COMPRESSION_JPEG = 7              '/* JPEG DCT compression */
        COMPRESSION_PACKBITS = 32773      '/* Macintosh RLE */
    End Enum

    Enum enumObjectType
        StrType = 0
        IntType = 1
        DblType = 2
        DteType = 3
    End Enum

    Public Enum TagNames As Integer
        ExifIFD = &H8769
        GpsIFD = &H8825
        NewSubfileType = &HFE
        SubfileType = &HFF
        ImageWidth = &H100
        ImageHeight = &H101
        BitsPerSample = &H102
        Compression = &H103
        PhotometricInterp = &H106
        ThreshHolding = &H107
        CellWidth = &H108
        CellHeight = &H109
        FillOrder = &H10A
        DocumentName = &H10D
        ImageDescription = &H10E
        EquipMake = &H10F
        EquipModel = &H110
        StripOffsets = &H111
        Orientation = &H112
        SamplesPerPixel = &H115
        RowsPerStrip = &H116
        StripBytesCount = &H117
        MinSampleValue = &H118
        MaxSampleValue = &H119
        XResolution = &H11A
        YResolution = &H11B
        PlanarConfig = &H11C
        PageName = &H11D
        XPosition = &H11E
        YPosition = &H11F
        FreeOffset = &H120
        FreeByteCounts = &H121
        GrayResponseUnit = &H122
        GrayResponseCurve = &H123
        T4Option = &H124
        T6Option = &H125
        ResolutionUnit = &H128
        PageNumber = &H129
        TransferFuncition = &H12D
        SoftwareUsed = &H131
        DateTime = &H132
        Artist = &H13B
        HostComputer = &H13C
        Predictor = &H13D
        WhitePoint = &H13E
        PrimaryChromaticities = &H13F
        ColorMap = &H140
        HalftoneHints = &H141
        TileWidth = &H142
        TileLength = &H143
        TileOffset = &H144
        TileByteCounts = &H145
        InkSet = &H14C
        InkNames = &H14D
        NumberOfInks = &H14E
        DotRange = &H150
        TargetPrinter = &H151
        ExtraSamples = &H152
        SampleFormat = &H153
        SMinSampleValue = &H154
        SMaxSampleValue = &H155
        TransferRange = &H156
        JPEGProc = &H200
        JPEGInterFormat = &H201
        JPEGInterLength = &H202
        JPEGRestartInterval = &H203
        JPEGLosslessPredictors = &H205
        JPEGPointTransforms = &H206
        JPEGQTables = &H207
        JPEGDCTables = &H208
        JPEGACTables = &H209
        YCbCrCoefficients = &H211
        YCbCrSubsampling = &H212
        YCbCrPositioning = &H213
        REFBlackWhite = &H214
        ICCProfile = &H8773
        Gamma = &H301
        ICCProfileDescriptor = &H302
        SRGBRenderingIntent = &H303
        ImageTitle = &H320
        Copyright = &H8298
        ResolutionXUnit = &H5001
        ResolutionYUnit = &H5002
        ResolutionXLengthUnit = &H5003
        ResolutionYLengthUnit = &H5004
        PrintFlags = &H5005
        PrintFlagsVersion = &H5006
        PrintFlagsCrop = &H5007
        PrintFlagsBleedWidth = &H5008
        PrintFlagsBleedWidthScale = &H5009
        HalftoneLPI = &H500A
        HalftoneLPIUnit = &H500B
        HalftoneDegree = &H500C
        HalftoneShape = &H500D
        HalftoneMisc = &H500E
        HalftoneScreen = &H500F
        JPEGQuality = &H5010
        GridSize = &H5011
        ThumbnailFormat = &H5012
        ThumbnailWidth = &H5013
        ThumbnailHeight = &H5014
        ThumbnailColorDepth = &H5015
        ThumbnailPlanes = &H5016
        ThumbnailRawBytes = &H5017
        ThumbnailSize = &H5018
        ThumbnailCompressedSize = &H5019
        ColorTransferFunction = &H501A
        ThumbnailData = &H501B
        ThumbnailImageWidth = &H5020
        ThumbnailImageHeight = &H502
        ThumbnailBitsPerSample = &H5022
        ThumbnailCompression = &H5023
        ThumbnailPhotometricInterp = &H5024
        ThumbnailImageDescription = &H5025
        ThumbnailEquipMake = &H5026
        ThumbnailEquipModel = &H5027
        ThumbnailStripOffsets = &H5028
        ThumbnailOrientation = &H5029
        ThumbnailSamplesPerPixel = &H502A
        ThumbnailRowsPerStrip = &H502B
        ThumbnailStripBytesCount = &H502C
        ThumbnailResolutionX = &H502D
        ThumbnailResolutionY = &H502E
        ThumbnailPlanarConfig = &H502F
        ThumbnailResolutionUnit = &H5030
        ThumbnailTransferFunction = &H5031
        ThumbnailSoftwareUsed = &H5032
        ThumbnailDateTime = &H5033
        ThumbnailArtist = &H5034
        ThumbnailWhitePoint = &H5035
        ThumbnailPrimaryChromaticities = &H5036
        ThumbnailYCbCrCoefficients = &H5037
        ThumbnailYCbCrSubsampling = &H5038
        ThumbnailYCbCrPositioning = &H5039
        ThumbnailRefBlackWhite = &H503A
        ThumbnailCopyRight = &H503B
        LuminanceTable = &H5090
        ChrominanceTable = &H5091
        FrameDelay = &H5100
        LoopCount = &H5101
        PixelUnit = &H5110
        PixelPerUnitX = &H5111
        PixelPerUnitY = &H5112
        PaletteHistogram = &H5113
        ExifExposureTime = &H829A
        ExifFNumber = &H829D
        ExifExposureProg = &H8822
        ExifSpectralSense = &H8824
        ExifISOSpeed = &H8827
        ExifOECF = &H8828
        ExifVer = &H9000
        ExifDTOrig = &H9003
        ExifDTDigitized = &H9004
        ExifCompConfig = &H9101
        ExifCompBPP = &H9102
        ExifShutterSpeed = &H9201
        ExifAperture = &H9202
        ExifBrightness = &H9203
        ExifExposureBias = &H9204
        ExifMaxAperture = &H9205
        ExifSubjectDist = &H9206
        ExifMeteringMode = &H9207
        ExifLightSource = &H9208
        ExifFlash = &H9209
        ExifFocalLength = &H920A
        ExifMakerNote = &H927C
        ExifUserComment = &H9286
        ExifDTSubsec = &H9290
        ExifDTOrigSS = &H9291
        ExifDTDigSS = &H9292
        ExifFPXVer = &HA000
        ExifColorSpace = &HA001
        ExifPixXDim = &HA002
        ExifPixYDim = &HA003
        ExifRelatedWav = &HA004
        ExifInterop = &HA005
        ExifFlashEnergy = &HA20B
        ExifSpatialFR = &HA20C
        ExifFocalXRes = &HA20E
        ExifFocalYRes = &HA20F
        ExifFocalResUnit = &HA210
        ExifSubjectLoc = &HA214
        ExifExposureIndex = &HA215
        ExifSensingMethod = &HA217
        ExifFileSource = &HA300
        ExifSceneType = &HA301
        ExifCfaPattern = &HA302
        GpsVer = &H0
        GpsLatitudeRef = &H1
        GpsLatitude = &H2
        GpsLongitudeRef = &H3
        GpsLongitude = &H4
        GpsAltitudeRef = &H5
        GpsAltitude = &H6
        GpsGpsTime = &H7
        GpsGpsSatellites = &H8
        GpsGpsStatus = &H9
        GpsGpsMeasureMode = &HA
        GpsGpsDop = &HB
        GpsSpeedRef = &HC
        GpsSpeed = &HD
        GpsTrackRef = &HE
        GpsTrack = &HF
        GpsImgDirRef = &H10
        GpsImgDir = &H11
        GpsMapDatum = &H12
        GpsDestLatRef = &H13
        GpsDestLat = &H14
        GpsDestLongRef = &H15
        GpsDestLong = &H16
        GpsDestBearRef = &H17
        GpsDestBear = &H18
        GpsDestDistRef = &H19
        GpsDestDist = &H1A
    End Enum

    Public Enum ExifDataTypes As Short
        UnsignedByte = 1
        AsciiString = 2
        UnsignedShort = 3
        UnsignedLong = 4
        UnsignedRational = 5
        SignedByte = 6
        Undefined = 7
        SignedShort = 8
        SignedLong = 9
        SignedRational = 10
        SingleFloat = 11
        DoubleFloat = 12
    End Enum

    <WebMethod(Description:="Asynchronously invoke GenCertProd")> _
    Public Function CallGenCertProd(ByVal sXML As String) As Boolean

        ' Calls GenCertProd with the supplied XML document

        ' Misc. Declarations
        Dim success As Boolean
        success = True

        ' Web Service declarations
        Dim request As Net.HttpWebRequest

        ' Create the request
        If InStr(sXML, "http:") > 0 Or InStr(sXML, "https:") > 0 Then
            request = CType(Net.WebRequest.Create(sXML), Net.HttpWebRequest)
        Else
            request = CType(Net.WebRequest.Create("http://hciscormsvc.certegrity.com/CertSvc/service.asmx/GenCertProd?sXML=" & sXML), Net.HttpWebRequest)
        End If
        request.Method = "GET"

        ' Set timeout at 5 minute = 300 seconds
        Dim timeout As Integer
        timeout = 1000 * 300
        request.Timeout = timeout

        ' Create the state object used to access the web request
        Dim state As WebRequestState
        state = New WebRequestState(request)

        ' Begin the async request
        Dim result As IAsyncResult
        result = request.BeginGetResponse(New AsyncCallback(AddressOf RequestComplete), state)

        ' Register a timeout for the async request
        Try
            ThreadPool.RegisterWaitForSingleObject(result.AsyncWaitHandle, New WaitOrTimerCallback(AddressOf TimeoutCallback), state, timeout, True)
        Catch ex As Exception
            success = False
        End Try

        ' ============================================
        ' Return results
CloseOut:
        Return success

    End Function

    ' Method called when a request times out
    Private Sub TimeoutCallback(ByVal state As Object, ByVal timeOut As Boolean)
        If (timeOut) Then
            ' Abort the request
            CType(state, WebRequestState).Request.Abort()
        End If
    End Sub

    ' Method called when the request completes
    Private Sub RequestComplete(ByVal result As IAsyncResult)
        ' Get the request
        Dim request As WebRequest
        request = DirectCast(result.AsyncState, WebRequestState).Request
    End Sub

    ' Stores web request for access during async processing
    Private Class WebRequestState
        ' Holds the request object
        Public Request As WebRequest

        Public Sub New(ByVal newRequest As WebRequest)
            Request = newRequest
        End Sub
    End Class

    <WebMethod(Description:="Schedules an execution of GenCertProd")> _
    Public Function SchedGenCertProd(ByVal sXML As String) As Boolean

        ' The purpose of this web service is to schedule an execution of GenCertProd

        ' This service stores the parameters into a queue record (CX_CERT_PROD_QUEUE) and then
        ' executes CallGenCertProd

        ' The parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <Product>
        '        <Debug>	    - A flag to indicate the service is to run in Debug mode or not
        '                                   "Y"  - Yes for debug mode on.. logging on
        '                                   "N"  - No for debug mode off.. logging off
        '                                   "T"  - Test mode on.. logging off
        '        <QueueId>      - The Id for a report in siebeldb.CX_CERT_PROD_QUEUE, if 
        '                           empty than look for:
        '        <DocId>        - An existing document id to replace - only valid if a single document product is to 
        '                           to be created
        '	     <CrseId>	    - The Course Id (S_CRSE.ROW_ID) of the relevant course
        '        <SkillLevel>   - Skill level code of the certificate (S_CRSE.SKILL_LEVEL_CD)
        '        <SrcId>      	- The Source Id of the individual record to generate (depending on the type of product)
        '        <IdentStart>   - The registration/exam range starting number depending on whether it is a 
        '				            "P"assport or "V"oucher (CX_SESS_REG.MS_IDENT or S_CRSE_TSTRUN.MS_IDENT).
        '        <IdentEnd>     - The registration/exam range ending number depending on whether it is a 
        '				            "P"assport or "V"oucher (CX_SESS_REG.MS_IDENT or S_CRSE_TSTRUN.MS_IDENT)
        '        <TypeProd>     - The type of product to generate ("R","C","W","P" or "V")
        '        <ProdId>  	    - The certification product id (if blank, the most recent primary certificate 
        '				            image) (CX_CERT_PROD.ROW_ID)
        '        <OutputDest>   - Output destination (file only [file], web site [web], mobile platform [mobile], 
        '				            link [link] or image only [image]). 
        '        <SrcQuery>	    - The "where" clause of the data query (if blank, then a query is derived appropriate for the 
        '				            course type or the default query for the certificate product)
        '        <OutFormat>    - The output format of the file - the 3 character extension for the file type (If blank, 
        '				            then default to "pdf").
        '        <JurisId>	    - The jurisdiction id (CX_JURISDICTION_X.ROW_ID).
        '        <OrgId>	    - The organization id (S_ORG_EXT.ROW_ID).
        '        <ConId>	    - The contact id of the individual (S_CONTACT.ROW_ID).
        '        <Domain>	    - The domain (CX_SUB_DOMAIN.DOMAIN). If specified, then create a domain user reference
        '        <NotifyFlg>	- A flag whether to notify the individual specified by the contact id above when the 
        '				            product is available.
        '        <AttachFlg>	- Whether to attach a link prepared to the notification.
        '        <MultiFlg>	    - A flag to indicate whether multiple record results should be stored in a single document or not.
        '        <ReqdFlg>      - A flag to indicate that the document required setting on the result document should be set on the contact association
        '        <SpecialMsg>   - A special message to include in the automated notification
        '        <EmpId>        - The employee id of the individual initiating this task
        '   </Product>

        ' Miscellaneous declarations
        Dim temp As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim debug, ReportId, errmsg, logging, Database, results As String
        Dim mypath As String
        Dim bResults As Boolean

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        Dim logfile, logstring As String
        Dim LogStartTime As String = Now.ToString
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("SGCPDebugLog")

        ' Product data declarations
        Dim QueueId, CrseId, SkillLevel, SrcId, IdentStart, IdentEnd, ProdId As String
        Dim OutputDest, SrcQuery, OutFormat, JurisId, OrgId, ConId As String
        Dim TypeProd, Domain, NotifyFlg, AttachFlg, MultiFlg, ReqdFlg As String
        Dim AccessFlg, PublicKey, EmpId, ExistDocId, SpecialMsg As String
        Dim ProdQueueId As String

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        ReportId = ""
        logging = "Y"
        logstring = ""
        errmsg = ""
        Database = ""

        QueueId = ""
        SkillLevel = ""
        CrseId = ""
        SrcId = ""
        TypeProd = ""
        IdentStart = ""
        IdentEnd = ""
        ProdId = ""
        OutputDest = ""
        SrcQuery = ""
        OutFormat = ""
        JurisId = ""
        OrgId = ""
        ConId = ""
        Domain = ""
        NotifyFlg = ""
        AttachFlg = ""
        MultiFlg = ""
        ReqdFlg = ""
        AccessFlg = ""
        PublicKey = ""
        EmpId = ""
        ExistDocId = ""
        SpecialMsg = ""
        ProdQueueId = ""

        debug = "Y"
        bResults = False
        results = "false"

        ' ============================================
        ' Check parameters
        If sXML = "" Then
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        HttpUtility.UrlDecode(sXML)
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//Product")
        For i = 0 To oNodeList.Count - 1
            Try
                debug = GetNodeValue("Debug", oNodeList.Item(i))
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error reading parameters. " & ex.ToString
                GoTo CloseOut2
            End Try
        Next
        debug = UCase(debug)

        ' ============================================
        ' Get system defaults
        '  Get database connection information
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("SchedGenCertProd_debug")
            If temp = "Y" And debug <> "T" Then debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get hcidb1 defaults from web.config. "
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\SchedGenCertProd.log"
            Try
                log4net.GlobalContext.Properties("SGCPLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' ============================================
        ' Open database connection 
        '  hcidb1
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            GoTo CloseOut
        End If
        If debug = "Y" Then mydebuglog.Debug("  Opened hcidb1 connection")

        ' ============================================
        ' Retrieve record from the XML document
        Try
            For i = 0 To oNodeList.Count - 1
                ' Get the ReportId field and type of output field
                QueueId = GetNodeValue("QueueId", oNodeList.Item(i))
                ExistDocId = Trim(GetNodeValue("DocId", oNodeList.Item(i)))
                CrseId = Trim(GetNodeValue("CrseId", oNodeList.Item(i)))
                If InStr(CrseId, "%") Then CrseId = HttpUtility.UrlEncode(CrseId)
                If InStr(CrseId, " ") Then CrseId = CrseId.Replace(" ", "+")
                SkillLevel = GetNodeValue("SkillLevel", oNodeList.Item(i))
                SrcId = Trim(GetNodeValue("SrcId", oNodeList.Item(i)))
                If InStr(SrcId, "%") Then SrcId = HttpUtility.UrlDecode(SrcId)
                If InStr(SrcId, " ") Then SrcId = SrcId.Replace(" ", "+")
                IdentStart = GetNodeValue("IdentStart", oNodeList.Item(i))
                IdentEnd = GetNodeValue("IdentEnd", oNodeList.Item(i))
                TypeProd = GetNodeValue("TypeProd", oNodeList.Item(i))
                ProdId = GetNodeValue("ProdId", oNodeList.Item(i))
                OutputDest = GetNodeValue("OutputDest", oNodeList.Item(i))
                SrcQuery = GetNodeValue("SrcQuery", oNodeList.Item(i))
                OutFormat = GetNodeValue("OutFormat", oNodeList.Item(i))
                JurisId = Trim(GetNodeValue("JurisId", oNodeList.Item(i)))
                If InStr(JurisId, "%") Then JurisId = HttpUtility.UrlDecode(JurisId)
                If InStr(JurisId, " ") Then JurisId = JurisId.Replace(" ", "+")
                OrgId = Trim(GetNodeValue("OrgId", oNodeList.Item(i)))
                If InStr(OrgId, "%") Then OrgId = HttpUtility.UrlDecode(OrgId)
                If InStr(OrgId, " ") > 0 Then OrgId = OrgId.Replace(" ", "+")
                ConId = Trim(GetNodeValue("ConId", oNodeList.Item(i)))
                If InStr(ConId, "%") Then ConId = HttpUtility.UrlDecode(ConId)
                If InStr(ConId, " ") > 0 Then ConId = ConId.Replace(" ", "+")
                AccessFlg = GetNodeValue("AccessFlg", oNodeList.Item(i))
                Domain = GetNodeValue("Domain", oNodeList.Item(i))
                NotifyFlg = GetNodeValue("NotifyFlg", oNodeList.Item(i))
                AttachFlg = GetNodeValue("AttachFlg", oNodeList.Item(i))
                MultiFlg = GetNodeValue("MultiFlg", oNodeList.Item(i))
                ReqdFlg = GetNodeValue("ReqdFlg", oNodeList.Item(i))
                SpecialMsg = GetNodeValue("SpecialMsg", oNodeList.Item(i))
                If InStr(SpecialMsg, "%") Then SpecialMsg = HttpUtility.UrlDecode(SpecialMsg)
                EmpId = Trim(GetNodeValue("EmpId", oNodeList.Item(i)))
                If InStr(EmpId, "%") Then EmpId = HttpUtility.UrlDecode(EmpId)
                If InStr(EmpId, " ") > 0 Then EmpId = EmpId.Replace(" ", "+")

                ' Reset fields as needed
                If ConId <> "" And AccessFlg = "" Then AccessFlg = "Y" ' If a contact is explicitly identified, give them access
                If NotifyFlg = "Y" And AccessFlg = "N" Then AccessFlg = "Y" ' If you notify them, give them access
                If AccessFlg.Trim = "" Then AccessFlg = "N"
                If ReqdFlg.Trim = "" Then ReqdFlg = "N"
                If Domain = "" Then Domain = "CSI"
                If MultiFlg = "" Then MultiFlg = "N"
                OutFormat = Trim(LCase(OutFormat))
                If OutFormat = "" Then OutFormat = "pdf"
                OutputDest = Trim(LCase(OutputDest))
                If OutputDest = "" Then OutputDest = "file"
                If ExistDocId = "" Then ExistDocId = "NULL"
                If IdentStart = "" Then IdentStart = "NULL"
                If IdentEnd = "" Then IdentEnd = "NULL"

                ' Debug report data found
                If debug = "Y" Then
                    mydebuglog.Debug("  ======================" & vbCrLf & "  Parameters found--")
                    mydebuglog.Debug("   QueueId: " & QueueId)
                    mydebuglog.Debug("   ExistDocId: " & ExistDocId)
                    mydebuglog.Debug("   CrseId: " & CrseId)
                    mydebuglog.Debug("   SkillLevel: " & SkillLevel)
                    mydebuglog.Debug("   SrcId: " & SrcId)
                    mydebuglog.Debug("   IdentStart: " & IdentStart)
                    mydebuglog.Debug("   IdentEnd: " & IdentEnd)
                    mydebuglog.Debug("   TypeProd: " & TypeProd)
                    mydebuglog.Debug("   ProdId: " & ProdId)
                    mydebuglog.Debug("   OutputDest: " & OutputDest)
                    mydebuglog.Debug("   SrcQuery: " & SrcQuery)
                    mydebuglog.Debug("   OutFormat: " & OutFormat)
                    mydebuglog.Debug("   JurisId: " & JurisId)
                    mydebuglog.Debug("   OrgId: " & OrgId)
                    mydebuglog.Debug("   ConId: " & ConId)
                    mydebuglog.Debug("   AccessFlg: " & AccessFlg)
                    mydebuglog.Debug("   Domain: " & Domain)
                    mydebuglog.Debug("   NotifyFlg: " & NotifyFlg)
                    mydebuglog.Debug("   AttachFlg: " & AttachFlg)
                    mydebuglog.Debug("   MultiFlg: " & MultiFlg)
                    mydebuglog.Debug("   SpecialMsg: " & SpecialMsg)
                    mydebuglog.Debug("   ReqdFlg: " & ReqdFlg)
                    mydebuglog.Debug("   EmpId: " & EmpId)
                    mydebuglog.Debug("  ======================" & vbCrLf)
                End If
            Next
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Error Opening Log. "
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Verify that we have a record
        If SrcId = "" Then
            errmsg = errmsg & vbCrLf & "Source Id missing. "
            GoTo CloseOut
        End If
        If TypeProd = "" Then
            errmsg = errmsg & vbCrLf & "Product Type missing. "
            GoTo CloseOut
        End If
        If CrseId = "" Then
            errmsg = errmsg & vbCrLf & "Course Id missing. "
            GoTo CloseOut
        End If

        ' ============================================
        ' Check to see if a duplicate of an unfinished entry
        SqlS = "SELECT ROW_ID " & _
        "FROM siebeldb.dbo.CX_CERT_PROD_QUEUE " & _
        "WHERE EXECUTED IS NULL AND CRSE_ID='" & CrseId & "' AND SRC_ID='" & SrcId & "' " & _
        "AND PROD_TYPE='" & TypeProd & "' AND CON_ID='" & ConId & "' AND PROD_ID='" & ProdId & "'"
        If debug = "Y" Then mydebuglog.Debug("  Check to see if submitted twice and just waiting: " & vbCrLf & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                        If temp <> "" Then
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "    >> Found duplicate: " & temp & vbCrLf)
                            bResults = True
                            GoTo CloseOut
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading CX_CERT_PROD_QUEUE record: " & ex.ToString
                    End Try
                End While
            End If
            dr.Close()
            dr = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Error reading CX_CERT_PROD_QUEUE record: " & ex.ToString
        End Try

        ' ============================================
        ' Generate unique record id 
        Try
            ProdQueueId = LoggingService.GenerateRecordId("CX_CERT_PROD_QUEUE", "N", debug)
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to generate product queue id. "
            GoTo CloseOut
        End Try
        If ProdQueueId = "" Then
            errmsg = errmsg & vbCrLf & "Unable to generate product queue id. "
            GoTo CloseOut
        End If
        If debug = "Y" Then mydebuglog.Debug("  New ProdQueueId: " & ProdQueueId)

        ' ============================================
        ' Create CX_CERT_PROD_QUEUE Queue Entry
        SqlS = "INSERT INTO siebeldb.dbo.CX_CERT_PROD_QUEUE " & _
        "(CONFLICT_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,PROD_TYPE,TRAIN_TYPE" & _
        ",CRSE_ID,SRC_ID,IDENT_START,IDENT_END,NOTIFY_FLG,ATTACH_FLG,MULTI_OUT_FLG,FORMAT,DOMAIN" & _
        ",JURIS_ID,CON_ID,OU_ID,DEST_CODE,EXECUTED,ROW_ID,NO_RESULTS_FLG,ACCESS_FLG,REQD_FLG" & _
        ",PROD_ID,SPECIAL_MSG,DOC_ID) " & _
        "VALUES(0,GETDATE(),'0-1',GETDATE(),'0-1',0,'" & TypeProd & "','" & SkillLevel & "'" & _
        ",'" & CrseId & "','" & SrcId & "'," & IdentStart & "," & IdentEnd & ",'" & NotifyFlg & "','" & AttachFlg & "','" & MultiFlg & "','" & OutFormat & "','" & Domain & "'" & _
        ",'" & JurisId & "','" & ConId & "','" & OrgId & "','" & OutputDest & "',NULL,'" & ProdQueueId & "',NULL,'" & AccessFlg & "','" & ReqdFlg & "'" & _
        ",'" & ProdId & "','" & SpecialMsg & "'," & ExistDocId & ")"
        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Inserting into CX_CERT_PROD_QUEUE: " & vbCrLf & SqlS & vbCrLf)
        Try
            cmd.CommandText = SqlS
            returnv = cmd.ExecuteNonQuery()
            If returnv = 0 Then
                errmsg = errmsg & "Unable to insert Product Queue Entry"
            End If
        Catch ex As Exception
            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Unable to insert Product Queue Entry: " & ex.ToString & vbCrLf & SqlS)
        End Try

        ' ============================================
        ' Execute CallGenCertProd service if queue entry created
        SqlS = "SELECT ROW_ID FROM siebeldb.dbo.CX_CERT_PROD_QUEUE WHERE ROW_ID='" & ProdQueueId & "'"
        If debug = "Y" Then mydebuglog.Debug("  Verify product queue entry: " & vbCrLf & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                        If temp <> "" Then
                            sXML = "<Product><Debug>" & debug & "</Debug><QueueId>" & temp & "</QueueId><CrseId></CrseId><SkillLevel></SkillLevel><SrcId></SrcId><IdentStart></IdentStart><IdentEnd></IdentEnd><TypeProd></TypeProd><ProdId></ProdId><OutputDest>link</OutputDest><SrcQuery></SrcQuery><OutFormat></OutFormat><JurisId></JurisId><OrgId></OrgId><ConId></ConId><AccessFlg></AccessFlg><Domain></Domain><NotifyFlg></NotifyFlg><AttachFlg></AttachFlg><MultiFlg></MultiFlg><ReqdFlg></ReqdFlg></Product>"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Service call: " & vbCrLf & sXML)
                            Try
                                If sXML <> "" Then
                                    Dim http As New simplehttp()
                                    'results = http.geturl("http://hciscormsvc.certegrity.com/CertSvc/service.asmx/GenCertProd?sXML=" & sXML, "192.168.7.61", 443, "", "")
                                    results = GenCertProd(sXML).InnerText()
                                End If

                            Catch ex As Exception
                                errmsg = errmsg & "Error generating certification products. " & ex.ToString & vbCrLf & " in query: " & vbCrLf & sXML & vbCrLf
                            End Try
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & " GenCertProd results: " & results)
                            If InStr(LCase(results), "success") > 0 Or InStr(LCase(results), "true") > 0 Then bResults = True
                            'If InStr(LCase(results), "error") = 0 Then bResults = True

                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CallGenCertProd Results: " & results & vbCrLf)
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading CX_CERT_PROD_QUEUE record: " & ex.ToString
                    End Try
                End While
            End If
            dr.Close()
            dr = Nothing
        Catch ex As Exception
            errmsg = errmsg & "Error reading CX_CERT_PROD_QUEUE record: " & ex.ToString
        End Try

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            ' hcidb1
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to close the database connection. "
        End Try

CloseOut2:
        ' ============================================
        ' CLOSE THE LOG FILE
        If Trim(errmsg) <> "" Then myeventlog.Error("SchedGenCertProd : Error: " & Trim(errmsg))
        myeventlog.Info("SchedGenCertProd : Results: " & bResults.ToString() & " for QueueId (passed in ): " & QueueId & " QueueId (queued): " & temp)
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "Error: " & Trim(errmsg))
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Log Performance Data
        If debug <> "T" Then
            Try
                'LoggingService.LogPerformanceDataAsync(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, debug)
                Dim VersionNum As String = "100"
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return bResults
    End Function

    <WebMethod(Description:="Generates current certification cards when applicable")> _
    Public Function GenCurrentCards(ByVal ConId As String, ByVal Debug As String) As Boolean

        ' This service schedules the generation of any missing certification card products 
        ' when provided a Contact Id

        ' The parameter is as follows:
        '	Debug	    - A flag to indicate the service is to run in Debug mode or not
        '				"Y"  - Yes for debug mode on.. logging on
        '				"N"  - No for debug mode off.. logging off
        '				"T"  - Test mode on.. logging off
        '	ConId	    - The contact id of the individual (S_CONTACT.ROW_ID).

        ' Miscellaneous declarations
        Dim temp As String
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i As Integer
        Dim ReportId, errmsg, logging, Database, results As String
        Dim mypath As String
        Dim bResults As Boolean

        ' Database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim SqlS As String
        Dim ConnS As String

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("GCCDebugLog")
        Dim logfile, logstring As String
        Dim LogStartTime As String = Now.ToString

        ' Product data declarations
        Dim sXML As String
        Dim QueueId, CrseId, SkillLevel, SrcId, IdentStart, IdentEnd, ProdId As String
        Dim OutputDest, SrcQuery, OutFormat, JurisId, OrgId As String
        Dim TypeProd, Domain, NotifyFlg, AttachFlg, MultiFlg, ReqdFlg, OuId As String
        Dim AccessFlg, PublicKey, EmpId, ExistDocId, SpecialMsg, SessPartId, CardProdId As String
        Dim pXML(100) As String
        Dim SpCnt, CpCnt As Integer
        Dim ProdQueueId As String

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service
        Dim http As New simplehttp()

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        ReportId = ""
        logging = "Y"
        logstring = ""
        errmsg = ""
        Database = ""
        SpCnt = 0
        CpCnt = 0
        sXML = 0
        JurisId = ""
        CrseId = ""
        OuId = ""
        CardProdId = ""
        SessPartId = ""
        QueueId = ""
        SkillLevel = ""
        SrcId = ""
        TypeProd = ""
        IdentStart = ""
        IdentEnd = ""
        ProdId = ""
        OutputDest = ""
        SrcQuery = ""
        OutFormat = ""
        OrgId = ""
        Domain = ""
        NotifyFlg = ""
        AttachFlg = ""
        MultiFlg = ""
        ReqdFlg = ""
        AccessFlg = ""
        PublicKey = ""
        EmpId = ""
        ExistDocId = ""
        SpecialMsg = ""
        ProdQueueId = ""

        debug = "Y"
        bResults = False
        results = "False"

        ' ============================================
        ' Check and fix parameters
        If ConId = "" Then
            errmsg = errmsg & vbCrLf & "No Contact Id supplied"
            GoTo CloseOut2
        End If
        ConId = UCase(Trim(HttpUtility.UrlEncode(ConId)))
        If InStr(ConId, "%") > 0 Then ConId = UCase(Trim(HttpUtility.UrlDecode(ConId)))
        If InStr(ConId, " ") > 0 Then Replace(ConId, " ", "+")
        debug = UCase(Trim(debug))

        ' ============================================
        ' Get system defaults
        '  Get database connection information
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("siebeldb").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GenCurrentCards_debug")
            If temp = "Y" And debug <> "T" Then debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get defaults from web.config. "
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\GenCurrentCards.log"
            Try
                log4net.GlobalContext.Properties("GCCLogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
            Catch ex As Exception
                errmsg = errmsg & vbCrLf & "Error Opening Log. "
                GoTo CloseOut2
            End Try

            If debug = "Y" Then
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("Trace Log Started " & Now.ToString & vbCrLf)
                mydebuglog.Debug("Parameters-")
                mydebuglog.Debug("  ConId:" & ConId)
            End If
        End If

        ' ============================================
        ' Open database connection 
        '  hcidb1
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            GoTo CloseOut
        End If

        ' ============================================
        ' Locate any ungenerated participation cards to be generated and store to table
        '  This query compensates for duplicate entries in CX_PART_CURRCLM

        SqlS = "SELECT CP.CURRENT_SPART_ID, CPD.ROW_ID, SP.JURIS_ID, CRS.ROW_ID AS CRSE_ID, SP.OU_ID, CRS.X_SUMMARY_CD " & _
         "FROM siebeldb.dbo.CX_PART_CURRCLM CP " & _
         "INNER JOIN siebeldb.dbo.CX_PARTICIPANT_X P ON P.ROW_ID=CP.PART_ID " & _
         "INNER JOIN siebeldb.dbo.S_CONTACT CN ON CN.X_PART_ID=P.ROW_ID " & _
         "INNER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=CP.CURRENT_SPART_ID " & _
         "INNER JOIN siebeldb.dbo.CX_CERT_PROD_CRSE CPC ON CPC.CRSE_ID=SP.CRSE_TST_ID " & _
         "INNER JOIN siebeldb.dbo.CX_CERT_PROD CPD ON CPD.ROW_ID=CPC.CERT_ID " & _
         "INNER JOIN siebeldb.dbo.S_CRSE CRS ON CRS.ROW_ID=SP.CRSE_TST_ID " & _
         "WHERE CN.ROW_ID='" & ConId & "' AND CPC.ROW_ID IS NOT NULL AND CPD.PROD_TYPE='C' AND " & _
         "CP.CURRENT_CARD_ID IS NULL AND CP.CURRENT_EXP_DT>=GETDATE() AND CPD.STATUS_CD='Active' AND CPC.CERT_STATUS IN ('Y','') " & _
         "GROUP BY CP.CURRENT_SPART_ID, CPD.ROW_ID, SP.JURIS_ID, CRS.ROW_ID, SP.OU_ID, CRS.X_SUMMARY_CD"
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locate participant cards to generate: " & vbCrLf & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        SessPartId = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                        CardProdId = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                        JurisId = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                        CrseId = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                        OuId = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                        Domain = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                        If CrseId <> "" Then
                            SpCnt = SpCnt + 1
                            sXML = "<Product><Debug>N</Debug><QueueId></QueueId><CrseId>" & CrseId & _
                                "</CrseId><SkillLevel>Participant</SkillLevel><SrcId>" & SessPartId & _
                                "</SrcId><IdentStart></IdentStart><IdentEnd></IdentEnd><TypeProd>C</TypeProd><ProdId>" & CardProdId & _
                                "</ProdId><OutputDest>file</OutputDest><SrcQuery></SrcQuery><OutFormat>jpg</OutFormat><JurisId>" & JurisId & _
                                "</JurisId><OrgId>" & OuId & "</OrgId><ConId>" & ConId & "</ConId><AccessFlg>Y</AccessFlg><Domain>" & Domain & _
                                "</Domain><NotifyFlg>N</NotifyFlg><AttachFlg>N</AttachFlg><MultiFlg>N</MultiFlg><ReqdFlg>N</ReqdFlg></Product>"
                            ' Schedule the generation of participant cards
                            Try
                                If sXML <> "" Then
                                    If Debug = "Y" Then mydebuglog.Debug("    > sXML: " & sXML)
                                    results = CallGenCertProd(sXML)
                                    'results = http.geturl("http://hciscormsvc.certegrity.com/CertSvc/service.asmx/CallGenCertProd?sXML=" & sXML, "192.168.7.61", 443, "", "")
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & "Error generating participation certification products. " & ex.ToString & vbCrLf & " in query: " & vbCrLf & sXML & vbCrLf
                            End Try
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading CX_PART_CURRCLM records: " & ex.ToString
                    End Try
                End While
            End If
            Try
                dr.Close()
                dr = Nothing
            Catch ex As Exception
            End Try
        Catch ex As Exception
            errmsg = errmsg & "Error reading CX_PART_CURRCLM records: " & ex.ToString
        End Try
        If Debug = "Y" Then mydebuglog.Debug("    > SpCnt found: " & SpCnt.ToString)

        ' ============================================
        ' Locate any trainer cards to be generated 
        SqlS = "SELECT '<Product><Debug>N</Debug><QueueId></QueueId><CrseId>'+C.ROW_ID+'</CrseId><SkillLevel>Trainer</SkillLevel><SrcId>'+CP.ROW_ID+'</SrcId><IdentStart> " & _
        "</IdentStart><IdentEnd></IdentEnd><TypeProd>C</TypeProd><ProdId>'+CPD.ROW_ID+'</ProdId><OutputDest>file</OutputDest><SrcQuery></SrcQuery><OutFormat>jpg</OutFormat> " & _
        "<JurisId></JurisId><OrgId>'+CP.X_ACCOUNT_ID+'</OrgId><ConId>Z3XJ8YEGS7W3</ConId><AccessFlg>Y</AccessFlg><Domain>'+CM.X_SUMMARY_CD+'</Domain><NotifyFlg>N</NotifyFlg> " & _
        "<AttachFlg>N</AttachFlg><MultiFlg>N</MultiFlg><ReqdFlg>N</ReqdFlg></Product>' AS CARD_LINK " & _
        "FROM siebeldb.dbo.S_CURRCLM_PER CP " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_CURRCLM CM ON CM.ROW_ID=CP.CURRCLM_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.CX_CONTACT_CURRCLM CC ON CP.ROW_ID=CC.CURRENT_CERT_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TSTRUN T ON T.ROW_ID=CP.X_CRSE_TSTRUN_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST CT ON CT.ROW_ID=T.CRSE_TST_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE C ON C.ROW_ID=CT.CRSE_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.CX_CERT_PROD_CRSE CPC ON CPC.CRSE_ID=CT.CRSE_ID " & _
        "LEFT OUTER JOIN siebeldb.dbo.CX_CERT_PROD CPD ON CPD.ROW_ID=CPC.CERT_ID " & _
        "WHERE CP.PERSON_ID='" & ConId & "' AND CP.GRANT_DT IS NOT NULL AND CC.CURRENT_EXP_DT>=GETDATE() AND CPC.ROW_ID IS NOT NULL AND " & _
        "CPD.PROD_TYPE='C' AND (CC.CURRENT_CARD_ID='' OR CC.CURRENT_CARD_ID='0' OR CC.CURRENT_CARD_ID IS NULL) AND CPC.CERT_STATUS IN ('Y','') " & _
        "GROUP BY CM.NAME, C.NAME, CP.GRANT_DT, CP.GRANTED_FLG, CP.EXPIRATION_DT, CP.ROW_ID, CP.X_ACCOUNT_ID, CM.X_SUMMARY_CD, C.ROW_ID, CPD.ROW_ID "
        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locate trainer cards to generate: " & vbCrLf & SqlS)
        Try
            cmd.CommandText = SqlS
            dr = cmd.ExecuteReader()
            If Not dr Is Nothing Then
                While dr.Read()
                    Try
                        sXML = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                        If sXML <> "" Then
                            ' Schedule the generation of trainer cards
                            CpCnt = CpCnt + 1
                            Try
                                If sXML <> "" Then
                                    If Debug = "Y" Then mydebuglog.Debug("    > sXML: " & sXML)
                                    results = CallGenCertProd(sXML)
                                    'results = http.geturl("http://hciscormsvc.certegrity.com/CertSvc/service.asmx/CallGenCertProd?sXML=" & sXML, "192.168.7.61", 443, "", "")
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & "Error generating trainer certification products. " & ex.ToString & vbCrLf & " in query: " & vbCrLf & sXML & vbCrLf
                            End Try
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading CX_CONTACT_CURRCLM records: " & ex.ToString
                    End Try
                End While
            End If
            Try
                dr.Close()
                dr = Nothing
            Catch ex As Exception
            End Try
        Catch ex As Exception
            errmsg = errmsg & "Error reading CX_CONTACT_CURRCLM records: " & ex.ToString
        End Try
        If Debug = "Y" Then mydebuglog.Debug("    > CpCnt found: " & CpCnt.ToString & vbCrLf)

        Try
            dr.Close()
            dr = Nothing
        Catch ex As Exception
        End Try
        If CpCnt = 0 And SpCnt = 0 Then bResults = False Else bResults = True

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            ' hcidb1
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
            http = Nothing
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to close the database connection. "
        End Try

CloseOut2:
        ' ============================================
        ' CLOSE THE LOG FILE
        If Trim(errmsg) <> "" Then myeventlog.Error("GenCurrentCards : Error: " & Trim(errmsg))
        myeventlog.Info("GenCurrentCards : Results: " & bResults.ToString() & " for ConId: " & ConId)
        If Debug = "Y" Or (logging = "Y" And Debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "Error: " & Trim(errmsg))
                If Debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & "Results " & bResults)
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Log Performance Data
        If Debug <> "T" Then
            Try
                'LoggingService.LogPerformanceDataAsync(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, Debug)
                Dim VersionNum As String = "100"
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, Debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return bResults
    End Function

    Private Function openRASReport(RASFolderID As String, REPORT_NAME As String, REP_FILENAME As String, ByRef Report As CrystalDecisions.ReportAppServer.ClientDoc.ReportClientDocument, ByRef mydebuglog As log4net.ILog, Debug As String) As Boolean
        Dim boEnterpriseSession As CrystalDecisions.Enterprise.EnterpriseSession
        Dim boInfoObject As CrystalDecisions.Enterprise.InfoObject
        'Dim Report As CrystalDecisions.ReportAppServer.ClientDoc.ReportClientDocument
        'Dim myExportOptions As CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptionsClass
        'Dim tempByteArray As CrystalDecisions.ReportAppServer.CommonObjectModel.ByteArray

        Dim boSessionMgr As CrystalDecisions.Enterprise.SessionMgr
        Dim boInfoStore As CrystalDecisions.Enterprise.InfoStore
        Dim boEnterpriseService As CrystalDecisions.Enterprise.EnterpriseService
        Dim boInfoObjects As CrystalDecisions.Enterprise.InfoObjects
        Dim boReportName As String
        Dim boQuery As String
        Dim boReportAppFactory As CrystalDecisions.ReportAppServer.ClientDoc.ReportAppFactory

        Dim isSuccess As Boolean = True

        ' Open the Report 
        Try
            mydebuglog.Debug(vbCrLf & "  Opening from  SAP RAS " & REPORT_NAME)
            If Debug = "Y" Then mydebuglog.Debug("  Report: reports\" & REP_FILENAME)

            '**** Open CR report from RAS in HCIDW ****
            'Log on to the Enterprise CMS
            boSessionMgr = New SessionMgr()
            'boEnterpriseSession = boSessionMgr.Logon(ConfigurationManager.AppSettings.Get("cruser"), ConfigurationManager.AppSettings.Get("crpwd"), ConfigurationManager.AppSettings.Get("cms") & ":6400", "secEnterprise")
            boEnterpriseSession = boSessionMgr.Logon(ConfigurationManager.AppSettings.Get("cruser"), ConfigurationManager.AppSettings.Get("crpwd"), ConfigurationManager.AppSettings.Get("cms") & "", "secEnterprise")
            mydebuglog.Debug(vbCrLf & "  Login to SAP RAS succeed!")

            boEnterpriseService = boEnterpriseSession.GetService("", "InfoStore")
            boInfoStore = New CrystalDecisions.Enterprise.InfoStore(boEnterpriseService)
            boReportName = REP_FILENAME.Replace(".rpt", "") '"1b0c8d173331321e78.rpt"
            'Retrieve the report object from the InfoStore, only need the SI_ID for RAS
            boQuery = "Select SI_ID From CI_INFOOBJECTS Where SI_KIND = 'CRYSTALREPORT' AND SI_NAME = '" _
                        & boReportName & "' AND SI_Instance=0 " & "AND SI_PARENT_FOLDER=" & RASFolderID
            If Debug = "Y" Then mydebuglog.Debug("  SAP RAS Query: " & boQuery)
            boInfoObjects = boInfoStore.Query(boQuery)
            If boInfoObjects.Count > 0 Then
                mydebuglog.Debug(vbCrLf & "  Query Report succeeds!")
            Else
                mydebuglog.Debug(vbCrLf & "  Query Report failed!")
            End If
            boInfoObject = boInfoObjects(1)
            boEnterpriseService = Nothing
            'Retrieve the RASReportFactory
            boEnterpriseService = boEnterpriseSession.GetService("RASReportFactory")
            boReportAppFactory = CType(boEnterpriseService.Interface, CrystalDecisions.ReportAppServer.ClientDoc.ReportAppFactory)
            'Open the report from Enterprise
            Report = boReportAppFactory.OpenDocument(boInfoObject.ID, 0)
            mydebuglog.Debug(vbCrLf & "  Report opened!")
        Catch ex As Exception
            mydebuglog.Debug(vbCrLf & "  Error: " & ex.ToString)
            isSuccess = False
            GoTo closeout
        End Try

closeout:
        'Close service
        If boEnterpriseSession.IsServerLogonSession Then boEnterpriseSession.Logoff()
        If Not boEnterpriseService Is Nothing Then boEnterpriseService.Dispose()
        If Not boSessionMgr Is Nothing Then boSessionMgr.Dispose()

        Return isSuccess
    End Function

    <WebMethod(Description:="Generate a personalized certification product on request")> _
    Public Function GenCertProd(ByVal sXML As String) As XmlDocument

        ' The purpose of this web service is to generate the certification product specified

        ' If supplied as a product generation queue record id then it searches the queue, and uses the 
        ' information found to generate a product.

        ' If supplied the parameters in an XML document, which contains the same data that is in the queue, then
        ' use it to generate the product and return the results.

        ' The parameter is as follows:
        '   sXML        -   An XML document in the following form:
        '     <Product>
        '        <Debug>	    - A flag to indicate the service is to run in Debug mode or not
        '                                   "Y"  - Yes for debug mode on.. logging on
        '                                   "N"  - No for debug mode off.. logging off
        '                                   "T"  - Test mode on.. logging off
        '        <QueueId>      - The Id for a report in siebeldb.CX_CERT_PROD_QUEUE, if 
        '                           empty than look for:
        '        <DocId>        - An existing document id to replace - only valid if a single document product is to 
        '                           to be created
        '	     <CrseId>	    - The Course Id (S_CRSE.ROW_ID) of the relevant course
        '        <SkillLevel>   - Skill level code of the certificate (S_CRSE.SKILL_LEVEL_CD)
        '        <SrcId>      	- The Source Id of the individual record to generate (depending on the type of product)
        '        <IdentStart>   - The registration/exam range starting number depending on whether it is a 
        '				            "P"assport or "V"oucher (CX_SESS_REG.MS_IDENT or S_CRSE_TSTRUN.MS_IDENT).
        '        <IdentEnd>     - The registration/exam range ending number depending on whether it is a 
        '				            "P"assport or "V"oucher (CX_SESS_REG.MS_IDENT or S_CRSE_TSTRUN.MS_IDENT)
        '        <TypeProd>     - The type of product to generate ("R","C","W","P" or "V")
        '        <ProdId>  	    - The certification product id (if blank, the most recent primary certificate 
        '				            image) (CX_CERT_PROD.ROW_ID)
        '        <OutputDest>   - Output destination (file only [file], web site [web], mobile platform [mobile], 
        '				            link [link] or image only [image]). 
        '        <SrcQuery>	    - The "where" clause of the data query (if blank, then a query is derived appropriate for the 
        '				            course type or the default query for the certificate product)
        '        <OutFormat>    - The output format of the file - the 3 character extension for the file type (If blank, 
        '				            then default to "pdf").
        '        <JurisId>	    - The jurisdiction id (CX_JURISDICTION_X.ROW_ID).
        '        <OrgId>	    - The organization id (S_ORG_EXT.ROW_ID).
        '        <ConId>	    - The contact id of the individual (S_CONTACT.ROW_ID).
        '        <Domain>	    - The domain (CX_SUB_DOMAIN.DOMAIN). If specified, then create a domain user reference
        '        <NotifyFlg>	- A flag whether to notify the individual specified by the contact id above when the 
        '				            product is available.
        '        <AttachFlg>	- Whether to attach a link prepared to the notification.
        '        <MultiFlg>	    - A flag to indicate whether multiple record results should be stored in a single document or not.
        '        <ReqdFlg>      - A flag to indicate that the document required setting on the result document should be set on the contact association
        '        <SpecialMsg>   - A special message to include in the automated notification
        '        <EmpId>        - The employee id of the individual initiating this task
        '   </Product>

        ' web.config Parameters used:
        '   email           - connection string to scanner database
        '   dbuser          - database username
        '   dbpass          - database user password
        '   attachments     - path to email attachments directory
        '   basepath        - path to where local files can be found

        ' Miscellaneous declarations
        Dim results, temp As String
        Dim iresults As Integer
        Dim lresults As Integer
        Dim iDoc As XmlDocument = New XmlDocument()
        Dim i, j, k, l As Integer
        Dim debug, ReportId, errmsg, logging, Database As String
        Dim mypath, basepath As String
        Dim OutputPath As String
        Dim crEDTDiskFile As Integer
        Dim bResults As Boolean

        ' PDF declarations
        Dim PdfPassword, PdfKeywords, PdfSubject As String
        Dim pdfid As Integer

        ' HCIDB database declarations
        Dim con As SqlConnection
        Dim cmd As SqlCommand
        Dim dr As SqlDataReader
        Dim dt As DataTable
        Dim SqlS As String
        Dim ConnS As String
        Dim returnv As Integer
        Dim dbuser, dbpass As String
        Dim tempstream As FileStream
        Dim myparam As SqlParameter
        Dim MyCB As SqlCommandBuilder
        'Dim da As SqlDataAdapter
        Dim ds As Data.DataSet
        Dim SaveTries As Integer

        ' DMS database declarations
        Dim dcon As SqlConnection
        Dim dcmd As SqlCommand
        Dim addDoc As SqlCommand
        Dim ddr As SqlDataReader
        Dim dConnS, NEW_DOCS, REG_ID, supervisor As String
        Dim dmsdocs(100, 21) As String     ' Array to store document ids and filenames
        Dim dmsbin(1, 1) As Byte
        Dim dmsBytes As New System.Collections.Generic.List(Of Byte()) 'Ren Hou; 1-3-2017; modified to fix error

        Dim DmsKeyField, DmsConId, DmsJurisId, DmsOrgId, DmsCertId, DataTypeId, DmsTrainerId As Integer
        Dim DmsWshopId, DmsSessId, DmsOffrId, DmsPartId, DmsWregId, DmsSregId, DmsExamId As Integer
        Dim DmsEmpId As String

        ' Crystal Reports declarations
        'Dim Report As New ReportDocument
        Dim Report As CrystalDecisions.ReportAppServer.ClientDoc.ReportClientDocument
        'Dim myExportOptions As New ExportOptions
        Dim myExportOptions As CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptionsClass
        Dim tempByteArray As CrystalDecisions.ReportAppServer.CommonObjectModel.ByteArray
        Dim myDiskFileDestinationOptions As New DiskFileDestinationOptions
        ' Dim ADOrs As ADODB.Recordset
        Dim adors As ADODB.Recordset

        ' CDO Definition declarations
        Dim CDOFields(100) As String        ' Store the field names
        Dim NumRows, NumCols As Integer
        Dim CDODefFn, strLine As String
        Dim CDOStream As System.IO.StreamReader

        ' Datatable declarations
        Dim dtRow As DataRow
        Dim dtColumn As DataColumn
        Dim adoField As ADODB.Field

        ' Logging declarations
        Dim myeventlog As log4net.ILog
        Dim mydebuglog As log4net.ILog
        Dim fs As FileStream
        myeventlog = log4net.LogManager.GetLogger("EventLog")
        mydebuglog = log4net.LogManager.GetLogger("GCPDebugLog")
        Dim logfile, logstring As String
        Dim LogStartTime As String = Now.ToString

        ' Product data declarations
        Dim QueueId, NewQueueId, CrseId, SkillLevel, SrcId, IdentStart, IdentEnd, ProdId As String
        Dim OutputDest, SrcQuery, OutFormat, JurisId, OrgId, ConId, ConIdUserKey As String
        Dim TypeProd, Domain, NotifyFlg, AttachFlg, MultiFlg, ReqdFlg As String
        Dim AccessFlg, UGAId, PublicKey, SPECIAL_NOTICE, EmpId, ExistDocId, ActivityId As String
        Dim JURIS_CERT_ID_FLG, JURIS_CERT_EMAIL, TEMP_PROD_ID, JURISDICTION, CERT_NUM, ID_POOL_ID, NEW_POOL_ID As String
        Dim PResX, PResY, PWidth, PHeight As Integer
        Dim OrignalImage As Image
        Dim rQUEUE_ID, rCERT_POOL_ID, rCONTACT_ID, rCRSE_ID, rREG_ID, rGENERATED, rDESTINATION, rDOC_ID, rPROD_ID As String

        Dim KeyVal As String
        Dim NumFiles, PageWidth As Integer

        ' Image declarations
        Dim lFileLength As Integer

        ' Output report declarations
        Dim ReportQuery As String
        Dim REP_FILENAME, CERT_TYPE, CERT_QUERY, FORMAT_NAME, FORMAT_CODE, FORMAT_EXTENSION, PRICE_LIST, START_DATE As String
        Dim REPORT_NAME, OUT_FILENAME, END_DATE, PARAMETER, DESTINATION As String
        Dim CONTACT_ID, ACCOUNT_ID, REP_DESC, SUB_ID, SQL_REP, SQL_MOD As String
        Dim ADDL_DESC, DESCRIPTION, ENT_CON_ID, SUPPRESS_BLANK_FLG As String
        Dim UserKey, Extension As String
        Dim RecordsRead, RecordsPrinted As Integer
        Dim b(100) As Byte

        ' Email declarations
        Dim ReplyTo, SendTo, Subject, Body, Body2, Letter, EOL, SpecialMsg As String
        Dim FST_NAME, LAST_NAME, EMAIL_ADDR, ACCESS_URL, FROM_ID, FROM_NAME As String
        Dim eFST_NAME, eLAST_NAME, eEMAIL_ADDR, eCON_ID, eJOB_TITLE, eWORK_PH_NUM, SIGNATURE, eLOGIN As String
        Dim MsgXml, LANG_CD As String
        Dim http As New simplehttp()

        ' Web service declarations
        Dim LoggingService As New com.certegrity.cloudsvc.Service
        Dim DmsService As New local.hq.datafluxapp.dms.Service

        ' ============================================
        ' Variable setup
        mypath = HttpRuntime.AppDomainAppPath
        ReportId = ""
        logging = "Y"
        logstring = ""
        errmsg = ""
        SaveTries = 0
        crEDTDiskFile = 1
        Database = ""
        RecordsRead = 0
        Letter = ""
        NumFiles = 0
        NumRows = 0
        NumCols = 0
        iresults = 0
        lresults = 0
        lFileLength = 0
        DmsKeyField = 0
        DmsConId = 0
        DmsJurisId = 0
        DmsOrgId = 0
        DmsWshopId = 0
        DmsSessId = 0
        DmsOffrId = 0
        DmsPartId = 0
        DmsWregId = 0
        DmsSregId = 0
        DmsExamId = 0
        DmsCertId = 0
        DataTypeId = 0
        DmsTrainerId = 0
        DmsEmpId = "0"
        UGAId = ""

        QueueId = ""
        NewQueueId = ""
        SkillLevel = ""
        CrseId = ""
        SrcId = ""
        TypeProd = ""
        IdentStart = ""
        IdentEnd = ""
        ProdId = ""
        OutputDest = ""
        SrcQuery = ""
        OutFormat = ""
        JurisId = ""
        OrgId = ""
        ConId = ""
        ConIdUserKey = ""
        Domain = ""
        NotifyFlg = ""
        AttachFlg = ""
        MultiFlg = ""
        ReqdFlg = ""
        AccessFlg = ""
        ReportQuery = ""
        PublicKey = ""
        UserKey = ""
        Extension = ""
        SpecialMsg = ""
        EmpId = ""
        KeyVal = ""
        ExistDocId = ""
        ActivityId = ""

        SendTo = ""
        Subject = ""
        Body = ""
        Body2 = ""
        ReplyTo = ""
        FROM_ID = ""
        FROM_NAME = ""
        FST_NAME = ""
        LAST_NAME = ""
        EMAIL_ADDR = ""
        EOL = Chr(10) & Chr(13)
        eFST_NAME = ""
        eLAST_NAME = ""
        eEMAIL_ADDR = ""
        eCON_ID = ""
        eJOB_TITLE = ""
        eWORK_PH_NUM = ""
        eLOGIN = ""
        SIGNATURE = ""
        LANG_CD = "ENU"

        REPORT_NAME = ""
        REP_FILENAME = ""
        CERT_TYPE = ""
        CERT_QUERY = ""
        OUT_FILENAME = ""
        FORMAT_NAME = ""
        FORMAT_CODE = ""
        FORMAT_EXTENSION = ""
        PRICE_LIST = ""
        START_DATE = ""
        END_DATE = ""
        PARAMETER = ""
        CONTACT_ID = ""
        ENT_CON_ID = ""
        ACCOUNT_ID = ""
        REP_DESC = ""
        ADDL_DESC = ""
        DESCRIPTION = ""
        SUB_ID = ""
        SQL_REP = ""
        SQL_MOD = ""
        DESTINATION = ""
        ACCESS_URL = ""
        SUPPRESS_BLANK_FLG = "N"
        SPECIAL_NOTICE = ""
        results = "Success"
        NEW_DOCS = ""
        supervisor = ""
        REG_ID = ""
        Body2 = ""
        RecordsPrinted = 0
        PageWidth = 0
        debug = "Y"
        PResX = 0
        PResY = 0
        PWidth = 0
        PHeight = 0
        bResults = False
        JURIS_CERT_EMAIL = ""
        TEMP_PROD_ID = ""
        JURISDICTION = ""
        NEW_POOL_ID = ""
        ID_POOL_ID = ""
        CERT_NUM = ""

        rQUEUE_ID = ""
        rCERT_POOL_ID = ""
        rCONTACT_ID = ""
        rCRSE_ID = ""
        rREG_ID = ""
        rGENERATED = ""
        rDESTINATION = ""
        rDOC_ID = ""
        rPROD_ID = ""

        PdfPassword = ""
        PdfKeywords = ""
        PdfSubject = ""

        ' ============================================
        ' Check parameters
        If sXML = "" Then
            results = "Failure"
            errmsg = errmsg & vbCrLf & "No parameters. "
            GoTo CloseOut2
        End If
        HttpUtility.UrlDecode(sXML)
        iDoc.LoadXml(sXML)
        Dim oNodeList As XmlNodeList = iDoc.SelectNodes("//Product")
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
        'debug = "Y"

        ' Write XML query to file if debug is set
        If logging = "Y" Then
            logfile = "C:\Logs\GenCertProdXML.log"
            Try
                If System.IO.File.Exists(logfile) Then
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
        ' Get system defaults
        '  Get database connection information
        Try
            ConnS = System.Configuration.ConfigurationManager.ConnectionStrings("email").ConnectionString
            If ConnS = "" Then ConnS = "server=HCIDBSQL\HCIDB;uid=sa;pwd=k3v5c2!k3v5c2;database=siebeldb"
            basepath = System.Configuration.ConfigurationManager.AppSettings.Get("basepath")
            If basepath = "" Then basepath = "C:\Inetpub\CertSvc\"
            dbuser = System.Configuration.ConfigurationManager.AppSettings.Get("dbuser")
            If dbuser = "" Then dbuser = "SIEBEL"
            dbpass = System.Configuration.ConfigurationManager.AppSettings.Get("dbpass")
            If dbpass = "" Then dbpass = "SIEBEL"
            temp = System.Configuration.ConfigurationManager.AppSettings.Get("GenCertProd_debug")
            If temp = "Y" And debug <> "T" Then debug = "Y"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get hcidb1 defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try

        ' ============================================
        ' Open log file if applicable
        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            logfile = "C:\Logs\GenCertProd.log"
            Try
                log4net.GlobalContext.Properties("GCPLogFileName") = logfile
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
                mydebuglog.Debug("  ReportId: " & ReportId)
                mydebuglog.Debug("  input xml:" & HttpUtility.UrlDecode(sXML))
            End If
        End If

        ' ============================================
        '  Get dms connection information
        Try
            dConnS = System.Configuration.ConfigurationManager.ConnectionStrings("dms").ConnectionString
            If dConnS = "" Then dConnS = "server=HCIDBSQL\HCIDB;uid=DMS;pwd=5241200;database=DMS"
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to get dms defaults from web.config. "
            results = "Failure"
            GoTo CloseOut2
        End Try
        If debug = "Y" Then
            Try
                mydebuglog.Debug("  Basepath: " & basepath)
                mydebuglog.Debug("  dbuser: " & dbuser & vbCrLf & "  dbpass: " & dbpass)
                mydebuglog.Debug(vbCrLf & "Report-")
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Open database connection 
        '  hcidb
        errmsg = OpenDBConnection(ConnS, con, cmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        If debug = "Y" Then mydebuglog.Debug("  Opened hcidb1 connection")
        '  dms connection #1
        errmsg = OpenDBConnection(dConnS, dcon, dcmd)
        If errmsg <> "" Then
            results = "Failure"
            GoTo CloseOut
        End If
        If debug = "Y" Then mydebuglog.Debug("  Opened dms connection")

        ' ============================================
        ' Retrieve record from the XML document
        Try
            For i = 0 To oNodeList.Count - 1
                ' Get the ReportId field and type of output field
                JURIS_CERT_ID_FLG = "N"
                QueueId = GetNodeValue("QueueId", oNodeList.Item(i))
                If QueueId <> "" And InStr(QueueId, " ") > 0 Then QueueId = QueueId.Replace(" ", "+")
                If debug = "T" Then QueueId = "[test id here]"
                errmsg = ""

                ' If this doesn't exist, it's a new job get from XML, 
                ' otherwise pull existing from CX_CERT_PROD_QUEUE
                If QueueId = "" Then
                    ExistDocId = Trim(GetNodeValue("DocId", oNodeList.Item(i)))
                    CrseId = Trim(GetNodeValue("CrseId", oNodeList.Item(i)))
                    If InStr(CrseId, "%") > 0 Then CrseId = HttpUtility.UrlEncode(CrseId)
                    If InStr(CrseId, " ") > 0 Then CrseId = CrseId.Replace(" ", "+")
                    SkillLevel = GetNodeValue("SkillLevel", oNodeList.Item(i))
                    SrcId = Trim(GetNodeValue("SrcId", oNodeList.Item(i)))
                    If InStr(SrcId, "%") > 0 Then SrcId = HttpUtility.UrlDecode(SrcId)
                    If InStr(SrcId, " ") > 0 Then SrcId = SrcId.Replace(" ", "+")
                    IdentStart = GetNodeValue("IdentStart", oNodeList.Item(i))
                    IdentEnd = GetNodeValue("IdentEnd", oNodeList.Item(i))
                    TypeProd = GetNodeValue("TypeProd", oNodeList.Item(i))
                    ProdId = GetNodeValue("ProdId", oNodeList.Item(i))
                    OutputDest = GetNodeValue("OutputDest", oNodeList.Item(i))
                    SrcQuery = GetNodeValue("SrcQuery", oNodeList.Item(i))
                    OutFormat = GetNodeValue("OutFormat", oNodeList.Item(i))
                    JurisId = Trim(GetNodeValue("JurisId", oNodeList.Item(i)))
                    If InStr(JurisId, "%") > 0 Then JurisId = HttpUtility.UrlDecode(JurisId)
                    If InStr(JurisId, " ") > 0 Then JurisId = JurisId.Replace(" ", "+")
                    OrgId = Trim(GetNodeValue("OrgId", oNodeList.Item(i)))
                    If InStr(OrgId, "%") > 0 Then OrgId = HttpUtility.UrlDecode(OrgId)
                    If InStr(OrgId, " ") > 0 Then OrgId = OrgId.Replace(" ", "+")
                    ConId = Trim(GetNodeValue("ConId", oNodeList.Item(i)))
                    If InStr(ConId, "%") > 0 Then ConId = HttpUtility.UrlDecode(ConId)
                    If InStr(ConId, " ") > 0 Then ConId = ConId.Replace(" ", "+")
                    If ConId <> "" Then ConIdUserKey = GenerateUserKey(ConId)
                    AccessFlg = GetNodeValue("AccessFlg", oNodeList.Item(i))
                    Domain = GetNodeValue("Domain", oNodeList.Item(i))
                    NotifyFlg = GetNodeValue("NotifyFlg", oNodeList.Item(i))
                    AttachFlg = GetNodeValue("AttachFlg", oNodeList.Item(i))
                    MultiFlg = GetNodeValue("MultiFlg", oNodeList.Item(i))
                    ReqdFlg = GetNodeValue("ReqdFlg", oNodeList.Item(i))
                    SpecialMsg = GetNodeValue("SpecialMsg", oNodeList.Item(i))
                    If InStr(SpecialMsg, "%") > 0 Then SpecialMsg = HttpUtility.UrlDecode(SpecialMsg)
                    EmpId = Trim(GetNodeValue("EmpId", oNodeList.Item(i)))
                    If InStr(EmpId, "%") > 0 Then EmpId = HttpUtility.UrlDecode(EmpId)
                    If InStr(EmpId, " ") > 0 Then EmpId = EmpId.Replace(" ", "+")
                Else
                    If debug = "T" Then
                        ' Get record from test table
                        SqlS = "SELECT TOP 1 * FROM reports.dbo.TEST_CERT_PROD"
                    Else
                        ' Get record from queue
                        SqlS = "SELECT Q.PROD_TYPE,C.SKILL_LEVEL_CD,Q.CRSE_ID,Q.SRC_ID,Q.IDENT_START,Q.IDENT_END,Q.NOTIFY_FLG,Q.ATTACH_FLG," & _
                        "Q.MULTI_OUT_FLG,Q.FORMAT,Q.DOMAIN,Q.JURIS_ID,Q.CON_ID,Q.OU_ID,Q.DEST_CODE,Q.PROD_ID,Q.ACCESS_FLG,Q.REQD_FLG,Q.SPECIAL_MSG,Q.EMP_ID,Q.DOC_ID " & _
                        "FROM siebeldb.dbo.CX_CERT_PROD_QUEUE Q " & _
                        "LEFT OUTER JOIN siebeldb.dbo.S_CRSE C ON C.ROW_ID=Q.CRSE_ID " & _
                        "WHERE Q.ROW_ID='" & QueueId & "' AND Q.EXECUTED IS NULL"
                    End If
                    If debug = "Y" Then mydebuglog.Debug("  Get product queue entry query: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Found product queue entry")
                                    TypeProd = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                    SkillLevel = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                    CrseId = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                    SrcId = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                    IdentStart = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                                    IdentEnd = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                                    NotifyFlg = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                                    AttachFlg = Trim(CheckDBNull(dr(7), enumObjectType.StrType)).ToString
                                    MultiFlg = Trim(CheckDBNull(dr(8), enumObjectType.StrType)).ToString
                                    OutFormat = Trim(CheckDBNull(dr(9), enumObjectType.StrType)).ToString
                                    Domain = Trim(CheckDBNull(dr(10), enumObjectType.StrType)).ToString
                                    JurisId = Trim(CheckDBNull(dr(11), enumObjectType.StrType)).ToString
                                    ConId = Trim(CheckDBNull(dr(12), enumObjectType.StrType)).ToString
                                    If ConId <> "" Then ConIdUserKey = GenerateUserKey(ConId)
                                    OrgId = Trim(CheckDBNull(dr(13), enumObjectType.StrType)).ToString
                                    OutputDest = Trim(CheckDBNull(dr(14), enumObjectType.StrType)).ToString
                                    ProdId = Trim(CheckDBNull(dr(15), enumObjectType.StrType)).ToString
                                    AccessFlg = Trim(CheckDBNull(dr(16), enumObjectType.StrType)).ToString
                                    ReqdFlg = Trim(CheckDBNull(dr(17), enumObjectType.StrType)).ToString
                                    SpecialMsg = Trim(CheckDBNull(dr(18), enumObjectType.StrType)).ToString
                                    EmpId = Trim(CheckDBNull(dr(19), enumObjectType.StrType)).ToString
                                    ExistDocId = Trim(CheckDBNull(dr(20), enumObjectType.StrType)).ToString
                                    SrcQuery = ""
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading product queue entry. " & ex.ToString
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "Error reading product queue entry."
                            results = "Failure"
                            GoTo CloseOut
                        End If
                        Try
                            dr.Close()
                            dr = Nothing
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading product queue entry."
                        results = "Failure"
                        GoTo CloseOut
                    End Try

                    ' Verify that we have a job
                    If SrcId = "" And (IdentStart = "" And IdentEnd = "") Then
                        errmsg = errmsg & "The record was already generated."
                        results = "Failure"
                        GoTo CloseOut
                    End If
                End If

                ' Log job
                If IdentStart <> "" Then
                    logstring = "for records '" & IdentStart & "' to '" & IdentEnd & "'"
                Else
                    logstring = "for record id '" & SrcId & "'"
                End If
                If debug = "Y" Or (logging = "Y" And debug <> "T") Then mydebuglog.Debug(Now.ToString & ": Generating product type '" & TypeProd & "' to '" & OutputDest & "' " & logstring & " and contact id '" & ConId & "'")

                ' Reset fields as needed
                If ConId <> "" And AccessFlg = "" Then AccessFlg = "Y" ' If a contact is explicitly identified, give them access
                If NotifyFlg = "Y" And AccessFlg = "N" Then AccessFlg = "Y" ' If you notify them, give them access
                If AccessFlg.Trim = "" Then AccessFlg = "N"
                If ReqdFlg.Trim = "" Then ReqdFlg = "N"
                If Domain = "" Then Domain = "CSI"
                If MultiFlg = "" Then MultiFlg = "N"
                If OutFormat = "" Then OutFormat = "pdf"

                ' Generate QueueId for the purposes of logging the job and generating a filename
                If QueueId = "" Then
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Creating a QueueId")
CheckFN:
                    ' Generate and check for uniqueness a random filename
                    NewQueueId = LoggingService.GenerateRecordId("CX_CERT_PROD_QUEUE", "N", debug)
                    If NewQueueId = "" Then
                        NewQueueId = Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65)
                    End If
                    OutputPath = basepath & "temp\" & NewQueueId & ".pdf"
                    If debug = "Y" Then mydebuglog.Debug("   ... checking to see if exists: " & OutputPath)
                    If (My.Computer.FileSystem.FileExists(OutputPath)) Then GoTo CheckFN

                    ' Create queue entry
                    SqlS = "INSERT INTO siebeldb.dbo.CX_CERT_PROD_QUEUE " & _
                    "(CONFLICT_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,PROD_TYPE,TRAIN_TYPE" & _
                    ",CRSE_ID,SRC_ID,IDENT_START,IDENT_END,NOTIFY_FLG,ATTACH_FLG,MULTI_OUT_FLG,FORMAT,DOMAIN" & _
                    ",JURIS_ID,CON_ID,OU_ID,DEST_CODE,EXECUTED,ROW_ID,NO_RESULTS_FLG,ACCESS_FLG,REQD_FLG" & _
                    ",PROD_ID,SPECIAL_MSG,DOC_ID) " & _
                    "VALUES(0,GETDATE(),'0-1',GETDATE(),'0-1',0,'" & TypeProd & "','" & SkillLevel & "'" & _
                    ",'" & CrseId & "','" & SrcId & "',"
                    If IdentStart = "" Then temp = "NULL" Else temp = IdentStart
                    SqlS = SqlS & temp & ","
                    If IdentEnd = "" Then temp = "NULL" Else temp = IdentEnd
                    SqlS = SqlS & temp & ",'" & NotifyFlg & "','" & AttachFlg & "','" & MultiFlg & "','" & OutFormat & "','" & Domain & "'" & _
                    ",'" & JurisId & "','" & ConId & "','" & OrgId & "','" & OutputDest & "',NULL,'" & NewQueueId & "',NULL,'" & AccessFlg & "','" & ReqdFlg & "'" & _
                    ",'" & ProdId & "','" & SpecialMsg & "',"
                    If ExistDocId = "" Then temp = "NULL" Else temp = ExistDocId
                    SqlS = SqlS & temp & ")"
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Creating a CX_CERT_PROD_QUEUE entry: " & vbCrLf & SqlS & vbCrLf)
                    Try
                        cmd.CommandText = SqlS
                        returnv = cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        errmsg = errmsg & "Unable to insert CX_CERT_PROD_QUEUE record: " & ex.ToString & vbCrLf
                    End Try

                    ' Verify queue created
                    SqlS = "SELECT ROW_ID FROM siebeldb.dbo.CX_CERT_PROD_QUEUE WHERE ROW_ID='" & NewQueueId & "'"
                    If debug = "Y" Then mydebuglog.Debug("  Verify product queue entry: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                    If temp <> "" Then
                                        QueueId = temp
                                        If debug = "Y" Then mydebuglog.Debug("   ... New QueueId: " & QueueId)
                                    End If
                                Catch ex As Exception
                                    errmsg = errmsg & "Error reading new CX_CERT_PROD_QUEUE record: " & ex.ToString
                                End Try
                            End While
                        End If
                        Try
                            dr.Close()
                            dr = Nothing
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading new CX_CERT_PROD_QUEUE record: " & ex.ToString
                    End Try
                End If

                ' Debug report data found
                If debug = "Y" Then
                    mydebuglog.Debug("   ======================" & vbCrLf & "   Parameters found--")
                    mydebuglog.Debug("   QueueId: " & QueueId)
                    mydebuglog.Debug("   ExistDocId: " & ExistDocId)
                    mydebuglog.Debug("   CrseId: " & CrseId)
                    mydebuglog.Debug("   SkillLevel: " & SkillLevel)
                    mydebuglog.Debug("   SrcId: " & SrcId)
                    mydebuglog.Debug("   IdentStart: " & IdentStart)
                    mydebuglog.Debug("   IdentEnd: " & IdentEnd)
                    mydebuglog.Debug("   TypeProd: " & TypeProd)
                    mydebuglog.Debug("   ProdId: " & ProdId)
                    mydebuglog.Debug("   OutputDest: " & OutputDest)
                    mydebuglog.Debug("   SrcQuery: " & SrcQuery)
                    mydebuglog.Debug("   OutFormat: " & OutFormat)
                    mydebuglog.Debug("   JurisId: " & JurisId)
                    mydebuglog.Debug("   OrgId: " & OrgId)
                    mydebuglog.Debug("   ConId: " & ConId)
                    mydebuglog.Debug("   ConIdUserKey: " & ConIdUserKey)
                    mydebuglog.Debug("   AccessFlg: " & AccessFlg)
                    mydebuglog.Debug("   Domain: " & Domain)
                    mydebuglog.Debug("   NotifyFlg: " & NotifyFlg)
                    mydebuglog.Debug("   AttachFlg: " & AttachFlg)
                    mydebuglog.Debug("   MultiFlg: " & MultiFlg)
                    mydebuglog.Debug("   SpecialMsg: " & SpecialMsg)
                    mydebuglog.Debug("   EmpId: " & EmpId)
                    mydebuglog.Debug("   ======================" & vbCrLf)
                End If

                ' -----
                ' LOCATE CERTIFICATION PRODUCT 
                ' Based on supplied parameters lookup the certification product
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Locating Product")
                If ProdId <> "" Then
                    SqlS = "SELECT P.ROW_ID, P.RFILENAME, P.CERT_TYPE, P.CERT_QUERY, P.NAME, P.DEF_FORMAT, " & _
                    "P.SPECIAL_NOTICE, P.RES_X, P.RES_Y, P.WIDTH, P.HEIGHT, PC.JURIS_CERT_ID_FLG, PC.JURIS_CERT_EMAIL, PC.TEMP_PROD_ID, J.NAME " & _
                    "FROM siebeldb.dbo.CX_CERT_PROD P " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_CERT_PROD_CRSE PC ON PC.CERT_ID=P.ROW_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X J ON J.ROW_ID=PC.JURIS_ID " & _
                    "WHERE P.ROW_ID='" & ProdId & "' AND PC.CRSE_ID='" & CrseId & "'"
                Else
                    SqlS = "SELECT P.ROW_ID, P.RFILENAME, (SELECT CASE WHEN P.CERT_TYPE IS NULL THEN R.SKILL_LEVEL_CD ELSE P.CERT_TYPE END) AS CERT_TYPE, " & _
                    "P.CERT_QUERY, P.NAME, P.DEF_FORMAT, P.SPECIAL_NOTICE, P.RES_X, P.RES_Y, P.WIDTH, P.HEIGHT, C.JURIS_CERT_ID_FLG, C.JURIS_CERT_EMAIL, " & _
                    "C.TEMP_PROD_ID, J.NAME " & _
                    "FROM siebeldb.dbo.CX_CERT_PROD_CRSE C " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_CERT_PROD P ON P.ROW_ID=C.CERT_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE R ON R.ROW_ID=C.CRSE_ID " & _
                    "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X J ON J.ROW_ID=C.JURIS_ID " & _
                    "WHERE C.CRSE_ID='" & CrseId & "' AND C.PRIMARY_FLG='Y' AND P.PROD_TYPE='" & TypeProd & "'"
                    If Domain <> "" Then SqlS = SqlS & " AND C.DOMAIN='" & Domain & "'"
                    If SkillLevel <> "" Then SqlS = SqlS & " AND P.CERT_TYPE='" & SkillLevel & "'"
                End If
                If debug = "Y" Then mydebuglog.Debug("  Get certification product: " & vbCrLf & SqlS)
                Try
                    cmd.CommandText = SqlS
                    dr = cmd.ExecuteReader()
                    If Not dr Is Nothing Then
                        While dr.Read()
                            Try
                                If ProdId = "" Then ProdId = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                REP_FILENAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                CERT_TYPE = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                CERT_QUERY = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                REPORT_NAME = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                                If OutFormat = "" Then OutFormat = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                                SPECIAL_NOTICE = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                                PResX = CheckDBNull(dr(7), enumObjectType.IntType)
                                PResY = CheckDBNull(dr(8), enumObjectType.IntType)
                                PWidth = CheckDBNull(dr(9), enumObjectType.IntType)
                                PHeight = CheckDBNull(dr(10), enumObjectType.IntType)
                                JURIS_CERT_ID_FLG = Trim(CheckDBNull(dr(11), enumObjectType.StrType))
                                JURIS_CERT_EMAIL = Trim(CheckDBNull(dr(12), enumObjectType.StrType))
                                TEMP_PROD_ID = Trim(CheckDBNull(dr(13), enumObjectType.StrType))
                                JURISDICTION = Trim(CheckDBNull(dr(14), enumObjectType.StrType))
                            Catch ex As Exception
                                errmsg = errmsg & "Error locating product for this course. " & ex.ToString
                            End Try
                        End While
                    Else
                        errmsg = errmsg & "Error locating product for this course. "
                        results = "Failure"
                        GoTo ErrorQueue
                    End If
                    Try
                        dr.Close()
                        dr = Nothing
                    Catch ex As Exception
                    End Try
                Catch ex As Exception
                    errmsg = errmsg & "Error locating product for this course. " & ex.ToString
                    results = "Failure"
                    GoTo ErrorQueue
                End Try

                ' If report not found, then obviously there is an issue
                If REP_FILENAME = "" Then
                    errmsg = errmsg & "The product was not found for this course"
                    GoTo ErrorQueue
                End If
                PWidth = PWidth * 3
                PHeight = PHeight * 3
                Dim CropRect As New Rectangle(0, 0, PWidth, PHeight)

                ' Debug product information found
                If debug = "Y" Then
                    mydebuglog.Debug("  Product found--")
                    mydebuglog.Debug("   ProdId: " & ProdId)
                    mydebuglog.Debug("   REP_FILENAME: " & REP_FILENAME)
                    mydebuglog.Debug("   CERT_TYPE: " & CERT_TYPE)
                    mydebuglog.Debug("   CERT_QUERY: " & vbCrLf & CERT_QUERY)
                    mydebuglog.Debug("   Resolution: " & PResX.ToString & " x " & PResY.ToString)
                    mydebuglog.Debug("   Dimensions: " & PWidth.ToString & " x " & PHeight.ToString)
                    mydebuglog.Debug("   JURIS_CERT_ID_FLG: " & JURIS_CERT_ID_FLG)
                    mydebuglog.Debug("   JURIS_CERT_EMAIL: " & JURIS_CERT_EMAIL)
                    mydebuglog.Debug("   TEMP_PROD_ID: " & TEMP_PROD_ID)
                    mydebuglog.Debug("   JURISDICTION: " & JURISDICTION)
                End If

                ' -----
                ' PROCESS A JURISDICTION SPECIFIC CERTIFICATION NUMBER PRODUCT
                If JURIS_CERT_ID_FLG = "Y" Then
                    If debug = "Y" Then mydebuglog.Debug("  JURIS_CERT_ID_FLG  " & vbCrLf)
                    ' Checking any supplied existing document id
                    If ExistDocId <> "" Then
                        ' Retrieve existing information based on supplied document id
                        SqlS = "SELECT Q.ROW_ID AS QUEUE_ID, R.CERT_POOL_ID, C.ROW_ID AS CONTACT_ID, Q.CRSE_ID, R.REG_ID, R.GENERATED, R.DESTINATION, R.DOC_ID, R.PROD_ID " & _
                            "FROM siebeldb.dbo.CX_CERT_PROD_RESULTS R " & _
                            "LEFT OUTER JOIN siebeldb.dbo.CX_CERT_PROD_QUEUE Q ON R.PROD_QUEUE_ID=Q.ROW_ID " & _
                            "LEFT OUTER JOIN siebeldb.dbo.CX_CERT_PROD_ID_POOL P ON P.ROW_ID=R.CERT_POOL_ID " & _
                            "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CON_ID " & _
                            "WHERE R.ROW_ID='" & ExistDocId & "'"
                        If debug = "Y" Then mydebuglog.Debug("  Checking existing document id: " & vbCrLf & SqlS)
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        rQUEUE_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                        rCERT_POOL_ID = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                        rCONTACT_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                        rCRSE_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType))
                                        rREG_ID = Trim(CheckDBNull(dr(4), enumObjectType.StrType))
                                        rGENERATED = Trim(CheckDBNull(dr(5), enumObjectType.StrType))
                                        rDESTINATION = Trim(CheckDBNull(dr(6), enumObjectType.StrType))
                                        rDOC_ID = Str(CheckDBNull(dr(7), enumObjectType.IntType))
                                        rPROD_ID = Trim(CheckDBNull(dr(8), enumObjectType.StrType))
                                    Catch ex As Exception
                                    End Try
                                End While
                            End If
                            Try
                                dr.Close()
                                dr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                        End Try

                        If debug = "Y" Then
                            mydebuglog.Debug("  Results found--")
                            mydebuglog.Debug("   rQUEUE_ID: " & rQUEUE_ID)
                            mydebuglog.Debug("   rCERT_POOL_ID: " & rCERT_POOL_ID)
                            mydebuglog.Debug("   rCONTACT_ID: " & rCONTACT_ID & " - ConId: " & ConId)
                            mydebuglog.Debug("   rCRSE_ID: " & rCRSE_ID & " - CrseId: " & CrseId)
                            mydebuglog.Debug("   rREG_ID: " & rREG_ID)
                            mydebuglog.Debug("   rGENERATED: " & rGENERATED)
                            mydebuglog.Debug("   rDESTINATION: " & rDESTINATION & " - OutputDest: " & OutputDest)
                            mydebuglog.Debug("   rDOC_ID: " & rDOC_ID & " - ExistDocId: " & ExistDocId)
                            mydebuglog.Debug("   rPROD_ID: " & rPROD_ID & " - ProdId: " & ProdId & vbCrLf)
                        End If

                        ' Test information for duplicate call
                        If CrseId = rCRSE_ID Then
                            If ConId = rCONTACT_ID Then
                                If ProdId = rPROD_ID Then
                                    If OutputDest = rDESTINATION Then
                                        If Val(ExistDocId) = Val(rDOC_ID) Then
                                            ' Don't generate another certificate..
                                            If debug = "Y" Then mydebuglog.Debug("  BYPASSING GENERATION OF NEW PRODUCT" & vbCrLf)
                                            If QueueId <> rQUEUE_ID Then
                                                ' Remove the queued entry that was just created
                                                SqlS = "DELETE FROM siebeldb.dbo.CX_CERT_PROD_QUEUE WHERE ROW_ID='" & QueueId & "'"
                                                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Remove the new queued entry: " & vbCrLf & SqlS & vbCrLf)
                                                Try
                                                    cmd.CommandText = SqlS
                                                    returnv = cmd.ExecuteNonQuery()
                                                Catch ex As Exception
                                                End Try
                                            End If

                                            If rDESTINATION <> "image" Then
                                                NumFiles = 1
                                                QueueId = rQUEUE_ID
                                                dmsdocs(1, 18) = rGENERATED
                                                GoTo CloseOut
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' Lock next free certification number/id
                    SqlS = "UPDATE siebeldb.dbo.CX_CERT_PROD_ID_POOL " & _
                        "SET PROD_QUEUE_ID='" & QueueId & "' " & _
                        "WHERE ROW_ID IN " & _
                        "(SELECT TOP 1 ROW_ID " & _
                        "FROM siebeldb.dbo.CX_CERT_PROD_ID_POOL " & _
                        "WHERE JURIS_ID='" & JurisId & "' AND PROD_RESULT_ID='' AND PROD_QUEUE_ID='' ORDER BY JURIS_CERT_ID ASC)"
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Lock next free CX_CERT_PROD_ID_POOL: " & vbCrLf & SqlS & vbCrLf)
                    Try
                        cmd.CommandText = SqlS
                        returnv = cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        errmsg = errmsg & "Unable to lock next free CX_CERT_PROD_ID_POOL: " & ex.ToString & vbCrLf
                    End Try

                    ' Locate assigned free certification number/id
                    SqlS = "SELECT TOP 1 ROW_ID, JURIS_CERT_ID, PROD_QUEUE_ID " & _
                        "FROM siebeldb.dbo.CX_CERT_PROD_ID_POOL " & _
                       "WHERE PROD_QUEUE_ID='" & QueueId & "' AND PROD_RESULT_ID=''"
                    If debug = "Y" Then mydebuglog.Debug("  Get juris certification number: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    ID_POOL_ID = Trim(CheckDBNull(dr(0), enumObjectType.StrType))
                                    CERT_NUM = Trim(CheckDBNull(dr(1), enumObjectType.StrType))
                                    If debug = "Y" Then
                                        mydebuglog.Debug("   ID_POOL_ID: " & ID_POOL_ID)
                                        mydebuglog.Debug("   CERT_NUM: " & CERT_NUM)
                                    End If
                                    temp = Trim(CheckDBNull(dr(2), enumObjectType.StrType))
                                    If temp = QueueId And temp <> "" And CERT_NUM <> "" Then
                                        ' A number WAS found - Update the query
                                        CERT_QUERY = CERT_QUERY.Replace("[CERTIFICATION NUMBER]", CERT_NUM)
                                    Else
                                        CERT_NUM = ""
                                        ID_POOL_ID = ""
                                    End If
                                Catch ex As Exception
                                End Try
                            End While
                        End If
                        Try
                            dr.Close()
                            dr = Nothing
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try

                    ' Determine if a free number if available or not                    
                    If CERT_NUM = "" Then
                        ' ----
                        ' A number WAS NOT found
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & " A number WAS NOT found")

                        ' Create a new CX_CERT_PROD_QUEUE.ROW_ID
                        NewQueueId = LoggingService.GenerateRecordId("CX_CERT_PROD_QUEUE", "N", debug)
                        If NewQueueId = "" Then
                            NewQueueId = Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65)
                        End If
                        If debug = "Y" Then mydebuglog.Debug("   ... NewQueueId: " & NewQueueId)

                        ' Create a new CX_CERT_PROD_QUEUE record
                        SqlS = "INSERT INTO siebeldb.dbo.CX_CERT_PROD_QUEUE " & _
                        "(CONFLICT_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,PROD_TYPE,TRAIN_TYPE" & _
                        ",CRSE_ID,SRC_ID,IDENT_START,IDENT_END,NOTIFY_FLG,ATTACH_FLG,MULTI_OUT_FLG,FORMAT,DOMAIN" & _
                        ",JURIS_ID,CON_ID,OU_ID,DEST_CODE,EXECUTED,ROW_ID,NO_RESULTS_FLG,ACCESS_FLG,REQD_FLG" & _
                        ",PROD_ID,SPECIAL_MSG,DOC_ID) " & _
                        "VALUES(0,GETDATE(),'0-1',GETDATE(),'0-1',0,'" & TypeProd & "','" & SkillLevel & "'" & _
                        ",'" & CrseId & "','" & SrcId & "',"
                        If IdentStart = "" Then temp = "NULL" Else temp = IdentStart
                        SqlS = SqlS & temp & ","
                        If IdentEnd = "" Then temp = "NULL" Else temp = IdentEnd
                        SqlS = SqlS & temp & ",'Y','" & AttachFlg & "','" & MultiFlg & "','" & OutFormat & "','" & Domain & "'" & _
                        ",'" & JurisId & "','" & ConId & "','" & OrgId & "','" & OutputDest & "',NULL,'" & NewQueueId & "','R','" & AccessFlg & "','" & ReqdFlg & "'" & _
                        ",'" & ProdId & "','" & SpecialMsg & "',"
                        If ExistDocId = "" Then temp = "NULL" Else temp = ExistDocId
                        SqlS = SqlS & temp & ")"
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Creating a cloned CX_CERT_PROD_QUEUE entry: " & vbCrLf & SqlS & vbCrLf)
                        Try
                            cmd.CommandText = SqlS
                            returnv = cmd.ExecuteNonQuery()
                        Catch ex As Exception
                            errmsg = errmsg & "Unable to insert CX_CERT_PROD_QUEUE record: " & ex.ToString & vbCrLf
                        End Try

                        ' Verify queue created
                        SqlS = "SELECT ROW_ID FROM siebeldb.dbo.CX_CERT_PROD_QUEUE WHERE ROW_ID='" & NewQueueId & "'"
                        If debug = "Y" Then mydebuglog.Debug("  Verify new product queue entry: " & vbCrLf & SqlS)
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                        If temp <> "" Then NewQueueId = temp
                                    Catch ex As Exception
                                        errmsg = errmsg & "Error reading new CX_CERT_PROD_QUEUE record: " & ex.ToString
                                    End Try
                                End While
                            End If
                            Try
                                dr.Close()
                                dr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                            errmsg = errmsg & "Error reading new CX_CERT_PROD_QUEUE record: " & ex.ToString
                        End Try

                        ' Generate a new CX_CERT_PROD_ID_POOL record
                        If NewQueueId <> "" Then
                            ' Create a new ID
                            NEW_POOL_ID = LoggingService.GenerateRecordId("CX_CERT_PROD_ID_POOL", "N", debug)
                            If NEW_POOL_ID = "" Then
                                NEW_POOL_ID = Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65)
                            End If
                            If debug = "Y" Then mydebuglog.Debug("   ... NEW_POOL_ID: " & NEW_POOL_ID)

                            ' Insert the blank record to be used by the staff member
                            SqlS = "INSERT INTO siebeldb.dbo.CX_CERT_PROD_ID_POOL " & _
                                "(CONFLICT_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY, " & _
                                "MODIFICATION_NUM,ROW_ID,JURIS_ID,PROD_QUEUE_ID ) " & _
                                "VALUES (0, GETDATE(), '0-1', GETDATE(), '0-1',  " & _
                                "0,'" & NEW_POOL_ID & "','" & JurisId & "','" & NewQueueId & "')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Creating a CX_CERT_PROD_ID_POOL record: " & vbCrLf & SqlS & vbCrLf)
                            Try
                                cmd.CommandText = SqlS
                                returnv = cmd.ExecuteNonQuery()
                            Catch ex As Exception
                                errmsg = errmsg & "Unable to insert CX_CERT_PROD_ID_POOL record: " & ex.ToString & vbCrLf
                            End Try

                            ' Verify pool id record created
                            SqlS = "SELECT ROW_ID FROM siebeldb.dbo.CX_CERT_PROD_ID_POOL WHERE ROW_ID='" & NEW_POOL_ID & "'"
                            If debug = "Y" Then mydebuglog.Debug("  Verify pool id record created: " & vbCrLf & SqlS)
                            Try
                                cmd.CommandText = SqlS
                                dr = cmd.ExecuteReader()
                                If Not dr Is Nothing Then
                                    While dr.Read()
                                        Try
                                            temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                            If temp <> "" Then NEW_POOL_ID = temp
                                            ID_POOL_ID = ""
                                            CERT_NUM = ""
                                        Catch ex As Exception
                                            errmsg = errmsg & "Error reading new CX_CERT_PROD_ID_POOL record: " & ex.ToString
                                        End Try
                                    End While
                                End If
                                Try
                                    dr.Close()
                                    dr = Nothing
                                Catch ex As Exception
                                End Try
                            Catch ex As Exception
                                errmsg = errmsg & "Error reading new CX_CERT_PROD_ID_POOL record: " & ex.ToString
                            End Try

                        End If

                        ' Generate a MESSAGE 0155 to JURIS_CERT_EMAIL
                        If JURIS_CERT_EMAIL <> "" Then
                            MsgXml = "<messages>" & _
                            "<message send_to=""" & JURIS_CERT_EMAIL & """ send_from=""" & JURIS_CERT_EMAIL & """ from_name="""" from_id="""" to_id="""">" & _
                            "<JURISDICTION>" & Trim(JURISDICTION) & "</JURISDICTION>" & _
                            "<REPORT_NAME>" & Trim(REPORT_NAME) & "</REPORT_NAME>" & _
                            "<SrcId>" & Trim(SrcId) & "</SrcId>" & _
                            "<NEW_POOL_ID>" & Trim(NEW_POOL_ID) & "</NEW_POOL_ID>" & _
                            "</message>" & _
                            "</messages>"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Sending MESSAGE 0155: ")
                            Try
                                If MsgXml <> "" Then
                                    results = XsltMerge(MsgXml, "751", "751", "EMAIL", LANG_CD, debug, mydebuglog)
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & "Error generating MESSAGE 0155: " & ex.ToString & vbCrLf & " in XML: " & vbCrLf & MsgXml & vbCrLf
                            End Try
                        End If

                        ' Set the product id to the temporary product id if applicable
                        If TEMP_PROD_ID <> "" Then

                            ' First make sure we have a product
                            SqlS = "SELECT P.ROW_ID, P.RFILENAME, P.CERT_TYPE, P.CERT_QUERY, P.NAME, P.DEF_FORMAT, " & _
                            "P.SPECIAL_NOTICE, P.RES_X, P.RES_Y, P.WIDTH, P.HEIGHT " & _
                            "FROM siebeldb.dbo.CX_CERT_PROD P " & _
                            "LEFT OUTER JOIN siebeldb.dbo.CX_CERT_PROD_CRSE PC ON PC.CERT_ID=P.ROW_ID " & _
                            "WHERE P.ROW_ID='" & TEMP_PROD_ID & "' AND PC.CRSE_ID='" & CrseId & "'"
                            If debug = "Y" Then mydebuglog.Debug("  Get certification product: " & vbCrLf & SqlS)
                            Try
                                cmd.CommandText = SqlS
                                dr = cmd.ExecuteReader()
                                If Not dr Is Nothing Then
                                    While dr.Read()
                                        Try
                                            temp = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                            If temp <> "" Then
                                                ProdId = temp
                                                REP_FILENAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                                CERT_TYPE = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                                CERT_QUERY = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                                REPORT_NAME = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                                                If OutFormat = "" Then OutFormat = Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString
                                                SPECIAL_NOTICE = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                                                PResX = CheckDBNull(dr(7), enumObjectType.IntType)
                                                PResY = CheckDBNull(dr(8), enumObjectType.IntType)
                                                PWidth = CheckDBNull(dr(9), enumObjectType.IntType)
                                                PHeight = CheckDBNull(dr(10), enumObjectType.IntType)
                                            Else
                                                ProdId = ""
                                            End If
                                        Catch ex As Exception
                                            errmsg = errmsg & "Error locating product for this course. " & ex.ToString
                                        End Try
                                    End While
                                Else
                                    errmsg = errmsg & "Error locating product for this course. "
                                    results = "Failure"
                                    GoTo ErrorQueue
                                End If
                                Try
                                    dr.Close()
                                    dr = Nothing
                                Catch ex As Exception
                                End Try
                            Catch ex As Exception
                                errmsg = errmsg & "Error locating product for this course. " & ex.ToString
                                results = "Failure"
                                GoTo ErrorQueue
                            End Try

                            ' If report not found, then obviously there is an issue
                            If ProdId = "" Then
                                errmsg = errmsg & "The product was not found for this course"
                                GoTo ErrorQueue
                            End If

                            ' Compute the bounding rectangle
                            PWidth = PWidth * 3
                            PHeight = PHeight * 3
                            Dim CropRect2 As New Rectangle(0, 0, PWidth, PHeight)
                            CropRect = CropRect2

                            ' Debug product information found
                            If debug = "Y" Then
                                mydebuglog.Debug("  Product found--")
                                mydebuglog.Debug("   ProdId: " & ProdId)
                                mydebuglog.Debug("   REP_FILENAME: " & REP_FILENAME)
                                mydebuglog.Debug("   CERT_TYPE: " & CERT_TYPE)
                                mydebuglog.Debug("   CERT_QUERY: " & vbCrLf & CERT_QUERY)
                                mydebuglog.Debug("   Resolution: " & PResX.ToString & " x " & PResY.ToString)
                                mydebuglog.Debug("   Dimensions: " & PWidth.ToString & " x " & PHeight.ToString)
                            End If

                        End If

                    End If
                End If

                ' -----
                ' DETERMINE CERTIFICATION TYPE IF NEEDED
                If SkillLevel = "" Or CERT_TYPE = "" Then
                    If CERT_TYPE = "" And SkillLevel <> "" Then CERT_TYPE = SkillLevel
                    If CERT_TYPE <> "" And SkillLevel = "" Then SkillLevel = CERT_TYPE
                    If CERT_TYPE = "" And SkillLevel = "" Then
                        SqlS = "SELECT SKILL_LEVEL_CD " & _
                        "FROM siebeldb.dbo.S_CRSE  " & _
                        "WHERE ROW_ID='" & CrseId & "'"
                        If debug = "Y" Then mydebuglog.Debug("  Get course type: " & vbCrLf & SqlS)
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Found record on query")
                                        CERT_TYPE = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                        SkillLevel = CERT_TYPE
                                    Catch ex As Exception
                                        errmsg = errmsg & "Error locating course. " & ex.ToString
                                    End Try
                                End While
                            Else
                                errmsg = errmsg & "The course was not found." & vbCrLf
                                results = "Failure"
                                GoTo CloseOut
                            End If
                            Try
                                dr.Close()
                                dr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                            errmsg = errmsg & "Error locating course. " & ex.ToString
                            results = "Failure"
                            GoTo CloseOut
                        End Try
                    End If
                End If

                ' -----
                ' VERIFY THAT WE HAVE ALL OF THE FIELDS WE NEED
                If TypeProd = "" Or SkillLevel = "" Or CrseId = "" Or ProdId = "" Then
                    errmsg = errmsg & "Critical information missing.  Cannot generate product." & vbCrLf
                    results = "Failure"
                    GoTo ErrorQueue
                End If

                ' -----
                ' VERIFY REPORT EXISTS
                ' Check to see if we have the report 
                If REP_FILENAME <> "" Then
                    ' Open Report and Login
                    Try
                        If debug = "Y" Then mydebuglog.Debug("   Generating: " & REPORT_NAME)
                        If debug = "Y" Then mydebuglog.Debug("   Report: " & basepath & "reports\" & REP_FILENAME)
                        'Report = New ReportDocument()
                        'Report.Load(basepath & "reports\" & REP_FILENAME, OpenReportMethod.OpenReportByTempCopy)
                        '*** Open report from SAP RAS Server
                        Report = New CrystalDecisions.ReportAppServer.ClientDoc.ReportClientDocument
                        Dim boReportFolderID As String = ConfigurationManager.AppSettings.Get("GenCertProd_RASFolderID")
                        If Not openRASReport(boReportFolderID, REPORT_NAME, REP_FILENAME, Report, mydebuglog, debug) Then
                            errmsg = errmsg & vbCrLf & "Error opening CR report from SAP RAS. "
                            results = "Failure"
                            GoTo CloseOut
                        Else
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  CR Report " & Report.SummaryInfo.Title & " opened successfully")
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & "Unable to open report or login to data server " & ex.ToString
                        results = "Failure"
                        GoTo CloseOut
                    End Try
                End If

                ' -----
                ' GET ACCOUNT INFORMATION FOR PASSPORTS DESCRIPTIONS
                Dim ACCT_NAME, ACCT_LOC As String
                ACCT_NAME = ""
                ACCT_LOC = ""
                If TypeProd = "P" And OrgId <> "" Then
                    SqlS = "SELECT NAME, LOC " & _
                    "FROM siebeldb.dbo.S_ORG_EXT " & _
                    "WHERE ROW_ID='" & OrgId & "'"
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Get Account Information at " & Now.ToString)
                    If debug = "Y" Then mydebuglog.Debug("  Get account name for passport description: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    ACCT_NAME = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                    ACCT_LOC = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                Catch ex As Exception
                                    errmsg = errmsg & "Error getting account name for passport description. " & ex.ToString
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "Error getting account name for passport description."
                            dr.Close()
                            dr = Nothing
                            cmd = Nothing
                            results = "Failure"
                            GoTo CloseOut
                        End If
                        Try
                            dr.Close()
                            dr = Nothing
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                        errmsg = errmsg & "Error getting account name for passport description. " & ex.ToString
                    End Try
                    Try
                        If debug = "Y" Then
                            mydebuglog.Debug("  > ACCT_NAME: " & ACCT_NAME)
                            mydebuglog.Debug("  > ACCT_LOC: " & ACCT_LOC)
                        End If
                    Catch ex As Exception
                        errmsg = errmsg & "Account name and/or location includes non-standard characters. "
                    End Try
                End If

                ' -----
                ' IMPORT THE FIELD DEFINITIONS
                '  Build from the data definition file, and create the data table to store the query results
                '  If specified predefined field names are found, then set an index to the column
                '  containing that fieldname.  The predefined field names supported are:

                '       Field Name / Index : Description

                '       DOC_KEY / DmsKeyField : The data to be stored to Documents.description
                '       CON_ID / DmsConId : An association key to S_CONTACT.ROW_ID
                '       JURIS_ID / DmsJurisId : An association key to CX_JURISDICTION_X.ROW_ID 
                '       ORG_ID / DmsOrgId : An association key to S_ORG_EXT.ROW_ID
                '       WSHOP_ID / DmsWshopId : An association key to S_CRSE_OFFR.ROW_ID
                '       SESS_ID /DmsSessId : An association key to CX_SESSIONS_X.ROW_ID
                '       OFFR_ID /DmsOffrId : An association key to CX_TRAIN_OFFR.ROW_ID
                '       PART_ID / DmsPartId : An association key to CX_SESS_PART_X.ROW_ID
                '       WREG_ID / DmsWregId : An association key to S_CRSE_REG.ROW_ID
                '       SREG_ID / DmsSregId : An association key to CX_SESS_REG.ROW_ID
                '       EXAM_ID / DmsExamId : An association key to S_CRSE_TSTRUN.ROW_ID
                '       CERT_ID / DmsCertId : An association key to S_CURRCLM_PER.ROW_ID
                '       TRAINER_ID / DmsTrainerId : An association key to S_CONTACT.X_TRAINER_ID

                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Preparing Data Objects")
                NumCols = 0
                dt = New DataTable
                CDODefFn = basepath & "reports\" & Replace(REP_FILENAME, ".rpt", ".ttx")
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Report Definition: " & CDODefFn)
                Try
                    CDOStream = System.IO.File.OpenText(CDODefFn)
                    While CDOStream.Peek <> -1
                        strLine = CDOStream.ReadLine()
                        If Len(strLine) <> 0 And Right(strLine, 2) <> "%%" Then
                            temp = Split(strLine, vbTab)(0)
                            If temp = "DOC_KEY" Then DmsKeyField = NumCols
                            If temp = "CON_ID" Then DmsConId = NumCols
                            If temp = "JURIS_ID" Then DmsJurisId = NumCols
                            If temp = "ORG_ID" Then DmsOrgId = NumCols
                            If temp = "WSHOP_ID" Then DmsWshopId = NumCols
                            If temp = "SESS_ID" Then DmsSessId = NumCols
                            If temp = "OFFR_ID" Then DmsOffrId = NumCols
                            If temp = "PART_ID" Then DmsPartId = NumCols
                            If temp = "WREG_ID" Then DmsWregId = NumCols
                            If temp = "SREG_ID" Then DmsSregId = NumCols
                            If temp = "EXAM_ID" Then DmsExamId = NumCols
                            If temp = "CERT_ID" Then DmsCertId = NumCols
                            If temp = "TRAINER_ID" Then DmsTrainerId = NumCols
                            If debug = "Y" Then mydebuglog.Debug("   > Field: #" & Str(NumCols) & " : " & temp)
                            NumCols = NumCols + 1
                            CDOFields(NumCols) = temp
                        End If
                    End While
                    CDOStream.DiscardBufferedData()
                    CDOStream.Close()
                    CDOStream = Nothing
                Catch ex As Exception
                    errmsg = errmsg & "Error opening report definition. " & ex.ToString
                    results = "Failure"
                    GoTo ErrorQueue
                End Try

                '  Debug output document related information found
                If debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & " Indexes found" & vbCrLf & "   > DOC_KEY:" & DmsKeyField)
                    mydebuglog.Debug("   > CON_ID:" & DmsConId)
                    mydebuglog.Debug("   > JURIS_ID:" & DmsJurisId)
                    mydebuglog.Debug("   > ORG_ID:" & DmsOrgId)
                    mydebuglog.Debug("   > WSHOP_ID:" & DmsWshopId)
                    mydebuglog.Debug("   > SESS_ID:" & DmsSessId)
                    mydebuglog.Debug("   > OFFR_ID:" & DmsOffrId)
                    mydebuglog.Debug("   > PART_ID:" & DmsPartId)
                    mydebuglog.Debug("   > WREG_ID:" & DmsWregId)
                    mydebuglog.Debug("   > SREG_ID:" & DmsSregId)
                    mydebuglog.Debug("   > EXAM_ID:" & DmsExamId)
                    mydebuglog.Debug("   > TRAINER_ID:" & DmsTrainerId)
                    mydebuglog.Debug("   > CERT_ID:" & DmsCertId & vbCrLf)
                End If

                ' -----
                ' BUILD PRODUCT RESULTS QUERY
                '  Take the query supplied in XML or in the CX_CERT_PROD table and add the
                '  parameters.  Note that the fields returned by the query MUST match the
                '  fields in the reports ttx file just retrieved

                '  Determine if we have a certification product query.  If not, create one
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Determining product query and executing it at " & Now.ToString)
                If CERT_QUERY = "" Then
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Determining query for a " & SkillLevel & " product type '" & TypeProd & "'")
                    Select Case SkillLevel
                        Case "Participant"
                            Select Case TypeProd
                                Case "R"
                                    CERT_QUERY = "SELECT DATENAME(month, SP.CREATED)+' '+CAST(DATEPART(Dd,SP.CREATED) AS VARCHAR)+', '+ " & _
                                    "CAST(DATEPART(Yy,SP.CREATED) AS VARCHAR) AS COMPLETED, " & _
                                    "CR.NAME AS COURSE, RTRIM(C.FST_NAME)+' '+C.LAST_NAME AS FULLNAME, " & _
                                    "(SELECT CASE WHEN O.ROW_ID IS NOT NULL THEN O.NAME+(SELECT CASE WHEN O.LOC IS NOT " & _
                                    "NULL AND O.LOC<>'' THEN ', '+O.LOC ELSE '' END) ELSE '' END) AS ORGANIZATION, " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.ADDR ELSE PA.ADDR END) AS ADDR, " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.CITY ELSE PA.CITY END)+' '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.STATE ELSE PA.STATE END)+', '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.ZIPCODE ELSE PA.ZIPCODE END)+' '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.COUNTRY ELSE PA.COUNTRY END) AS CSZ, " & _
                                    "RTRIM(C.FST_NAME)+' '+C.LAST_NAME+', Participation No. '+CAST(SP.PART_NUM AS VARCHAR) AS DOC_KEY, " & _
                                    "C.ROW_ID AS CON_ID, O.ROW_ID AS ORG_ID, R.JURIS_ID, R.ROW_ID AS SREG_ID, SP.ROW_ID AS PART_ID, " & _
                                    "R.TRAIN_OFFR_ID AS OFFR_ID " & _
                                    "FROM siebeldb.dbo.CX_SESS_REG R  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=R.CONTACT_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=R.SESS_PART_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT O ON O.ROW_ID=R.OU_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=C.PR_PER_ADDR_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG CA ON CA.ROW_ID=C.PR_OU_ADDR_ID " & _
                                    "WHERE R.ROW_ID='[ROW_ID]'"
                                Case "C"
                                    CERT_QUERY = "EXEC reports.dbo.OpenHCIKeys; SELECT CONVERT(VARCHAR,S.CREATED,101) AS CERTIFIED_DATE, W.NAME AS CRSE_NAME, " & _
                                    "CONVERT(VARCHAR,S.EXP_DATE,101) AS EXPIRATION_DATE, LTRIM(CAST(S.PART_NUM AS VARCHAR)) AS PART_NUM, " & _
                                    "S.PERMIT_NUM,  S.STATE_MSG,  P.ACCOUNT_NAME AS PART_ACCT_NAME, " & _
                                    "P.ACCOUNT_LOC AS PART_ACCT_LOC, P.ADDR AS PART_ADDR, CONVERT(VARCHAR,CONVERT(DATETIME,reports.dbo.HCI_Decrypt(P.ENC_BIRTH_DT)),101) AS DOB, " & _
                                    "P.CITY AS PART_CITY, P.FST_NAME AS PART_FST_NAME, P.LAST_NAME AS PART_LAST_NAME, " & _
                                    "reports.dbo.HCI_Decrypt(P.ENC_SOC_SECURITY_NUM) AS PART_SSN, P.STATE AS PART_STATE, P.ZIPCODE AS PART_ZIP, " & _
                                    "P.COUNTRY AS PART_COUNTRY, J.NAME AS JURISDICTION,  J.JURISDICTION AS JURIS_CODE, " & _
                                    "RTRIM(P.FST_NAME)+' '+RTRIM(P.LAST_NAME)+', Participation No. '+CAST(S.PART_NUM AS VARCHAR) AS DOC_KEY, " & _
                                    "P.CON_ID AS CON_ID, S.OU_ID AS ORG_ID, S.JURIS_ID, S.ROW_ID AS PART_ID, S.SESS_ID, " & _
                                    "S.CRSE_TSTRUN_ID AS EXAM_ID " & _
                                    "FROM siebeldb.dbo.CX_SESS_PART_X S " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X J ON J.ROW_ID=S.JURIS_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_PARTICIPANT_X P ON S.PART_ID = P.ROW_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE W ON W.ROW_ID = S.CRSE_TST_ID " & _
                                    "WHERE S.ROW_ID='[ROW_ID]'"
                                Case "W"
                                    CERT_QUERY = ""
                                Case "P"
                                    CERT_QUERY = "SELECT LTRIM(CAST(S.MS_IDENT AS VARCHAR)) AS ID_NUMBER, R.ROW_ID AS REG_NUMBER, CR.NAME AS PROGRAM, " & _
                                    "'Registration No. '+CAST(R.MS_IDENT AS VARCHAR) AS DOC_KEY, " & _
                                    "R.CONTACT_ID AS CON_ID, R.OU_ID AS ORG_ID, R.JURIS_ID, S.ROW_ID AS OFFR_ID, CONVERT(VARCHAR,R.EXP_DT,101) AS EXP_DT  " & _
                                    "FROM siebeldb.dbo.CX_SESS_REG R  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_TRAIN_OFFR S ON S.ROW_ID=R.TRAIN_OFFR_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=R.CRSE_ID  " & _
                                    "WHERE R.ROW_ID='[ROW_ID]'"
                                Case "V"
                                    CERT_QUERY = "SELECT T.ROW_ID AS REG_NUMBER, " & _
                                    "E.ROW_ID AS ID_NUMBER, E.NAME AS EXAM,  " & _
                                    "'Exam No. '+CAST(T.MS_IDENT AS VARCHAR) AS DOC_KEY, T.PERSON_ID AS CON_ID, " & _
                                    "C.PR_DEPT_OU_ID AS ORG_ID, E.X_JURIS_ID AS JURIS_ID, T.CRSE_OFFR_ID AS WSHOP_ID, " & _
                                    "T.X_PART_ID AS PART_ID, T.ROW_ID AS EXAM_ID " & _
                                    "FROM siebeldb.dbo.S_CRSE_TSTRUN T " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST E ON E.ROW_ID=T.CRSE_TST_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=T.PERSON_ID " & _
                                    "WHERE T.ROW_ID='[ROW_ID]'"
                                Case "F"
                                    CERT_QUERY = "EXEC reports.dbo.OpenHCIKeys; SELECT P.FST_NAME AS PART_FST_NAME, P.LAST_NAME AS PART_LAST_NAME, " & _
                                    "CONVERT(VARCHAR,SS.DATE_HELD,101) AS DATE_HELD, CONVERT(VARCHAR,S.EXP_DATE,101) AS EXPIRATION_DATE, " & _
                                    "CONVERT(VARCHAR,S.CREATED,101) AS CERTIFIED_DATE, W.NAME AS CRSE_NAME, " & _
                                    "LTRIM(CAST(S.PART_NUM AS VARCHAR)) AS PART_NUM, P.ACCOUNT_NAME AS PART_ACCT_NAME, " & _
                                    "P.ACCOUNT_LOC AS PART_ACCT_LOC, P.ADDR AS PART_ADDR, P.CITY AS PART_CITY, " & _
                                    "CONVERT(VARCHAR,CONVERT(DATETIME,reports.dbo.HCI_Decrypt(P.ENC_BIRTH_DT)),101) AS DOB, reports.dbo.HCI_Decrypt(P.ENC_SOC_SECURITY_NUM) AS PART_SSN, " & _
                                    "P.STATE AS PART_STATE, P.ZIPCODE AS PART_ZIP, P.COUNTRY AS PART_COUNTRY, " & _
                                    "A.COUNTY AS PART_COUNTY, J.NAME AS JURISDICTION,  J.JURISDICTION AS JURIS_CODE," & _
                                    "T.FST_NAME AS TRAINER_FST_NAME, T.LAST_NAME AS TRAINER_LAST_NAME, T.X_TRAINER_NUM AS TRAINER_NUM, " & _
                                    "RTRIM(P.FST_NAME)+' '+RTRIM(P.LAST_NAME)+', Participation No. '+CAST(S.PART_NUM AS VARCHAR)+', Session No. '+CAST(SS.SESS_ID AS VARCHAR) AS DOC_KEY, " & _
                                    "T.ROW_ID AS CON_ID, S.OU_ID AS ORG_ID, S.JURIS_ID, S.ROW_ID AS PART_ID, SS.ROW_ID AS SESS_ID " & _
                                    "FROM siebeldb.dbo.CX_SESS_PART_X S " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_SESSIONS_X SS ON SS.ROW_ID=S.SESS_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT T ON T.ROW_ID=SS.PR_CON_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG A ON A.ROW_ID=S.ADDR_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X J ON J.ROW_ID=S.JURIS_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_PARTICIPANT_X P ON S.PART_ID = P.ROW_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE W ON W.ROW_ID = S.CRSE_TST_ID " & _
                                    "WHERE S.ROW_ID='[ROW_ID]'"
                                Case "T"
                                    CERT_QUERY = "SELECT DATENAME(month, T.TEST_DT)+' '+CAST(DATEPART(Dd,T.TEST_DT) AS VARCHAR)+', '+ " & _
                                    "CAST(DATEPART(Yy,T.TEST_DT) AS VARCHAR) AS COMPLETED, " & _
                                    "E.NAME AS PROGRAM, RTRIM(C.FST_NAME)+' '+C.LAST_NAME AS FULLNAME,  " & _
                                    "(SELECT CASE WHEN O.ROW_ID IS NOT NULL THEN O.NAME+(SELECT CASE WHEN O.LOC IS NOT  " & _
                                    "NULL AND O.LOC<>'' THEN ', '+O.LOC ELSE '' END) ELSE '' END) AS ORGANIZATION, " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.ADDR ELSE PA.ADDR END) AS ADDR, " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.CITY ELSE PA.CITY END)+' '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.STATE ELSE PA.STATE END)+', '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.ZIPCODE ELSE PA.ZIPCODE END)+' '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.COUNTRY ELSE PA.COUNTRY END) AS CSZ, " & _
                                    "(SELECT CASE WHEN TC.ROW_ID IS NOT NULL THEN (RTRIM(C.FST_NAME)+' '+C.LAST_NAME+', Trainer No. '+TC.X_CERT_ID) " & _
                                    "ELSE (RTRIM(C.FST_NAME)+' '+C.LAST_NAME+', Participation No. '+CAST(SP.PART_NUM AS VARCHAR)) END) AS DOC_KEY, " & _
                                    "C.ROW_ID AS CON_ID, O.ROW_ID AS ORG_ID, (SELECT CASE WHEN TC.ROW_ID IS NOT NULL THEN J.ROW_ID ELSE SP.JURIS_ID END) AS JURIS_ID, " & _
                                    "T.ROW_ID AS EXAM_ID, SP.ROW_ID AS PART_ID, TC.ROW_ID AS CERT_ID " & _
                                    "FROM siebeldb.dbo.S_CRSE_TSTRUN T " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST E ON E.ROW_ID=T.CRSE_TST_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=E.CRSE_ID  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=T.PERSON_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=T.X_PART_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT O ON O.ROW_ID=C.PR_DEPT_OU_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=C.PR_PER_ADDR_ID  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG CA ON CA.ROW_ID=C.PR_OU_ADDR_ID  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CURRCLM_PER TC ON TC.X_CRSE_TSTRUN_ID=T.ROW_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X J ON J.STATE=TC.X_STATE " & _
                                    "WHERE T.ROW_ID='[ROW_ID]'"
                                Case Else
                                    errmsg = errmsg & "Unable to find this type of product"
                                    results = "Failure"
                                    GoTo CloseOut
                            End Select
                        Case "Trainer"
                            Select Case TypeProd
                                Case "C"
                                    CERT_QUERY = "~S_CURRCLM_PER~"
                                Case "W"
                                    CERT_QUERY = "~S_CURRCLM_PER~"
                                Case "V"
                                    CERT_QUERY = "SELECT T.ROW_ID AS REG_NUMBER, " & _
                                    "E.ROW_ID AS ID_NUMBER, E.NAME AS EXAM,  " & _
                                    "'Exam No. '+CAST(T.MS_IDENT AS VARCHAR) AS DOC_KEY, T.PERSON_ID AS CON_ID, " & _
                                    "C.PR_DEPT_OU_ID AS ORG_ID, E.X_JURIS_ID AS JURIS_ID, T.CRSE_OFFR_ID AS WSHOP_ID, " & _
                                    "T.X_PART_ID AS PART_ID, T.ROW_ID AS EXAM_ID " & _
                                    "FROM siebeldb.dbo.S_CRSE_TSTRUN T " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST E ON E.ROW_ID=T.CRSE_TST_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=T.PERSON_ID " & _
                                    "WHERE T.ROW_ID='[ROW_ID]'"
                                Case "F"
                                    CERT_QUERY = "~S_CURRCLM_PER~"
                                Case "T"
                                    CERT_QUERY = "SELECT DATENAME(month, T.TEST_DT)+' '+CAST(DATEPART(Dd,T.TEST_DT) AS VARCHAR)+', '+ " & _
                                    "CAST(DATEPART(Yy,T.TEST_DT) AS VARCHAR) AS COMPLETED, " & _
                                    "E.NAME AS PROGRAM, RTRIM(C.FST_NAME)+' '+C.LAST_NAME AS FULLNAME,  " & _
                                    "(SELECT CASE WHEN O.ROW_ID IS NOT NULL THEN O.NAME+(SELECT CASE WHEN O.LOC IS NOT  " & _
                                    "NULL AND O.LOC<>'' THEN ', '+O.LOC ELSE '' END) ELSE '' END) AS ORGANIZATION, " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.ADDR ELSE PA.ADDR END) AS ADDR, " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.CITY ELSE PA.CITY END)+' '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.STATE ELSE PA.STATE END)+', '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.ZIPCODE ELSE PA.ZIPCODE END)+' '+ " & _
                                    "(SELECT CASE WHEN CA.ROW_ID IS NOT NULL THEN CA.COUNTRY ELSE PA.COUNTRY END) AS CSZ, " & _
                                    "(SELECT CASE WHEN TC.ROW_ID IS NOT NULL THEN (RTRIM(C.FST_NAME)+' '+C.LAST_NAME+', Trainer No. '+TC.X_CERT_ID) " & _
                                    "ELSE (RTRIM(C.FST_NAME)+' '+C.LAST_NAME+', Participation No. '+CAST(SP.PART_NUM AS VARCHAR)) END) AS DOC_KEY, " & _
                                    "C.ROW_ID AS CON_ID, O.ROW_ID AS ORG_ID, (SELECT CASE WHEN TC.ROW_ID IS NOT NULL THEN J.ROW_ID ELSE SP.JURIS_ID END) AS JURIS_ID, " & _
                                    "T.ROW_ID AS EXAM_ID, SP.ROW_ID AS PART_ID, TC.ROW_ID AS CERT_ID " & _
                                    "FROM siebeldb.dbo.S_CRSE_TSTRUN T " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST E ON E.ROW_ID=T.CRSE_TST_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE CR ON CR.ROW_ID=E.CRSE_ID  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=T.PERSON_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_SESS_PART_X SP ON SP.ROW_ID=T.X_PART_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ORG_EXT O ON O.ROW_ID=C.PR_DEPT_OU_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_PER PA ON PA.ROW_ID=C.PR_PER_ADDR_ID  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_ADDR_ORG CA ON CA.ROW_ID=C.PR_OU_ADDR_ID  " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CURRCLM_PER TC ON TC.X_CRSE_TSTRUN_ID=T.ROW_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.CX_JURISDICTION_X J ON J.STATE=TC.X_STATE " & _
                                    "WHERE T.ROW_ID='[ROW_ID]'"
                                Case Else
                                    errmsg = errmsg & "Unable to find this type of product"
                                    results = "Failure"
                                    GoTo CloseOut
                            End Select
                        Case "Master Trainer"
                            Select Case TypeProd
                                Case "V"
                                    CERT_QUERY = "SELECT T.ROW_ID AS REG_NUMBER, " & _
                                    "E.ROW_ID AS ID_NUMBER, E.NAME AS EXAM,  " & _
                                    "'Exam No. '+CAST(T.MS_IDENT AS VARCHAR) AS DOC_KEY, T.PERSON_ID AS CON_ID, " & _
                                    "C.PR_DEPT_OU_ID AS ORG_ID, E.X_JURIS_ID AS JURIS_ID, T.CRSE_OFFR_ID AS WSHOP_ID, " & _
                                    "T.X_PART_ID AS PART_ID, T.ROW_ID AS EXAM_ID " & _
                                    "FROM siebeldb.dbo.S_CRSE_TSTRUN T " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CRSE_TST E ON E.ROW_ID=T.CRSE_TST_ID " & _
                                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.ROW_ID=T.PERSON_ID " & _
                                    "WHERE T.ROW_ID='[ROW_ID]'"
                                Case "F"
                                    CERT_QUERY = ""
                                Case Else
                                    errmsg = errmsg & "Unable to find this type of product"
                                    results = "Failure"
                                    GoTo CloseOut
                            End Select
                        Case Else
                            errmsg = errmsg & "Unable to find this type of product"
                            results = "Failure"
                            GoTo CloseOut
                    End Select
                End If

                '  Apply the supplied query if applicable.  The query supplied will be the 
                '   just the "where" clause.
                If SrcQuery <> "" Then
                    ReportQuery = Left(CERT_QUERY, InStr(UCase(CERT_QUERY), "WHERE") - 1) & SrcQuery
                Else
                    ReportQuery = CERT_QUERY
                End If

                '  Apply the supplied parameters to the query
                If SrcId.Trim <> "" Then
                    ReportQuery = Replace(ReportQuery, "[ROW_ID]", SrcId)
                End If
                If IdentStart.Trim <> "" And IdentStart.Trim <> "0" Then
                    ReportQuery = Replace(ReportQuery, "ROW_ID='[ROW_ID]'", "MS_IDENT>=" & IdentStart)
                    If IdentEnd.Trim <> "" Then
                        ReportQuery = ReportQuery & " AND "
                        j = InStr(UCase(ReportQuery), "WHERE") + 5
                        temp = ReportQuery.Substring(j, 2)
                        ReportQuery = ReportQuery & temp & "MS_IDENT<=" & IdentEnd
                    End If
                    ReportQuery = ReportQuery & " ORDER BY R.MS_IDENT"
                Else
                    If IdentEnd.Trim <> "" And IdentEnd.Trim <> "0" Then
                        ReportQuery = Replace(ReportQuery, "ROW_ID='[ROW_ID]'", "MS_IDENT<=" & IdentStart)
                        ReportQuery = ReportQuery & " ORDER BY R.MS_IDENT"
                    End If
                End If

                'Ren Hou; Added to modify query for Encryption changes;
                If ReportQuery.Contains("OpenHCIKeys") Then
                    ReportQuery = ReportQuery + "; EXEC reports.dbo.CloseHCIKeys;"
                End If

                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Report query:" & vbCrLf & ReportQuery)

                '***** New codes for  ADODB connection to perform query for report data ******
                '**** ADODB  *****
                ' Open a connection without using a Data Source Name (DSN)
                Dim Cnxn1 As ADODB.Connection = New ADODB.Connection
                Cnxn1.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("adodb").ConnectionString
                If Cnxn1.ConnectionString = "" Then Cnxn1.ConnectionString = "Provider=SQLNCLI11; Server=HCIDBSQL\HCIDB; Database=;" _
                    & "User ID=sa;Password=k3v5c2!k3v5c2;" _
                    & "DataTypeCompatibility=80;ApplicationIntent=ReadOnly"
                '& "MARS Connection=True;"
                Cnxn1.Open()
                ' ADODB command
                Dim adocmd As ADODB.Command = New ADODB.Command
                adocmd.ActiveConnection = Cnxn1
                'Set timeout
                adocmd.CommandTimeout = 100
                ' Set the criteria
                Dim adoSqlStr As String = ReportQuery
                adocmd.CommandText = adoSqlStr
                adocmd.CommandType = ADODB.CommandTypeEnum.adCmdText

                ' Create a recordset by executing the command.
                Try
                    adors = adocmd.Execute
                Catch ex As Exception
                    errmsg = errmsg & "Unable to Exec ADODB Command: " & ex.ToString
                End Try
                'Count records
                RecordsRead = 0
                Do While Not adors.EOF
                    RecordsRead = RecordsRead + 1
                    adors.MoveNext()
                Loop
                If debug = "Y" Then mydebuglog.Debug("   > ADO Recordset found: " & RecordsRead.ToString)

                ' Set report data source to adors
                If RecordsRead > 0 Then
                    Try
                        Report.DatabaseController.SetDataSource(adors, Report.DataDefController.Database.Tables(0).Name, Report.DataDefController.Database.Tables(0).Name)
                    Catch ex2 As Exception
                        ' Write record that is causing error to xml document
                        errmsg = errmsg & "Error setting report datasource: " & ex2.ToString & vbCrLf
                        Try
                            adors.Save(mypath & "GenCertProd.xml", ADODB.PersistFormatEnum.adPersistXML)
                        Catch ex As Exception
                        End Try
                        If debug = "Y" Then mydebuglog.Debug(errmsg)
                        results = "Failure"
                        NumFiles = 0
                        GoTo CloseOut
                    End Try
                End If

                'Try
                '    '*** Load datatable with ado recordset
                '    ' Load to datatable
                '    Dim oledbDa As New OleDbDataAdapter
                '    mydebuglog.Debug("   oledbDa fill: " & oledbDa.Fill(dt, adors))
                '    If debug = "Y" Then mydebuglog.Debug("   > dt datatable found: " & dt.Rows.Count.ToString)
                'Catch ex As Exception
                '    errmsg = errmsg & "Unable to load ADO Recordset into data table."
                '    results = "Failure"
                '    GoTo CloseOut
                'End Try

                ' -----
                ' PERFORM QUERY AND SAVE TO THE DATA TABLE
                Try
                    If RecordsRead > 0 Then
                        NumRows = RecordsRead         ' Number of records in the result set
                        NumCols = adors.Fields.Count      ' Number of columns in each row
                        If debug = "Y" Then mydebuglog.Debug("   > Records found: " & NumRows.ToString)
                        If debug = "Y" Then mydebuglog.Debug("   > Columms found: " & NumCols.ToString)
                    Else
                        errmsg = errmsg & "No records were found for this query."
                        results = "Failure"
                        GoTo ErrorQueue
                    End If
                    Try
                        dr.Close()
                        dr = Nothing
                    Catch ex As Exception
                    End Try
                Catch ex As Exception
                    errmsg = errmsg & "Unable to load datareader into dataset."
                    results = "Failure"
                    GoTo CloseOut
                End Try

                '  Display results for debugging
                If debug = "Y" Then
                    Try
                        ' Display column names
                        temp = "| "
                        'For Each dtColumn In dt.Columns
                        For Each adoField In adors.Fields
                            temp = temp & adoField.Name & " | "
                        Next
                        mydebuglog.Debug("   Fields:   " & temp)

                        j = 0
                        'For Each dtRow In dt.Rows
                        adors.MoveFirst()
                        Do While (j < RecordsRead)
                            temp = "| "
                            j = j + 1
                            For Each adoField In adors.Fields
                                temp = temp & adoField.Value.ToString & " | "
                            Next
                            mydebuglog.Debug("   Row # " & (j).ToString & " : " & temp)
                            adors.MoveNext()
                        Loop
                        'Next
                    Catch ex As Exception
                    End Try
                End If

                '***** End ADODB  ****

                '  Error out if needed
                If NumRows = 0 Then
                    errmsg = errmsg & "No records were found for this query."
                    results = "Failure"
                    GoTo ErrorQueue
                End If

                ' -----
                ' STORE DATATABLE TO ADODB DATASET
                '  This moves the data cached to the datatable object into the CDO
                Dim LabelRows(NumRows - 1, NumCols - 1) As String
                Try
                    ' Transfer the datatable into a temporary array
                    j = 0
                    k = 0
                    'For Each dtRow In dt.Rows
                    adors.MoveFirst()
                    Do While (j < RecordsRead)
                        temp = "| "
                        'For Each dtColumn In dt.Columns
                        For Each adoField In adors.Fields
                            'If debug = "Y" Then mydebuglog.Debug(j.ToString & " " & k.ToString & " - " & dtRow(dtColumn))
                            LabelRows(j, k) = Trim(adoField.Value.ToString)
                            k = k + 1
                            'Next
                        Next
                        j = j + 1
                        k = 0
                        'Next
                        adors.MoveNext()
                    Loop


                Catch ex As Exception
                    errmsg = errmsg & "Unable to load data into the CDO from the Datatable: " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try

                ' -----
                ' DEFINE TEMPORARY FILE
                '  Everything gets exported into pdf and then converted from there to other formats
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Generating product from report at " & Now.ToString)
                If QueueId <> "" Then
                    OUT_FILENAME = QueueId & ".pdf"
                Else
CheckFN2:
                    ' Generate and check for uniqueness a random filename
                    temp = LoggingService.GenerateRecordId("CX_CERT_PROD_QUEUE", "N", debug)
                    OUT_FILENAME = temp & ".pdf"
                    OutputPath = basepath & "temp\" & OUT_FILENAME
                    If debug = "Y" Then mydebuglog.Debug("   ... checking to see if exists: " & OutputPath)
                    If (My.Computer.FileSystem.FileExists(OutputPath)) Then GoTo CheckFN2
                End If

                ' Define path to where temp file is stored
                OutputPath = basepath & "temp\" & OUT_FILENAME
                If debug = "Y" Then mydebuglog.Debug("  Temp File: " & OutputPath)

                '*** Export Report *****
                Try

                    'Set Export Options
                    myExportOptions = New CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptionsClass()
                    myExportOptions.ExportFormatType = CrystalDecisions.ReportAppServer.ReportDefModel.CrReportExportFormatEnum.crReportExportFormatPDF
                    Dim RasPDFExpOpts As ReportDefModel.PDFExportFormatOptions = New ReportDefModel.PDFExportFormatOptions()
                    RasPDFExpOpts.CreateBookmarksFromGroupTree = False
                    myExportOptions.FormatOptions = RasPDFExpOpts
                    'Output Report to stream
                    tempByteArray = Report.PrintOutputController.ExportEx(myExportOptions)
                    Dim byteStreamOutput As Byte() = tempByteArray.ByteArray
                    'Dim xmlPath As String = (("C:\" & rcd.DisplayName & "_") + XmlExportFormat.Name & ".") + XmlExportFormat.FileExtension
                    Dim out_fs As FileStream = New FileStream(OutputPath, FileMode.Create, FileAccess.ReadWrite)
                    Dim maxSize As Integer = byteStreamOutput.Length
                    out_fs.Write(byteStreamOutput, 0, maxSize)
                    out_fs.Close()
                Catch ex As Exception
                    errmsg = errmsg & "Error generating product report: " & ex.ToString & vbCrLf
                    If debug = "Y" Then mydebuglog.Debug(errmsg)
                    results = "Failure"
                    If QueueId <> "" Then
                        errmsg = errmsg & "Error generating product"
                        GoTo ErrorQueue
                    End If
                    GoTo CloseOut
                End Try

                '*** End export

                '' -----
                '' DEFINE PRINT OPTIONS - RESET
                'Report.PrintOptions.PrinterName = ""

                '' -----
                '' EXPORT REPORT TO TEMP FILE
                'myDiskFileDestinationOptions = New CrystalDecisions.Shared.DiskFileDestinationOptions
                'myDiskFileDestinationOptions.DiskFileName = OutputPath

                'myExportOptions = Report.ExportOptions
                'With myExportOptions
                '    .DestinationOptions = myDiskFileDestinationOptions
                '    .ExportDestinationType = ExportDestinationType.DiskFile
                '    .ExportFormatType = ExportFormatType.PortableDocFormat
                'End With

                'Try
                '    Report.Export()
                'Catch ex As Exception
                '    errmsg = errmsg & "Error generating product report: " & ex.ToString & vbCrLf
                '    If debug = "Y" Then mydebuglog.Debug(errmsg)
                '    results = "Failure"
                '    If QueueId <> "" Then
                '        errmsg = errmsg & "Error generating product"
                '        GoTo ErrorQueue
                '    End If
                '    GoTo CloseOut
                'End Try

                ' -----
                ' Close the report and related objects
                Try
                    Release(Database, True)
                    Database = Nothing
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Database)
                Catch ex As Exception
                End Try

                Try
                    Release(myExportOptions, True)
                    myExportOptions = Nothing
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(myExportOptions)
                Catch ex As Exception
                End Try

                Try
                    Release(myDiskFileDestinationOptions, True)
                    myDiskFileDestinationOptions = Nothing
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(myDiskFileDestinationOptions)
                Catch ex As Exception
                End Try

                Try
                    Report.Close()
                    Report.Dispose()
                    Report = Nothing
                    Release(Report, True)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Report)
                Catch ex As Exception
                End Try

                Try
                    ds.Clear()
                    ds = Nothing
                Catch ex As Exception
                End Try

                Try
                    dr.Close()
                    dr = Nothing
                Catch ex As Exception
                End Try

                Try
                    'close out ADODB objects
                    adors.Close()
                    adors = Nothing
                    adocmd = Nothing
                    Cnxn1.Close()
                    Cnxn1 = Nothing
                Catch ex As Exception
                End Try

                ' -----
                ' CONVERT TEMP FILE TO ANOTHER FORMAT IF APPLICABLE
                '  If MultiFlg="Y", then each page will be broken into a separate file.
                '  Otherwise, the number of files is determined by the format where TIFF and PDF
                '  are automatically one, and JPEG automatically the number of pages produced in
                '  the reporting output.  The file name convention is:
                '       [OutputPath]xxx1.jpg
                '       [OutputPath]xxx2.jpg
                '           ...
                Try
                    VerySetLicenseKey("VERYPDF-PDFTOOLKITSDK-087848")
                    PDFToImageSetCode("VERYPDF-PDF2IMAGE-1009018849")
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Convert product to other formats at " & Now.ToString)
                    Select Case OutFormat
                        Case "pdf"
                            ' Generate an OwnerPassword
                            PdfPassword = Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & _
                            Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Trim(Str(Minute(Now))) & Trim(Str(Second(Now))) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65) & Chr(Str(Int(Rnd() * 26)) + 65) & Chr(Str(Int(RandNum.NextDouble() * 26)) + 65)
                            If debug = "Y" Then mydebuglog.Debug("  Generated password: " & PdfPassword)

                            ' Split it up if MultiFlg is set to "Y", otherwise, encrypt it as-is
                            If MultiFlg = "Y" Then
                                OUT_FILENAME = Replace(OUT_FILENAME, ".pdf", "")   ' Change filename
                                NumFiles = NumRows

                                ' Create temp directory
                                MkDir(basepath & "temp\" & OUT_FILENAME & "\")

                                ' Split pdf into constituent pieces into temporary directory
                                If debug = "Y" Then mydebuglog.Debug("  Splitting to: " & basepath & "temp\" & OUT_FILENAME & "\")
                                iresults = VerySplitMergePDF("burst " & OutputPath & " " & basepath & "temp\" & OUT_FILENAME & "\")

                                ' Kill original unsplit file
                                If debug = "Y" Then mydebuglog.Debug("  Removing: " & OutputPath)
                                Kill(OutputPath)

                                ' Copy and rename to base directory and remove
                                For j = 1 To NumFiles
                                    Rename(basepath & "temp\" & OUT_FILENAME & "\" & j.ToString("D4") & ".pdf", basepath & "temp\" & OUT_FILENAME & j.ToString("D4") & ".pdf")
                                    If debug = "Y" Then mydebuglog.Debug("  Moving to: " & OUT_FILENAME & j.ToString("D4") & ".pdf")
                                Next

                                ' Remove temp directory
                                RmDir(basepath & "temp\" & OUT_FILENAME & "\")

                                ' Add meta data and encrypt
                                PdfSubject = "A " & REPORT_NAME & " was prepared "
                                If debug = "Y" Then mydebuglog.Debug("   Subject: " & PdfSubject)
                                For j = 1 To NumFiles
                                    ' Compute the encryption path
                                    temp = basepath & "temp\" & OUT_FILENAME & j.ToString("D4") & "_e.pdf"
                                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Encoding to: " & temp)

                                    ' Add meta information to the pdf
                                    pdfid = VeryOpen(basepath & "temp\" & OUT_FILENAME & j.ToString("D4") & ".pdf")
                                    If pdfid <> 0 Then
                                        If debug = "Y" Then mydebuglog.Debug("   > adding meta data to pdf id: " & pdfid.ToString)
                                        PdfKeywords = LabelRows(j - 1, DmsKeyField)
                                        If debug = "Y" Then mydebuglog.Debug("   > keywords: " & PdfKeywords)
                                        iresults = VeryAddInfo(pdfid, REPORT_NAME, PdfSubject, Domain, PdfKeywords, Domain)

                                        ' Close the pdf
                                        If iresults <> 0 Then VeryClose(pdfid)
                                    End If

                                    ' Encrypt the pdf so that it can only be printed
                                    lresults = VeryEncryptPDF(basepath & "temp\" & OUT_FILENAME & j.ToString("D4") & ".pdf", temp, 128, 4, PdfPassword, "")
                                    If lresults < 0 Then
                                        errmsg = errmsg & "Error encrypting PDF " & temp & vbCrLf
                                        If debug = "Y" Then mydebuglog.Debug(errmsg)
                                        results = "Failure"
                                        GoTo CloseOut
                                    End If

                                    ' Kill the original pdf and rename the encrypted version to the original
                                    If debug = "Y" Then mydebuglog.Debug("   > renaming encrypted file to: " & basepath & "temp\" & OUT_FILENAME & j.ToString("D4") & ".pdf")
                                    Kill(basepath & "temp\" & OUT_FILENAME & j.ToString("D4") & ".pdf")
                                    Rename(temp, basepath & "temp\" & OUT_FILENAME & j.ToString("D4") & ".pdf")
                                Next
                            Else
                                ' Compute the encryption path
                                temp = Replace(OutputPath, ".pdf", "_e.pdf")
                                If debug = "Y" Then mydebuglog.Debug("  Encoding to: " & temp)

                                ' Add meta information to the pdf
                                pdfid = VeryOpen(OutputPath)
                                If pdfid <> 0 Then
                                    If debug = "Y" Then mydebuglog.Debug("  Adding meta data to pdf id: " & pdfid.ToString)
                                    PdfSubject = "A " & REPORT_NAME & " was prepared "
                                    If debug = "Y" Then mydebuglog.Debug("   > subject: " & PdfSubject)
                                    For j = 1 To NumRows
                                        PdfKeywords = PdfKeywords & LabelRows(j - 1, DmsKeyField) & Chr(10)
                                    Next
                                    ' Make single document output more descriptive
                                    If NumRows > 1 Then
                                        REPORT_NAME = REPORT_NAME & "s" ' Make name plural if needed
                                        LabelRows(0, DmsKeyField) = Trim(Replace(PdfKeywords, Chr(10), ", "))
                                        If Right(LabelRows(0, DmsKeyField), 1) = "," Then LabelRows(0, DmsKeyField) = Left(LabelRows(0, DmsKeyField), Len(LabelRows(0, DmsKeyField)) - 1)
                                        If Len(LabelRows(0, DmsKeyField)) > 300 Then LabelRows(0, DmsKeyField) = Left(LabelRows(0, DmsKeyField), 300) & "..."
                                    End If
                                    If debug = "Y" Then mydebuglog.Debug("   > keywords: " & PdfKeywords)
                                    iresults = VeryAddInfo(pdfid, REPORT_NAME, PdfSubject, Domain, PdfKeywords, Domain)

                                    ' Close the pdf
                                    If iresults <> 0 Then VeryClose(pdfid)
                                End If

                                ' Encrypt the pdf so that it can only be printed
                                lresults = VeryEncryptPDF(OutputPath, temp, 128, 4, PdfPassword, "")
                                If lresults < 0 Then
                                    errmsg = errmsg & "Error encrypting PDF " & temp & vbCrLf
                                    If debug = "Y" Then mydebuglog.Debug(errmsg)
                                    results = "Failure"
                                    GoTo CloseOut
                                End If

                                ' Kill the original pdf and rename the encrypted version to the original
                                If debug = "Y" Then mydebuglog.Debug("  Renaming encrypted file to: " & OutputPath)
                                Try
                                    Kill(OutputPath)
                                Catch ex As Exception
                                End Try
                                Rename(temp, OutputPath)
                                NumFiles = 1
                            End If

                        Case "tif"
                            ' Compute the output filename
                            temp = Replace(OutputPath, ".pdf", ".tif")
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Converting to: " & temp)

                            ' Split it up if MultiFlg is set to "Y"
                            If MultiFlg = "Y" Then
                                iresults = PDFToImageConverter(OutputPath, temp, "", "", 300, 300, 8, CompressionType.COMPRESSION_PACKBITS, 70, True, False, -1, -1)
                                OUT_FILENAME = Replace(OUT_FILENAME, ".pdf", "")   ' Change filename
                                NumFiles = NumRows
                            Else
                                iresults = PDFToImageConverter(OutputPath, temp, "", "", 300, 300, 8, CompressionType.COMPRESSION_PACKBITS, 70, True, True, -1, -1)
                                OUT_FILENAME = Replace(OUT_FILENAME, ".pdf", ".tif")   ' Change filename
                                NumFiles = 1
                            End If
                            If iresults < 0 Then
                                errmsg = errmsg & "Error converting PDF to TIF" & vbCrLf
                                If debug = "Y" Then mydebuglog.Debug(errmsg)
                                results = "Failure"
                                GoTo CloseOut
                            End If

                            ' Remove the results pdf file
                            Kill(OutputPath)
                            OutputPath = temp

                        Case "jpg"
                            ' Get width of images to create
                            PageWidth = GetPageWidth(OutputPath, 1)
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Image page width: " & PageWidth.ToString)

                            ' Perform conversion
                            temp = Replace(OutputPath, ".pdf", ".jpg")
                            If debug = "Y" Then mydebuglog.Debug("  Converting to: " & temp)

                            ' Crop if applicable
                            Dim ConImage As String
                            ConImage = ""
                            If PResX > 0 And PResY > 0 Then
                                PResX = PResX * 3
                                PResY = PResY * 3
                                iresults = PDFToImageConverterEx(OutputPath, temp, "", "", 0, PResX, PResY, 24, CompressionType.COMPRESSION_NONE, 100, True, False, -1, -1)
                                If iresults = 0 Then
                                    If MultiFlg <> "Y" Then
                                        If debug = "Y" Then mydebuglog.Debug("   .. cropping to: " & Str(CropRect.Width).Trim & " x " & Str(CropRect.Height).Trim)
                                        ConImage = Replace(temp, ".jpg", "0001.jpg")
                                        OrignalImage = Image.FromFile(ConImage, True)
                                        Dim CropImage = New Bitmap(CropRect.Width, CropRect.Height)
                                        Using grp = Graphics.FromImage(CropImage)
                                            grp.DrawImage(OrignalImage, New Rectangle(0, 0, CropRect.Width, CropRect.Height), CropRect, GraphicsUnit.Pixel)
                                            OrignalImage.Dispose()

                                            ' Resize cropped image
                                            Dim ReducedImage As New Bitmap(CInt(CropRect.Width / 3) + 3, CInt(CropRect.Height / 3) + 3)
                                            Dim ReducedImageDest As Graphics = Graphics.FromImage(ReducedImage)
                                            ReducedImageDest.DrawImage(CropImage, 0, 0, ReducedImage.Width + 1, ReducedImage.Height + 1)
                                            ReducedImage.Save(ConImage)
                                            ReducedImageDest.Dispose()
                                            ReducedImage.Dispose()

                                            'CropImage.Save(ConImage)
                                            CropImage.Dispose()
                                            CropImage = Nothing
                                        End Using
                                    Else
                                        For j = 1 To NumFiles
                                            ConImage = Replace(temp, ".jpg", "") & j.ToString("D4") & "." & OutFormat
                                            If debug = "Y" Then mydebuglog.Debug("   .. cropping " & ConImage & " to: " & Str(CropRect.Width).Trim & " x " & Str(CropRect.Height).Trim)
                                            OrignalImage = Image.FromFile(ConImage, True)
                                            Dim CropImage = New Bitmap(CropRect.Width, CropRect.Height)
                                            Using grp = Graphics.FromImage(CropImage)
                                                grp.DrawImage(OrignalImage, New Rectangle(0, 0, CropRect.Width, CropRect.Height), CropRect, GraphicsUnit.Pixel)
                                                OrignalImage.Dispose()

                                                ' Resize cropped image
                                                Dim ReducedImage As New Bitmap(CInt(CropRect.Width / 2), CInt(CropRect.Height / 2))
                                                Dim ReducedImageDest As Graphics = Graphics.FromImage(ReducedImage)
                                                ReducedImageDest.DrawImage(CropImage, 0, 0, ReducedImage.Width + 1, ReducedImage.Height + 1)
                                                ReducedImage.Save(ConImage)
                                                ReducedImageDest.Dispose()
                                                ReducedImage.Dispose()

                                                'CropImage.Save(ConImage)
                                                CropImage.Dispose()
                                                CropImage = Nothing
                                            End Using
                                        Next
                                    End If
                                End If
                            Else
                                iresults = PDFToImageConverterEx(OutputPath, temp, "", "", 1, PageWidth, 0, 24, CompressionType.COMPRESSION_JPEG, 100, True, False, -1, -1)
                                If iresults = 0 Then
                                    If MultiFlg <> "Y" Then
                                        ConImage = Replace(temp, ".jpg", "0001.jpg")
                                    Else
                                        For j = 1 To NumFiles
                                            ConImage = Replace(temp, ".jpg", "") & j.ToString("D4") & "." & OutFormat
                                        Next
                                    End If
                                End If
                            End If

                            ' Secure Image if card product
                            Dim UpdImage As System.Drawing.Image
                            Dim CmtFile, SecFile, StegCmd, StegFile As String

                            If TypeProd = "C" Then
                                If MultiFlg <> "Y" Then
                                    ' If single image, add steganography
                                    If LabelRows(0, DmsKeyField) <> "" Then
                                        ' Convert image to jpeg
                                        CmtFile = Replace(ConImage, ".jpg", "-c.jpg")
                                        If debug = "Y" Then mydebuglog.Debug("   .. converting: " & ConImage)
                                        UpdImage = Bitmap.FromFile(ConImage)
                                        UpdImage.Save(CmtFile, System.Drawing.Imaging.ImageFormat.Jpeg)
                                        UpdImage.Dispose()
                                        UpdImage = Nothing
                                        My.Computer.FileSystem.DeleteFile(ConImage)
                                        My.Computer.FileSystem.MoveFile(CmtFile, ConImage)

                                        ' Create security text
                                        SecFile = basepath & "temp\" & SrcId & ".txt"
                                        If debug = "Y" Then mydebuglog.Debug("   .. creating security text: " & SecFile)
                                        Using outfile As New StreamWriter(SecFile)
                                            outfile.Write(LabelRows(0, DmsKeyField))
                                            outfile.Close()
                                        End Using

                                        ' Add security file to the image via steganography
                                        StegFile = Replace(ConImage, ".jpg", "-s.jpg")
                                        If debug = "Y" Then mydebuglog.Debug("   .. creating steg file: " & StegFile)
                                        If System.IO.File.Exists(StegFile) Then My.Computer.FileSystem.DeleteFile(StegFile)
                                        Dim errorflg As String
                                        errorflg = ""
                                        If System.IO.File.Exists(ConImage) And System.IO.File.Exists(SecFile) Then
                                            StegCmd = basepath & "jpeg\steghide.exe embed -cf " & ConImage & " -ef " & SecFile & " -sf " & StegFile & " -p csi -f -q"
                                            If debug = "Y" Then mydebuglog.Debug("     > add steg cmd: " & StegCmd)
                                            errorflg = ShellandWait(StegCmd, mydebuglog, debug)
                                            If errorflg <> "Y" Then
                                                StegCmd = basepath & "jpeg\steghide.exe extract -sf " & ConImage & " -xf " & basepath & "temp\test.txt -v -p csi -f -q"
                                                If debug = "Y" Then mydebuglog.Debug("     > extract cmd: " & StegCmd)
                                            Else
                                                If debug = "Y" Then mydebuglog.Debug("   .. steg failed")
                                                If debug = "Y" Then mydebuglog.Debug("   .. renaming: " & ConImage & " to " & StegFile)
                                                My.Computer.FileSystem.MoveFile(ConImage, StegFile)
                                            End If
                                        End If
                                        If debug = "Y" Then mydebuglog.Debug("   .. deleting: " & SecFile)
                                        My.Computer.FileSystem.DeleteFile(SecFile)
                                        If debug = "Y" Then mydebuglog.Debug("   .. ConImage: " & ConImage)
                                        My.Computer.FileSystem.DeleteFile(ConImage)

                                        ' Add comments to steg file
                                        CmtFile = Replace(ConImage, ".jpg", "-c.jpg")
                                        Try
                                            If System.IO.File.Exists(CmtFile) Then My.Computer.FileSystem.DeleteFile(CmtFile)
                                        Catch ex As Exception
                                        End Try
                                        If debug = "Y" Then mydebuglog.Debug("   .. adding comment: """ & LabelRows(0, DmsKeyField) & """ to " & StegFile)
                                        UpdImage = Bitmap.FromFile(StegFile)
                                        SetImageProperty(UpdImage, TagNames.ImageDescription, StringToBytes(LabelRows(0, DmsKeyField)), ExifDataTypes.AsciiString)
                                        SetImageProperty(UpdImage, TagNames.Copyright, StringToBytes("Copyright " & Year(Now) & " Health Communications, Inc"), ExifDataTypes.AsciiString)
                                        SetImageProperty(UpdImage, TagNames.ImageTitle, StringToBytes(REPORT_NAME), ExifDataTypes.AsciiString)
                                        UpdImage.Save(CmtFile, System.Drawing.Imaging.ImageFormat.Jpeg)
                                        UpdImage.Dispose()
                                        UpdImage = Nothing

                                        ' Clean up files
                                        Try
                                            My.Computer.FileSystem.DeleteFile(Replace(ConImage, "0001.jpg", "0002.jpg"))
                                            My.Computer.FileSystem.DeleteFile(Replace(ConImage, "0001.jpg", "0003.jpg"))
                                            My.Computer.FileSystem.DeleteFile(Replace(ConImage, "0001.jpg", "0004.jpg"))
                                        Catch ex As Exception
                                        End Try
                                        If debug = "Y" Then mydebuglog.Debug("   .. renaming: " & CmtFile & " to " & ConImage)
                                        My.Computer.FileSystem.MoveFile(CmtFile, ConImage)
                                        If debug = "Y" Then mydebuglog.Debug("   .. deleting: " & StegFile)
                                        My.Computer.FileSystem.DeleteFile(StegFile)

                                    End If
                                Else
                                    ' If multiple images, add EXIF comments only
                                    For j = 1 To NumFiles
                                        ConImage = Replace(temp, ".jpg", "") & j.ToString("D4") & "." & OutFormat
                                        CmtFile = Replace(ConImage, ".jpg", "-c.jpg")
                                        If debug = "Y" Then mydebuglog.Debug("   .. adding comment: " & LabelRows(j - 1, DmsKeyField) & " to " & CmtFile)
                                        If LabelRows(j - 1, DmsKeyField) <> "" Then
                                            UpdImage = Bitmap.FromFile(ConImage)
                                            SetImageProperty(UpdImage, TagNames.ImageDescription, StringToBytes(LabelRows(j - 1, DmsKeyField)), ExifDataTypes.AsciiString)
                                            SetImageProperty(UpdImage, TagNames.Copyright, StringToBytes("Copyright " & Year(Now) & " Health Communications, Inc"), ExifDataTypes.AsciiString)
                                            SetImageProperty(UpdImage, TagNames.ImageTitle, StringToBytes(REPORT_NAME), ExifDataTypes.AsciiString)
                                            UpdImage.Save(CmtFile)
                                            UpdImage.Dispose()
                                            UpdImage = Nothing
                                            System.IO.File.Delete(ConImage)
                                            If debug = "Y" Then mydebuglog.Debug("   .. renaming: " & CmtFile & " to " & ConImage)
                                            Rename(CmtFile, ConImage)
                                        End If
                                    Next
                                End If
                            Else
                            End If

                            ' Automatically multi-flag .. set variable
                            MultiFlg = "Y"
                            If iresults < 0 Then
                                errmsg = errmsg & "Error converting PDF to JPG" & vbCrLf
                                If debug = "Y" Then mydebuglog.Debug(errmsg)
                                results = "Failure"
                                GoTo CloseOut
                            End If

                            ' Strip the format from the filename to get the base filename
                            OUT_FILENAME = Replace(OUT_FILENAME, ".pdf", "")                ' Set to base filename                             
                            NumFiles = NumRows

                            ' Remove the results pdf file
                            Kill(OutputPath)
                            OutputPath = temp

                        Case "gif"
                            ' Get width of images to create
                            PageWidth = GetPageWidth(OutputPath, 1)
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Image width: " & PageWidth.ToString)

                            ' Automatically set MultiFlg and split
                            MultiFlg = "Y"
                            temp = Replace(OutputPath, ".pdf", ".gif")
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Converting to: " & temp)
                            iresults = PDFToImageConverterEx(OutputPath, temp, "", "", 1, PageWidth, 0, 24, CompressionType.COMPRESSION_PACKBITS, 100, False, False, -1, -1)
                            If iresults < 0 Then
                                errmsg = errmsg & "Error converting PDF to GIF" & vbCrLf
                                If debug = "Y" Then mydebuglog.Debug(errmsg)
                                results = "Failure"
                                GoTo CloseOut
                            End If

                            ' Strip the format from the filename to get the base filename
                            OUT_FILENAME = Replace(OUT_FILENAME, ".pdf", "")                ' Set to base filename                             
                            NumFiles = NumRows

                            ' Remove the results pdf file
                            Kill(OutputPath)
                            OutputPath = temp
                    End Select
                Catch ex As Exception
                    errmsg = errmsg & "Unable to convert temp file to another format: " & ex.ToString
                    results = "Failure"
                    GoTo CloseOut
                End Try

                ' -----
                ' COMPUTE DMS DOCUMENT INFORMATION
                '  Store to the array dmsdocs the temporary ids and filenames of the files created
                '   dmsdocs(#, 1) = An id which represents the job
                '   dmsdocs(#, 2) = The filename of the document minus the temporary directory
                '   dmsdocs(#, 3) = The computed description of the document
                '   dmsdocs(#, 4) = The computed category
                '   dmsdocs(#, 5) = The contact association id
                '   dmsdocs(#, 6) = The jurisdiction association id
                '   dmsdocs(#, 7) = The organization association id
                '   dmsdocs(#, 8) = The created document id when there is an update
                '   dmsdocs(#, 9) = Set to "Y" if an existing document
                '   dmsdocs(#,10) = The trainer workshop id
                '   dmsdocs(#,11) = The participant session id
                '   dmsdocs(#,12) = The scheduled participant session id
                '   dmsdocs(#,13) = The participation id
                '   dmsdocs(#,14) = The workshop registration id
                '   dmsdocs(#,15) = The session registration id
                '   dmsdocs(#,16) = The exam id
                '   dmsdocs(#,17) = The trainer certification id
                '   dmsdocs(#,18) = The output product, determined by OutputDest
                '   dmsdocs(#,19) = The computed contact key code
                '   dmsdocs(#,20) = The computed keyword
                '   dmsdocs(#,21) = The created document version id when there is an update
                '   dmsdocs(#,22) = The trainer id
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Preparing DMS Information at " & Now.ToString)
                ReDim dmsdocs(NumFiles, 22)
                If MultiFlg = "Y" Then
                    For j = 1 To NumFiles
                        OUT_FILENAME = Replace(OUT_FILENAME, ".pdf", "")
                        OUT_FILENAME = Replace(OUT_FILENAME, ".tif", "")
                        OUT_FILENAME = Replace(OUT_FILENAME, ".jpg", "")
                        OUT_FILENAME = Replace(OUT_FILENAME, ".gif", "")
                        dmsdocs(j, 1) = OUT_FILENAME
                        dmsdocs(j, 2) = OUT_FILENAME & j.ToString("D4") & "." & OutFormat
                        dmsdocs(j, 3) = LabelRows(j - 1, DmsKeyField)
                        dmsdocs(j, 4) = GetCategory(SkillLevel, TypeProd)
                        mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & " DmsConId: " & DmsConId & " ConId:" & ConId & " DmsJurisId: " & DmsJurisId & " JurisId: " & JurisId & " DmsOrgId: " & DmsOrgId & " OrgId: " & OrgId)
                        If DmsConId <> 0 Then dmsdocs(j, 5) = LabelRows(j - 1, DmsConId)
                        If ConId <> "" And dmsdocs(j, 5) = "" Then dmsdocs(j, 5) = ConId
                        mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & " dmsdocs(j, 5): " & dmsdocs(j, 5))
                        If DmsJurisId <> 0 Then dmsdocs(j, 6) = LabelRows(j - 1, DmsJurisId)
                        If JurisId <> "" And dmsdocs(j, 6) = "" Then dmsdocs(j, 6) = JurisId
                        If DmsOrgId <> 0 Then dmsdocs(j, 7) = LabelRows(j - 1, DmsOrgId)
                        If OrgId <> "" And dmsdocs(j, 7) = "" Then dmsdocs(j, 7) = OrgId
                        dmsdocs(j, 8) = ""          ' Document ID will be stored here
                        dmsdocs(j, 9) = "N"
                        If DmsWshopId <> 0 Then dmsdocs(j, 10) = LabelRows(j - 1, DmsWshopId)
                        If DmsSessId <> 0 Then dmsdocs(j, 11) = LabelRows(j - 1, DmsSessId)
                        If DmsOffrId <> 0 Then dmsdocs(j, 12) = LabelRows(j - 1, DmsOffrId)
                        If DmsPartId <> 0 Then dmsdocs(j, 13) = LabelRows(j - 1, DmsPartId)
                        If DmsWregId <> 0 Then dmsdocs(j, 14) = LabelRows(j - 1, DmsWregId)
                        If DmsSregId <> 0 Then dmsdocs(j, 15) = LabelRows(j - 1, DmsSregId)
                        If DmsExamId <> 0 Then dmsdocs(j, 16) = LabelRows(j - 1, DmsExamId)
                        If DmsCertId <> 0 Then dmsdocs(j, 17) = LabelRows(j - 1, DmsCertId)
                        dmsdocs(j, 18) = ""
                        dmsdocs(j, 19) = GenerateUserKey(dmsdocs(j, 5))
                        dmsdocs(j, 20) = GetKeyword(SkillLevel, TypeProd)
                        dmsdocs(j, 21) = ""         ' Document Version ID will be stored here
                        If DmsTrainerId <> 0 Then dmsdocs(j, 22) = LabelRows(j - 1, DmsTrainerId)
                    Next
                Else
                    dmsdocs(1, 1) = Replace(OUT_FILENAME, "." & OutFormat, "")
                    If InStr(OUT_FILENAME, "." & OutFormat) = 0 Then
                        OUT_FILENAME = OUT_FILENAME & "." & OutFormat
                    End If
                    dmsdocs(1, 2) = OUT_FILENAME
                    dmsdocs(1, 3) = LabelRows(0, DmsKeyField)
                    If TypeProd = "P" And OrgId <> "" And ACCT_NAME <> "" Then
                        If ACCT_LOC <> "" Then ACCT_NAME = ACCT_NAME & ", " & ACCT_LOC
                        dmsdocs(1, 3) = ACCT_NAME & ", " & dmsdocs(1, 3)
                    End If
                    dmsdocs(1, 4) = GetCategory(SkillLevel, TypeProd)
                    If DmsConId <> 0 Then dmsdocs(1, 5) = LabelRows(0, DmsConId)
                    If ConId <> "" And dmsdocs(1, 5) = "" Then dmsdocs(1, 5) = ConId
                    If DmsJurisId <> 0 Then dmsdocs(1, 6) = LabelRows(0, DmsJurisId)
                    If JurisId <> "" And dmsdocs(1, 6) = "" Then dmsdocs(1, 6) = JurisId
                    If DmsOrgId <> 0 Then dmsdocs(1, 7) = LabelRows(0, DmsOrgId)
                    If OrgId <> "" And dmsdocs(1, 7) = "" Then dmsdocs(1, 7) = OrgId
                    dmsdocs(1, 8) = ""
                    dmsdocs(1, 9) = "N"
                    If DmsWshopId <> 0 Then dmsdocs(1, 10) = LabelRows(0, DmsWshopId)
                    If DmsSessId <> 0 Then dmsdocs(1, 11) = LabelRows(0, DmsSessId)
                    If DmsOffrId <> 0 Then dmsdocs(1, 12) = LabelRows(0, DmsOffrId)
                    If DmsPartId <> 0 Then dmsdocs(1, 13) = LabelRows(0, DmsPartId)
                    If DmsWregId <> 0 Then dmsdocs(1, 14) = LabelRows(0, DmsWregId)
                    If DmsSregId <> 0 Then dmsdocs(1, 15) = LabelRows(0, DmsSregId)
                    If DmsExamId <> 0 Then dmsdocs(1, 16) = LabelRows(0, DmsExamId)
                    If DmsCertId <> 0 Then dmsdocs(1, 17) = LabelRows(0, DmsCertId)
                    dmsdocs(1, 18) = ""
                    dmsdocs(1, 19) = GenerateUserKey(dmsdocs(1, 5))
                    dmsdocs(1, 20) = GetKeyword(SkillLevel, TypeProd)
                    dmsdocs(1, 21) = ""
                    If DmsTrainerId <> 0 Then dmsdocs(1, 22) = LabelRows(0, DmsTrainerId)
                End If
                LabelRows = Nothing

                If debug = "Y" Then
                    For j = 1 To NumFiles
                        Try
                            mydebuglog.Debug("  > Storing ID: " & dmsdocs(j, 1) & " for file: " & dmsdocs(j, 2) & vbCrLf & _
                            "     Description: " & dmsdocs(j, 3) & vbCrLf & "     Category Id: " & dmsdocs(1, 4) & "     Contact Id: " & dmsdocs(j, 5) & _
                            "     Jurisdiction Id: " & dmsdocs(j, 6) & "     Organization Id: " & dmsdocs(j, 7) & vbCrLf & _
                            "     Workshop Id: " & dmsdocs(j, 10) & "     Session Id: " & dmsdocs(j, 11) & _
                            "     Scheduled Session Id: " & dmsdocs(j, 12) & "     Participation Id: " & dmsdocs(j, 13) & vbCrLf & _
                            "     Workshop Reg Id: " & dmsdocs(j, 14) & "     Session Reg Id: " & dmsdocs(j, 15) & _
                            "     Exam Id: " & dmsdocs(j, 16) & "     Certification Id: " & dmsdocs(j, 17) & _
                            "     User Key: " & dmsdocs(j, 19) & "     Keyword: " & dmsdocs(j, 20) & "     Trainer Id: " & dmsdocs(j, 22))
                        Catch ex As Exception
                        End Try
                    Next
                End If

                ' -----
                ' STORE THE FILE(S) GENERATED TO THE DMS
                '  Iterate through each document and determine what to store
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Storing Documents to the DMS at " & Now.ToString)
                DataTypeId = TranslateDataType(OutFormat)
                For j = 1 To NumFiles

                    ' Check to see if the document exists in the DMS, if so, replace it
                    If ExistDocId <> "" And NumFiles = 1 Then
                        ' Existing document specified 
                        SqlS = "SELECT d.row_id, d.last_version_id " & _
                            "FROM DMS.dbo.Documents d " & _
                            "WHERE d.row_id=" & ExistDocId
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locating specified duplicate product: " & vbCrLf & SqlS)
                        Try
                            dcmd.CommandText = SqlS
                            ddr = dcmd.ExecuteReader()
                            If Not ddr Is Nothing Then
                                Try
                                    While ddr.Read()
                                        If ddr.HasRows Then
                                            dmsdocs(j, 8) = Trim(CheckDBNull(ddr(0), enumObjectType.StrType)).ToString
                                            dmsdocs(j, 21) = Trim(CheckDBNull(ddr(1), enumObjectType.StrType)).ToString
                                            If debug = "Y" Then
                                                mydebuglog.Debug(vbCrLf & "   > Found document id " & dmsdocs(j, 8))
                                                mydebuglog.Debug(vbCrLf & "   > Found document versions id " & dmsdocs(j, 21))
                                            End If
                                        End If
                                    End While
                                Catch ex As Exception
                                End Try
                            Else
                                errmsg = errmsg & "DMS database access error" & vbCrLf
                                If debug = "Y" Then mydebuglog.Debug(errmsg)
                                results = "Failure"
                                GoTo CloseOut
                            End If
                            Try
                                ddr.Close()
                                ddr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                            errmsg = errmsg & "DMS database access error:  " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try
                    End If

                    ' Check to see if there is a previously generated certification product
                    SqlS = "SELECT TOP 1 DOC_ID, ROW_ID " & _
                    "FROM siebeldb.dbo.CX_CERT_PROD_RESULTS " & _
                    "WHERE CERT_CRSE_ID='" & CrseId & "' AND CON_ID='" & dmsdocs(j, 5) & "' AND PROD_ID='" & ProdId & "' AND DESTINATION='" & OutputDest & "' "
                    Select Case TypeProd
                        Case "R"    ' Results Certification
                            SqlS = SqlS & "AND REG_ID='" & dmsdocs(j, 15) & "' "
                        Case "C"    ' Certification Card
                            Select Case SkillLevel
                                Case "Participant"
                                    SqlS = SqlS & "AND SESS_PART_ID='" & dmsdocs(j, 13) & "' "
                                Case "Trainer"
                                    SqlS = SqlS & "AND CURRCLM_PER_ID='" & dmsdocs(j, 17) & "' "
                                Case "Master Trainer"
                                    SqlS = SqlS & "AND CURRCLM_PER_ID='" & dmsdocs(j, 17) & "' "
                            End Select
                        Case "F"    ' Course Regulatory Form
                            Select Case SkillLevel
                                Case "Participant"
                                    SqlS = SqlS & "AND SESS_PART_ID='" & dmsdocs(j, 13) & "' "
                                Case "Trainer"
                                    SqlS = SqlS & "AND CURRCLM_PER_ID='" & dmsdocs(j, 17) & "' "
                                Case "Master Trainer"
                                    SqlS = SqlS & "AND CURRCLM_PER_ID='" & dmsdocs(j, 17) & "' "
                            End Select
                        Case "W"    ' Wall Certificate
                            Select Case SkillLevel
                                Case "Participant"
                                    SqlS = SqlS & "AND SESS_PART_ID='" & dmsdocs(j, 13) & "' "
                                Case "Trainer"
                                    SqlS = SqlS & "AND CURRCLM_PER_ID='" & dmsdocs(j, 17) & "' "
                                Case "Master Trainer"
                                    SqlS = SqlS & "AND CURRCLM_PER_ID='" & dmsdocs(j, 17) & "' "
                            End Select
                        Case "P"    ' Passport
                            If dmsdocs(j, 15) = "" Then
                                SqlS = SqlS & "AND (REG_ID='" & dmsdocs(j, 15) & "' OR REG_ID IS NULL) "
                            Else
                                SqlS = SqlS & "AND REG_ID='" & dmsdocs(j, 15) & "' "
                            End If
                        Case "V"    ' Voucher
                            SqlS = SqlS & "AND CRSE_TSTRUN_ID='" & dmsdocs(j, 16) & " '"
                        Case "T"    ' Trainer Exam Completion Certificate
                            SqlS = SqlS & "AND CURRCLM_PER_ID='" & dmsdocs(j, 17) & " '"
                    End Select
                    If IdentStart <> "0" And IdentStart <> "" Then
                        SqlS = SqlS & "AND IDENT_START=" & IdentStart & " "
                    End If
                    If IdentEnd <> "0" And IdentEnd <> "" Then
                        SqlS = SqlS & "AND IDENT_END=" & IdentEnd & " "
                    End If
                    If ExistDocId <> "" Then
                        SqlS = SqlS & "AND DOC_ID='" & ExistDocId & "' "
                    End If
                    SqlS = SqlS & " ORDER BY CREATED DESC"

                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Checking for duplicate product: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            Try
                                While dr.Read()
                                    ' Save document id to update
                                    dmsdocs(j, 8) = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                    If dmsdocs(j, 8) <> "" Then dmsdocs(j, 9) = "Y"
                                    If debug = "Y" Then mydebuglog.Debug("    > Found document id " & dmsdocs(j, 8) & " in record " & Trim(CheckDBNull(dr(1), enumObjectType.StrType)))
                                End While
                            Catch ex As Exception
                                errmsg = errmsg & "Error searching for duplicate product: " & ex.Message
                            End Try
                        Else
                            errmsg = errmsg & "Error searching for duplicate product."
                            results = "Failure"
                            GoTo CloseOut
                        End If
                        Try
                            dr.Close()
                            dr = Nothing
                        Catch ex As Exception
                        End Try

                        ' Locate last_version_id
                        If dmsdocs(j, 8) <> "" And dmsdocs(j, 21) = "" Then
                            SqlS = "SELECT d.row_id, d.last_version_id " & _
                                "FROM DMS.dbo.Documents d " & _
                                "WHERE d.row_id=" & dmsdocs(j, 8)
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locating specified duplicate product Document_Versions id: " & vbCrLf & SqlS)
                            Try
                                dcmd.CommandText = SqlS
                                ddr = dcmd.ExecuteReader()
                                If Not ddr Is Nothing Then
                                    Try
                                        While ddr.Read()
                                            If ddr.HasRows Then
                                                dmsdocs(j, 21) = Trim(CheckDBNull(ddr(1), enumObjectType.StrType)).ToString
                                                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   > Found document id version " & dmsdocs(j, 21))
                                            End If
                                        End While
                                    Catch ex As Exception
                                    End Try
                                Else
                                    errmsg = errmsg & "DMS database access error" & vbCrLf
                                    If debug = "Y" Then mydebuglog.Debug(errmsg)
                                    results = "Failure"
                                    GoTo CloseOut
                                End If
                                Try
                                    ddr.Close()
                                    ddr = Nothing
                                Catch ex As Exception
                                End Try
                            Catch ex As Exception
                                errmsg = errmsg & "DMS database access error:  " & ex.Message
                                results = "Failure"
                                GoTo CloseOut
                            End Try

                        End If

                    Catch ex As Exception
                        errmsg = errmsg & "Error searching for duplicate product: " & ex.Message
                        results = "Failure"
                        GoTo CloseOut
                    End Try

                    ' Get temp file and attach to MyData object
                    temp = basepath & "temp\" & dmsdocs(j, 2)
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Opening temp file to save: " & temp)
                    Dim mstream As New System.IO.FileStream(temp, FileMode.OpenOrCreate, FileAccess.Read)
                    lFileLength = mstream.Length
                    If debug = "Y" Then mydebuglog.Debug("    > Input stream length " & lFileLength.ToString)
                    Dim MyData(lFileLength) As Byte
                    Try
                        mstream.Read(MyData, 0, lFileLength)
                        mstream.Close()
                        mstream = Nothing
                    Catch ex As Exception
                        errmsg = errmsg & "Error reading temp file into filestream for saving: " & ex.Message
                        results = "Failure"
                        GoTo CloseOut
                    End Try

                    ' Fix fields if lengthy
                    If Len(REPORT_NAME) > 255 Then
                        REPORT_NAME = Left(REPORT_NAME, 250) & "..."
                        If debug = "Y" Then mydebuglog.Debug("   >> Truncating Report Name to " & REPORT_NAME)
                    End If
                    If Len(dmsdocs(j, 1)) > 15 Then
                        dmsdocs(j, 1) = Left(dmsdocs(j, 1), 15)
                        If debug = "Y" Then mydebuglog.Debug("   >> Truncating Ext Id to " & dmsdocs(j, 1))
                    End If
                    If Len(dmsdocs(j, 2)) > 254 Then
                        dmsdocs(j, 2) = Left(dmsdocs(j, 2), 250) & "..."
                        If debug = "Y" Then mydebuglog.Debug("   >> Truncating DFilename to " & dmsdocs(j, 2))
                    End If
                    If Len(dmsdocs(j, 3)) > 1999 Then
                        dmsdocs(j, 3) = Left(dmsdocs(j, 3), 1995) & "..."
                        If debug = "Y" Then mydebuglog.Debug("   >> Truncating Description to " & dmsdocs(j, 3))
                    End If

                    ' Update or create document record
                    Dim TryAgain As Boolean
                    TryAgain = False
                    If dmsdocs(j, 8) <> "" And dmsdocs(j, 21) <> "" Then
                        ' Update the existing Documents record
                        Try
                            SqlS = "UPDATE DMS.dbo.Documents " & _
                                "SET deleted=null, data_type_id=@data_type_id, ext_id=@ext_id, dfilename=@dfilename, " & _
                                "last_upd=@last_upd, last_upd_by=@last_upd_by, description=@description " & _
                                "WHERE row_id=" & dmsdocs(j, 8)
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Updating existing Documents: " & vbCrLf & SqlS & vbCrLf)
                            addDoc = New SqlCommand(SqlS, dcon)
                            addDoc.CommandTimeout = 30
                            addDoc.Parameters.Add("@data_type_id", SqlDbType.Int).Value = DataTypeId
                            addDoc.Parameters.Add("@ext_id", SqlDbType.VarChar, 15).Value = dmsdocs(j, 1)
                            addDoc.Parameters.Add("@dfilename", SqlDbType.VarChar, 255).Value = dmsdocs(j, 2)
                            addDoc.Parameters.Add("@last_upd", SqlDbType.DateTime).Value = Now
                            addDoc.Parameters.Add("@last_upd_by", SqlDbType.Int).Value = 1
                            addDoc.Parameters.Add("@description", SqlDbType.VarChar, 2000).Value = dmsdocs(j, 3)
                            addDoc.ExecuteNonQuery()
                            Try
                                addDoc.ExecuteNonQuery()
                            Catch ex3 As Exception
                                TryAgain = True
                            End Try
                            addDoc.Dispose()
                            addDoc = Nothing
                        Catch ex As Exception
                            errmsg = errmsg & "Error updating product in DMS: " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try

DocSave:
                        ' Update the existing Document_Versions record
                        Try
                            SqlS = "UPDATE DMS.dbo.Document_Versions " & _
                                "SET deleted=NULL, dimage=@dimage, dsize=@dsize, last_upd=@last_upd, last_upd_by=@last_upd_by,  " & _
                                "backed_up=getdate(), version=version+1 " & _
                                "WHERE row_id=" & dmsdocs(j, 21)
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Updating existing Document_Versions: " & vbCrLf & SqlS & vbCrLf)
                            addDoc = New SqlCommand(SqlS, dcon)
                            addDoc.CommandTimeout = 30
                            addDoc.Parameters.Add("@dimage", SqlDbType.Image, MyData.Length).Value = MyData
                            addDoc.Parameters.Add("@dsize", SqlDbType.Int).Value = lFileLength
                            addDoc.Parameters.Add("@last_upd", SqlDbType.DateTime).Value = Now
                            addDoc.Parameters.Add("@last_upd_by", SqlDbType.Int).Value = 1
                            addDoc.ExecuteNonQuery()
                            Try
                                addDoc.ExecuteNonQuery()
                            Catch ex3 As Exception
                                TryAgain = True
                            End Try
                            addDoc.Dispose()
                            addDoc = Nothing
                        Catch ex As Exception
                            errmsg = errmsg & "Error updating product in DMS: " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try

                        If TryAgain Then
                            ' Pause and try again up to 3 times
                            System.Threading.Thread.Sleep(1000)
                            SaveTries = SaveTries + 1
                            If SaveTries < 3 Then GoTo DocSave
                        End If

                    Else
                        ' Generate the Documents record
                        SqlS = "INSERT INTO DMS.dbo.Documents " & _
                            "(ext_id, data_type_id, dfilename, name, created_by, last_upd_by, description) " & _
                            "VALUES ('" & SqlString(dmsdocs(j, 1)) & "', " & DataTypeId.ToString & ", '" & SqlString(dmsdocs(j, 2)) & "', '" & _
                            SqlString(REPORT_NAME) & "', 1, 1, '" & SqlString(dmsdocs(j, 3)) & "')"
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Inserting new Document: " & vbCrLf & SqlS & vbCrLf)
                        dcmd.CommandText = SqlS
                        Try
                            returnv = dcmd.ExecuteNonQuery()
                        Catch ex As Exception
                            errmsg = errmsg & "Error saving product Document in DMS: " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try

                        ' Locate the document just created
                        Try
                            SqlS = "SELECT d.row_id " & _
                                "FROM DMS.dbo.Documents d " & _
                                "WHERE d.ext_id='" & dmsdocs(j, 1) & "' and dfilename='" & dmsdocs(j, 2) & "'"
                            If debug = "Y" Then mydebuglog.Debug("  Locating Document id for file " & dmsdocs(j, 2) & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            ddr = dcmd.ExecuteReader()
                            If ddr Is Nothing Then
                                errmsg = errmsg & "DMS database access error" & vbCrLf
                                If debug = "Y" Then mydebuglog.Debug(errmsg)
                                results = "Failure"
                                GoTo CloseOut
                            End If
                            While ddr.Read()
                                If ddr.HasRows Then
                                    ' Save document id to update
                                    dmsdocs(j, 8) = Trim(CheckDBNull(ddr(0), enumObjectType.StrType)).ToString
                                End If
                            End While
                            Try
                                ddr.Close()
                                ddr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                            errmsg = errmsg & "Error locating new Document in DMS: " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try
                        If debug = "Y" Then mydebuglog.Debug("   > Document Id found " & dmsdocs(j, 8))
                        If dmsdocs(j, 8) = "" Then
                            errmsg = errmsg & "Error locating Document in DMS"
                            results = "Failure"
                            GoTo CloseOut
                        End If

DocSave2:
                        ' Generate the Document_Versions record
                        Try
                            SqlS = "INSERT INTO DMS.dbo.Document_Versions " & _
                                "(doc_id, dimage, dsize, created_by, last_upd_by, backed_up) " & _
                                "Values(@docid, @dimage, @dsize, @created_by, @last_upd_by, getdate())"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Inserting new Document_Versions: " & vbCrLf & SqlS & vbCrLf)
                            addDoc = New SqlCommand(SqlS, dcon)
                            addDoc.CommandTimeout = 30
                            addDoc.Parameters.Add("@docid", SqlDbType.Int).Value = dmsdocs(j, 8)
                            addDoc.Parameters.Add("@dimage", SqlDbType.Image, MyData.Length).Value = MyData
                            addDoc.Parameters.Add("@dsize", SqlDbType.Int).Value = lFileLength
                            addDoc.Parameters.Add("@created_by", SqlDbType.Int).Value = 1
                            addDoc.Parameters.Add("@last_upd_by", SqlDbType.Int).Value = 1
                            Try
                                addDoc.ExecuteNonQuery()
                            Catch ex3 As Exception
                                TryAgain = True
                            End Try
                            addDoc.Dispose()
                            addDoc = Nothing
                        Catch ex As Exception
                            errmsg = errmsg & "Error saving Document_Versions in DMS: " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try

                        ' If failure, try again 3 times
                        If TryAgain Then
                            ' Pause and try again up to 3 times
                            System.Threading.Thread.Sleep(1000)
                            SaveTries = SaveTries + 1
                            If SaveTries < 3 Then GoTo DocSave2
                        End If

                        ' Locate the Document_Versions record just created
                        Try
                            SqlS = "SELECT d.row_id " & _
                                "FROM DMS.dbo.Document_Versions d " & _
                                "WHERE d.doc_id=" & dmsdocs(j, 8)
                            If debug = "Y" Then mydebuglog.Debug("  Locating Document_Versions record for Document " & dmsdocs(j, 8) & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            ddr = dcmd.ExecuteReader()
                            If ddr Is Nothing Then
                                errmsg = errmsg & "DMS database access error" & vbCrLf
                                If debug = "Y" Then mydebuglog.Debug(errmsg)
                                results = "Failure"
                                GoTo CloseOut
                            End If
                            While ddr.Read()
                                If ddr.HasRows Then
                                    ' Save Document_Versions id 
                                    dmsdocs(j, 21) = Trim(CheckDBNull(ddr(0), enumObjectType.StrType)).ToString
                                End If
                            End While
                            Try
                                ddr.Close()
                                ddr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                            errmsg = errmsg & "Error locating Document_Versions in DMS: " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try
                        If debug = "Y" Then mydebuglog.Debug("   > Document_Versions Id found " & dmsdocs(j, 21))
                        If dmsdocs(j, 21) = "" Then
                            errmsg = errmsg & "Error locating Document_Versions in DMS for " & dmsdocs(j, 8)
                            results = "Failure"
                            GoTo CloseOut
                        End If

                        ' Update the Documents record to link to the Document_Versions record
                        SqlS = "UPDATE DMS.dbo.Documents " & _
                            "SET last_version_id=" & dmsdocs(j, 21) & " WHERE row_id=" & dmsdocs(j, 8)
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Updating Document record: " & vbCrLf & SqlS & vbCrLf)
                        dcmd.CommandText = SqlS
                        Try
                            returnv = dcmd.ExecuteNonQuery()
                        Catch ex As Exception
                            errmsg = errmsg & "Error saving product Document in DMS: " & ex.Message
                            results = "Failure"
                            GoTo CloseOut
                        End Try
                    End If

                    ' Kill temp file if the file was successfully stored
                    If dmsdocs(j, 8) <> "" Then
                        If debug <> "Y" Then
                            Kill(basepath & "temp\" & dmsdocs(j, 2))
                        End If
                    End If

                    ' Error handling if the document cannot be stored - do not update queue so we try it again
                    If dmsdocs(j, 8) = "" Then
                        If debug = "Y" Then
                            mydebuglog.Debug("  Unable to locate the document. Values attempted to store/update: ")
                            mydebuglog.Debug("    dsize : " & lFileLength.ToString)
                            mydebuglog.Debug("    data_type_id : " & DataTypeId.ToString)
                            mydebuglog.Debug("    ext_id : " & dmsdocs(j, 1))
                            mydebuglog.Debug("    dfilename : " & dmsdocs(j, 2))
                            mydebuglog.Debug("    name : " & REPORT_NAME)
                            mydebuglog.Debug("    description : " & dmsdocs(j, 3))
                        End If
                        GoTo CloseOut
                    End If
                Next

                ' -----
                ' CREATE DOCUMENT CATEGORY RECORD
                '    Using dmsdocs(#, 4) = The computed category
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Creating Document Category records at " & Now.ToString)
                For j = 1 To NumFiles
                    If dmsdocs(j, 8) <> "" Then
                        SqlS = "INSERT INTO DMS.dbo.Document_Categories(doc_id, cat_id, created_by, last_upd_by, pr_flag) " & _
                            "VALUES (" & dmsdocs(j, 8) & ", " & dmsdocs(j, 4) & ", 1, 1, 'Y')"
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Categories record: " & vbCrLf & SqlS)
                        dcmd.CommandText = SqlS
                        Try
                            returnv = dcmd.ExecuteNonQuery()
                        Catch ex As Exception
                        End Try

                        ' Add to "image" category if the output destination is "web"
                        If OutputDest = "web" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Categories(doc_id, cat_id, created_by, last_upd_by, pr_flag) " & _
                                "VALUES (" & dmsdocs(j, 8) & ", 118, 1, 1, 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Categories record for 'Images': " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                Next

                ' -----
                ' CREATE DOCUMENT KEYWORD RECORDS
                '    Using dmsdocs(#, 20) = The computed keyword
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Creating Document Keyword records at " & Now.ToString)
                For j = 1 To NumFiles
                    If dmsdocs(j, 20) <> "" Then
                        ' Assign Value Based on Keyword
                        KeyVal = ""
                        Select Case dmsdocs(j, 20)
                            Case "3"    ' Trainer
                                KeyVal = dmsdocs(j, 17)
                            Case "7"    ' Participant
                                KeyVal = dmsdocs(j, 13)
                        End Select

                        ' Generate SQL for insert
                        SqlS = "INSERT INTO DMS.dbo.Document_Keywords(doc_id, key_id, created_by, last_upd_by, val, pr_flag) " & _
                            "VALUES (" & dmsdocs(j, 8) & ", " & dmsdocs(j, 20) & ", 1, 1, '" & KeyVal & "', 'Y')"
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Keywords record: " & vbCrLf & SqlS)
                        dcmd.CommandText = SqlS
                        Try
                            returnv = dcmd.ExecuteNonQuery()
                        Catch ex As Exception
                        End Try
                    End If
                Next

                ' -----
                ' CREATE DOCUMENT ASSOCIATION RECORDS
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Creating Document Association records at " & Now.ToString)
                For j = 1 To NumFiles
                    '  Must have a valid document id to proceed
                    If dmsdocs(j, 8) <> "" Then

                        ' -----
                        ' CONTACT ASSOCIATION
                        '  Using dmsdocs(#, 5) = The contact association (3) id
                        If dmsdocs(j, 5) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 3, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 5) & "', '" & AccessFlg & "', '" & AccessFlg & "', '" & ReqdFlg & "')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Contact: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' If the stored contact id is not the same as the parameter contact id, 
                        ' create a second such association
                        If dmsdocs(j, 5) <> ConId And ConId <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES(1, 1, 3, " & dmsdocs(j, 8) & ", '" & ConId & "', '" & AccessFlg & "', '" & AccessFlg & "', '" & ReqdFlg & "')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Contact: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' -----
                        ' OTHER ASSOCIATIONS
                        ' Jurisdiction Association
                        '   Using dmsdocs(#, 6) = The jurisdiction association (20) id
                        If dmsdocs(j, 6) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 20, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 6) & "', 'N', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Jurisdiction: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Organization Association
                        '   Using dmsdocs(#, 7) = The organization association (8) id
                        If dmsdocs(j, 7) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 8, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 7) & "', 'N', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Organization: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Workshop Association
                        '   Using dmsdocs(#, 10) = The workshop association (2) id
                        If dmsdocs(j, 10) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 2, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 10) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Workshop: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Session Association
                        '   Using dmsdocs(#, 11) = The session association (1) id
                        If dmsdocs(j, 11) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 1, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 11) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Session: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Scheduled Session Association
                        '   Using dmsdocs(#, 12) = The scheduled session association (23) id
                        If dmsdocs(j, 12) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 23, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 12) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Scheduled Session: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Participation Association
                        '   Using dmsdocs(#, 13) = The participation association (29) id
                        If dmsdocs(j, 13) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 30, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 13) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Participation: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Workshop Registration Association
                        '   Using dmsdocs(#, 14) = The workshop registration association (7) id
                        If dmsdocs(j, 14) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 7, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 14) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Workshop Registration: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Session Registration Association
                        '   Using dmsdocs(#, 15) = The session registration association (6) id
                        If dmsdocs(j, 15) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                               "VALUES (1, 1, 6, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 15) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Session Registration: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Exam Association
                        '   Using dmsdocs(#, 16) = The exam association (16) id
                        If dmsdocs(j, 16) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 16, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 16) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Exam: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Trainer Certification Association
                        '   Using dmsdocs(#, 17) = The trainer certification (22) id
                        If dmsdocs(j, 17) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 22, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 17) & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Trainer Certification: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If

                        ' Trainer Id Association
                        '   Using dmsdocs(#, 22) = The trainer (5) id
                        If dmsdocs(j, 22) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                "VALUES (1, 1, 5, " & dmsdocs(j, 8) & ", '" & dmsdocs(j, 22) & "', '" & AccessFlg & "', '" & AccessFlg & "', '" & ReqdFlg & "')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Associations record for Trainer: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                Next

                ' -----
                ' BACKUP PDF PASSWORD IF APPLICABLE
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Creating Document Keyword record(s)")
                If PdfPassword <> "" Then
                    For j = 1 To NumFiles
                        If dmsdocs(j, 8) <> "" Then
                            SqlS = "INSERT INTO DMS.dbo.Document_Keywords(created_by, last_upd_by, key_id, doc_id, val, pr_flag) " & _
                                "VALUES (1, 1, 2, " & dmsdocs(j, 8) & ", '" & PdfPassword & "', 'N')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Document_Keywords record for document: " & vbCrLf & SqlS)
                            dcmd.CommandText = SqlS
                            Try
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                            End Try
                        End If
                    Next
                End If

                ' -----
                ' LOCATE DOMAIN RECORDS
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Locating Domain and User records at " & Now.ToString)
                If Domain <> "" Then
                    ' Get Domain User Id for these records
                    SqlS = "SELECT U.row_id " & _
                    "FROM DMS.dbo.Groups G " & _
                    "LEFT OUTER JOIN DMS.dbo.User_Group_Access U ON U.access_id=G.row_id " & _
                    "WHERE U.type_id='G' and UPPER(G.name)='" & Domain.ToUpper & "'"
                    Try
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locating UGA id for the domain " & vbCrLf & SqlS)
                        dcmd.CommandText = SqlS
                        Try
                            ddr = dcmd.ExecuteReader()
                            If ddr Is Nothing Then
                                errmsg = errmsg & "DMS database access error" & vbCrLf
                                If debug = "Y" Then mydebuglog.Debug(errmsg)
                                results = "Failure"
                                GoTo CloseOut
                            End If
                            While ddr.Read()
                                If ddr.HasRows Then
                                    ' Save document id to update
                                    UGAId = Trim(CheckDBNull(ddr(0), enumObjectType.StrType)).ToString
                                End If
                            End While
                        Catch ex As Exception
                        End Try
                        Try
                            ddr.Close()
                            ddr = Nothing
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                    If debug = "Y" Then mydebuglog.Debug("    > Found Domain UGA id " & UGAId)

                    ' Get contact information for notification
                    SqlS = "SELECT TOP 1 D.CS_EMAIL, (SELECT CASE WHEN D.ETIPS_FLG='Y' THEN 'http://www.gettips.com/' ELSE D.HOME_URL END) + " & _
                    "(SELECT CASE WHEN RIGHT(D.HOME_URL,1)='/' THEN '' ELSE '/' END)+'servicelogin.html?RNL=mydocs', C.ROW_ID AS FROM_ID, C.FST_NAME+' '+C.LAST_NAME " & _
                    "FROM siebeldb.dbo.CX_SUB_DOMAIN D " & _
                    "LEFT OUTER JOIN siebeldb.dbo.S_CONTACT C ON C.EMAIL_ADDR = D.CS_EMAIL " & _
                    "WHERE UPPER(D.DOMAIN)='" & Domain.ToUpper & "'"
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Get reply-to email address for domain: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    ReplyTo = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                    'If debug = "Y" Then ReplyTo = "bobbittc@gettips.com"
                                    ACCESS_URL = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                    'If Domain.ToUpper = "PBSA" Then ACCESS_URL = "http://www.gettips.com/servicelogin.html?RNL=mydocs"
                                    FROM_ID = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                    FROM_NAME = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                Catch ex As Exception
                                    errmsg = errmsg & "Error locating domain contact information. " & ex.ToString
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "The domain contact information was not found."
                        End If
                        Try
                            dr.Close()
                            dr = Nothing
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                    If debug = "Y" Then
                        mydebuglog.Debug("    > Reply-to Address: " & ReplyTo)
                        mydebuglog.Debug("    > Access URL: " & ACCESS_URL)
                        mydebuglog.Debug("    > FROM_ID: " & FROM_ID)
                        mydebuglog.Debug("    > FROM_NAME: " & FROM_NAME)
                    End If
                End If

                ' -----
                ' LOCATE EMPLOYEE INFORMATION
                '  If an EmpId is supplied, then use this as the reply-to address on messages generated
                If EmpId <> "" Then
                    ' Locate contact information for recipient
                    SqlS = "SELECT FST_NAME, LAST_NAME, EMAIL_ADDR, X_CON_ID, JOB_TITLE, WORK_PH_NUM, LOGIN " & _
                    "FROM siebeldb.dbo.S_EMPLOYEE " & _
                    "WHERE ROW_ID='" & EmpId & "'"
                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Getting employee information: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        dr = cmd.ExecuteReader()
                        If Not dr Is Nothing Then
                            While dr.Read()
                                Try
                                    eFST_NAME = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                    eLAST_NAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                    eEMAIL_ADDR = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                    eCON_ID = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                    eJOB_TITLE = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                                    eWORK_PH_NUM = FormatPhone(Trim(CheckDBNull(dr(5), enumObjectType.StrType)).ToString)
                                    eLOGIN = Trim(CheckDBNull(dr(6), enumObjectType.StrType)).ToString
                                Catch ex As Exception
                                    errmsg = errmsg & "Error locating employee. " & ex.ToString
                                End Try
                            End While
                        Else
                            errmsg = errmsg & "The employee was not found."
                        End If
                        Try
                            dr.Close()
                            dr = Nothing
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                    If debug = "Y" Then mydebuglog.Debug("    > Employee Found: " & eFST_NAME & " " & eLAST_NAME & ", " & eEMAIL_ADDR & ", Id: " & eCON_ID)

                    ' If employee was found, then set the reply-from appropriately
                    If eLAST_NAME <> "" And eCON_ID <> "" Then
                        FROM_NAME = eFST_NAME & " " & eLAST_NAME
                        SIGNATURE = FROM_NAME & "&#60;br /&#62;" & eJOB_TITLE & "&#60;br /&#62;" & eWORK_PH_NUM
                        ReplyTo = eEMAIL_ADDR
                        FROM_ID = eCON_ID
                    End If

                    ' Locate DMS contact id for the employee
                    If eCON_ID <> "" Then
                        SqlS = "SELECT UGA.row_id " & _
                        "FROM DMS.dbo.Users U " & _
                        "LEFT OUTER JOIN DMS.dbo.User_Group_Access UGA ON UGA.access_id=U.row_id " & _
                        "WHERE U.ext_user_id='" & eCON_ID & "' AND UGA.type_id='U'"
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locating DMS User Id of employee: " & vbCrLf & SqlS)
                        Try
                            dcmd.CommandText = SqlS
                            ddr = dcmd.ExecuteReader()
                            If Not ddr Is Nothing Then
                                Try
                                    While ddr.Read()
                                        If ddr.HasRows Then
                                            ' Save document id to update
                                            DmsEmpId = Trim(CheckDBNull(ddr(0), enumObjectType.StrType)).ToString
                                            If debug = "Y" Then mydebuglog.Debug("    > Found DMS employee id: " & DmsEmpId)
                                        End If
                                    End While
                                Catch ex As Exception
                                End Try
                            Else
                                errmsg = errmsg & "DMS database access error" & vbCrLf
                            End If
                            Try
                                ddr.Close()
                                ddr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                        End Try
                    End If
                End If

                ' -----
                ' GENERATE DOCUMENT_USERS RECORDS
                Dim WebRegId, TempUserId, SaveUserId As String
                SaveUserId = ""
                supervisor = "1"
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Creating Document_Users records")
                For j = 1 To NumFiles
                    ' If a document id was created...
                    If dmsdocs(j, 8) <> "" Then
                        ' Locate DMS user account for this contact or the parameter contact
                        '  Give priority to the parameter contact
                        If ConId <> "" Or dmsdocs(j, 5) <> "" Then
                            WebRegId = ""
                            ' Determine the contact id to use
                            If ConId <> "" Then
                                TempUserId = ConId
                            Else
                                TempUserId = dmsdocs(j, 5)
                            End If

                            ' Lookup the web registration id of the contact id if not previously looked up
                            If SaveUserId <> TempUserId Then
                                SaveUserId = TempUserId
                                SqlS = "SELECT TOP 1 X_REGISTRATION_NUM " & _
                                "FROM siebeldb.dbo.S_CONTACT " & _
                                "WHERE ROW_ID='" & TempUserId & "'"
                                If debug = "Y" Then mydebuglog.Debug("  Get contact web registration id: " & vbCrLf & SqlS)
                                Try
                                    cmd.CommandText = SqlS
                                    dr = cmd.ExecuteReader()
                                    If Not dr Is Nothing Then
                                        While dr.Read()
                                            Try
                                                WebRegId = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                            Catch ex As Exception
                                                errmsg = errmsg & "Error locating domain contact information. " & ex.ToString
                                            End Try
                                        End While
                                    Else
                                        errmsg = errmsg & "The domain contact information was not found."
                                    End If
                                    Try
                                        dr.Close()
                                        dr = Nothing
                                    Catch ex As Exception
                                    End Try
                                Catch ex As Exception
                                End Try

                                ' Retrieve the DMS user id if the web registration id was found
                                If WebRegId <> "" Then
                                    SqlS = "SELECT uga.row_id " & _
                                    "FROM DMS.dbo.User_Group_Access uga " & _
                                    "LEFT OUTER JOIN DMS.dbo.Users u on u.row_id=uga.access_id " & _
                                    "WHERE uga.type_id='U' AND u.ext_id='" & WebRegId & "'"
                                    If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Locating UGA id for the supervisor " & vbCrLf & SqlS)
                                    Try
                                        dcmd.CommandText = SqlS
                                        ddr = dcmd.ExecuteReader()
                                        If ddr Is Nothing Then
                                            errmsg = errmsg & "DMS database access error" & vbCrLf
                                            If debug = "Y" Then mydebuglog.Debug(errmsg)
                                            results = "Failure"
                                            GoTo CloseOut
                                        Else
                                            While ddr.Read()
                                                If ddr.HasRows Then
                                                    ' Save user access id to update
                                                    supervisor = Trim(CheckDBNull(ddr(0), enumObjectType.StrType)).ToString
                                                End If
                                            End While
                                        End If
                                        Try
                                            ddr.Close()
                                            ddr = Nothing
                                        Catch ex As Exception
                                        End Try
                                    Catch ex As Exception
                                    End Try
                                End If
                            End If
                            If supervisor = "" Then supervisor = "1"
                            If debug = "Y" Then mydebuglog.Debug("    > Supervisor UGA id " & supervisor)
                        End If

                        ' Provide REDO access to the document for "supervisor"
                        'SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                        '      "SELECT TOP 1 1, 1, " & dmsdocs(j, 8) & ", " & supervisor & ", 'Y', 'REDO' " & _
                        '      "FROM DMS.dbo.Document_Users " & _
                        '      "WHERE NOT EXISTS (SELECT doc_id FROM DMS.dbo.Document_Users WHERE doc_id=" & dmsdocs(j, 8) & " AND user_access_id=" & supervisor & ")"
                        SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                              "VALUES (1, 1, " & dmsdocs(j, 8) & ", " & supervisor & ", 'Y', 'REDO')"
                        If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Supervisor Document_Users record: " & vbCrLf & SqlS)
                        Try
                            dcmd.CommandText = SqlS
                            returnv = dcmd.ExecuteNonQuery()
                        Catch ex As Exception
                            'errmsg = errmsg & "Could not update the Document_Users table for supervisor" & vbCrLf & ex.Message
                            'If debug = "Y" Then mydebuglog.Debug(errmsg)
                            'results = "Failure"
                        End Try

                        ' Provide REDO access to the document for a specified employee
                        If DmsEmpId <> "0" And DmsEmpId <> "" Then
                            'SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                            '      "SELECT TOP 1 1, 1, " & dmsdocs(j, 8) & ", " & DmsEmpId & ", 'Y', 'REDO' " & _
                            '      "FROM DMS.dbo.Document_Users " & _
                            '      "WHERE NOT EXISTS (SELECT doc_id FROM DMS.dbo.Document_Users WHERE doc_id=" & dmsdocs(j, 8) & " AND user_access_id=" & DmsEmpId & ")"
                            SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                                  "VALUES (1, 1, " & dmsdocs(j, 8) & ", " & DmsEmpId & ", 'Y', 'REDO')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Employee Document_Users record: " & vbCrLf & SqlS)
                            Try
                                dcmd.CommandText = SqlS
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                                'errmsg = errmsg & "Could not update the Document_Users table for employee" & vbCrLf & ex.Message
                                'If debug = "Y" Then mydebuglog.Debug(errmsg)
                                'results = "Failure"
                            End Try
                        End If

                        ' Provide RO access for domain users
                        If UGAId <> "" Then
                            'SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                            '      "SELECT TOP 1 1, 1, " & dmsdocs(j, 8) & ", " & UGAId & ", 'N', 'RO' " & _
                            '      "FROM DMS.dbo.Document_Users " & _
                            '      "WHERE NOT EXISTS (SELECT doc_id FROM DMS.dbo.Document_Users WHERE doc_id=" & dmsdocs(j, 8) & " AND user_access_id=" & UGAId & ")"
                            SqlS = "INSERT INTO DMS.dbo.Document_Users(created_by, last_upd_by, doc_id, user_access_id, owner_flag, access_type) " & _
                                  "VALUES (1, 1, " & dmsdocs(j, 8) & ", " & UGAId & ", 'N', 'RO')"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Creating Domain Document_Users record: " & vbCrLf & SqlS)
                            Try
                                dcmd.CommandText = SqlS
                                returnv = dcmd.ExecuteNonQuery()
                            Catch ex As Exception
                                'errmsg = errmsg & "Could not update the Document_Users table for domain" & vbCrLf & ex.Message
                                'If debug = "Y" Then mydebuglog.Debug(errmsg)
                                'results = "Failure"
                            End Try
                        End If

                        ' -----
                        ' Update document count for recipient
                        If AccessFlg = "Y" And dmsdocs(j, 5) <> "" Then
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "   Updating document count for id: " & dmsdocs(j, 5))
                            DmsService.UpdDMSDocCountAsync(dmsdocs(j, 5), "", "", "", debug)
                            'DmsService.UpdDMSDocCount(dmsdocs(j, 5), "", "", "", debug)
                        End If
                    End If
                Next

                ' -----
                ' GENERATE DESTINATION BASED ON OUTPUTDEST PARAMETER
                '  Generate the destination records for subsequent use in other systems
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Creating Destination records at " & Now.ToString)
                Extension = LCase(OutFormat)               ' Set output extension for use in links
                For j = 1 To NumFiles
                    k = 0
                    Dim oEncoder As New System.Text.ASCIIEncoding()
                    Dim bytes As Byte() = oEncoder.GetBytes(dmsdocs(j, 8))
                    UserKey = "[USERKEY]"
                    If dmsdocs(j, 8) <> "" Then
                        Select Case OutputDest
                            Case "file"
                                ' Store to DMS only - no destination computed
                                dmsdocs(j, 18) = ""
                                If debug = "Y" Then mydebuglog.Debug("   Generated file output: " & dmsdocs(j, 18))
                            Case "web"
                                ' Store for use on unsecured web page - reference GetDImage service
                                PublicKey = ReverseString(ToBase64(bytes))
                                temp = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2 Final//EN""> " & Chr(10) & _
                                "<HTML> " & Chr(10) & _
                                "<HEAD> " & Chr(10) & _
                                "<TITLE>" & REPORT_NAME & "</TITLE> " & Chr(10) & _
                                "</HEAD> " & Chr(10) & _
                                "<BODY> " & Chr(10) & _
                                "<TABLE WIDTH=""100%"" HEIGHT=""100%"" BORDER=""0""> " & Chr(10) & _
                                "<TR VALIGN=""MIDDLE""><TD ALIGN=""CENTER""> "
                                If OutFormat = "jpg" Then
                                    temp = temp & "<img src=""https://hciscorm.certegrity.com/media/GetDImage.ashx?Domain=" & Domain & "&PublicKey=" & PublicKey & "&ItemName=" & dmsdocs(j, 2) & """ border=""5"">" & Chr(10)
                                Else
                                    temp = temp & "<a href=""https://hciscorm.certegrity.com/media/GetDImage.ashx?Domain=" & Domain & "&PublicKey=" & PublicKey & "&ItemName=" & dmsdocs(j, 2) & """>Get " & REPORT_NAME & "</a> " & Chr(10)
                                End If
                                temp = temp & "</TD></TR></TABLE> " & Chr(10) & _
                                "</BODY> " & Chr(10) & _
                                "</HTML> " & Chr(10)
                                dmsdocs(j, 18) = temp
                                If debug = "Y" Then mydebuglog.Debug("   Generated web output: " & dmsdocs(j, 18))
                            Case "mobile"
                                ' Store for use on secure mobile platform - TBD
                                dmsdocs(j, 18) = ""
                            Case "link"
                                ' Store for use in secured link - reference GetSImage service
                                PublicKey = ReverseString(ToBase64(bytes))
                                If ConId <> "" Then
                                    UserKey = ConIdUserKey
                                Else
                                    UserKey = dmsdocs(j, 19)
                                End If
                                dmsdocs(j, 18) = "https://hciscorm.certegrity.com/media/GetSImage.ashx?Domain=" & Domain & "&PublicKey=" & PublicKey & "&UserKey=" & UserKey & "&Ext=" & Extension & "&Debug=N"
                                If debug = "Y" Then mydebuglog.Debug("   Generated link output: " & dmsdocs(j, 18))
                            Case "image"
                                ' Store as binary - store to array dmsbin instead of dmsdoc and length in dmsdocs
                                temp = basepath & "temp\" & dmsdocs(j, 2)
                                If debug = "Y" Then mydebuglog.Debug("    Opening temp file to store to binary: " & temp)
                                Try
                                    Dim mstream As New System.IO.FileStream(temp, FileMode.OpenOrCreate, FileAccess.Read)
                                    lFileLength = mstream.Length
                                    If lFileLength > k Then
                                        k = lFileLength
                                        mydebuglog.Debug("    NumFile: " & NumFiles & " lFileLength: " & lFileLength)
                                        'ReDim Preserve dmsbin(NumFiles, lFileLength)
                                        ReDim dmsbin(NumFiles, lFileLength)
                                    End If
                                    If debug = "Y" Then mydebuglog.Debug("    > Input stream length " & lFileLength.ToString)
                                    dmsdocs(j, 18) = lFileLength.ToString
                                    Dim MyData(lFileLength) As Byte
                                    mstream.Read(MyData, 0, lFileLength)
                                    mstream.Close()
                                    mstream = Nothing
                                    dmsBytes.Add(MyData) 'Ren Hou; 1-3-2017; modified to fix error
                                    'For l = 0 To lFileLength
                                    '    dmsbin(j, l) = MyData(l)
                                    'Next
                                    MyData = Nothing
                                    If debug = "Y" Then mydebuglog.Debug("   Generated image output of " & lFileLength.ToString & " bytes")
                                Catch ex As Exception
                                    errmsg = errmsg & "Could not open temp file to store to binary: " & vbCrLf & ex.Message
                                    results = "Failure"
                                    GoTo CloseOut
                                End Try
                        End Select
                    End If
                    bytes = Nothing
                    oEncoder = Nothing
                Next

                ' -----
                ' REMOVE TEMP FILES GENERATES
                If debug <> "Y" Then
                    For j = 1 To NumFiles
                        If dmsdocs(j, 8) <> "" Then
                            temp = basepath & "temp\" & dmsdocs(j, 2)
                            If logging = "Y" Then mydebuglog.Debug("  Removing temp file:" & temp)
                            Try
                                Kill(temp)
                            Catch ex As Exception
                            End Try
                        End If
                    Next
                End If

                ' -----
                ' SENT NOTIFICATION IF APPLICABLE
                '  If the notification flag is set and email address exist, then generate/send 
                '  an email message to each contact for each document generated.
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Locating Contacts and Generating Notifications at " & Now.ToString)

                '  If a ConId was supplied, then send a single message, otherwise
                '  generate multiple messages if more than one file was generated
                Dim dmsdocid As String = ""
                If SIGNATURE = "" Then SIGNATURE = FROM_NAME
                If ConId <> "" Then
                    If NumFiles > 0 Then
                        ' Locate contact information for recipient
                        SqlS = "SELECT FST_NAME, LAST_NAME, EMAIL_ADDR, ROW_ID, X_PR_LANG_CD " & _
                        "FROM siebeldb.dbo.S_CONTACT " & _
                        "WHERE ROW_ID='" & ConId & "'"
                        If debug = "Y" Then mydebuglog.Debug("  Get recipient contact information: " & vbCrLf & SqlS)
                        Try
                            cmd.CommandText = SqlS
                            dr = cmd.ExecuteReader()
                            If Not dr Is Nothing Then
                                While dr.Read()
                                    Try
                                        FST_NAME = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                        LAST_NAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                        EMAIL_ADDR = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                        LANG_CD = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                                        If LANG_CD = "" Then LANG_CD = "ENU"

                                        ' Prepare and fix user key link if applicable
                                        temp = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                        UserKey = GenerateUserKey(temp)
                                        If debug = "Y" Then mydebuglog.Debug("   > replacing [USERKEY] with '" & UserKey & "' where applicable")
                                        For j = 1 To NumFiles
                                            temp = dmsdocs(j, 18)
                                            dmsdocid = dmsdocs(j, 8)
                                            If temp.IndexOf("[USERKEY]") > 0 Then dmsdocs(j, 18) = Replace(temp, "[USERKEY]", UserKey) 'Ren Hou; 1-3-2017; Modified to fix issue
                                            'dmsdocs(j, 18) = Replace(temp, "[USERKEY]", UserKey)
                                        Next

                                    Catch ex As Exception
                                        errmsg = errmsg & "Error locating recipient contact. " & ex.ToString
                                    End Try
                                End While
                            Else
                                errmsg = errmsg & "The recipient contact was not found."
                            End If
                            Try
                                dr.Close()
                                dr = Nothing
                            Catch ex As Exception
                            End Try
                        Catch ex As Exception
                            errmsg = errmsg & "Error locating recipient contact. " & ex.ToString
                        End Try
                        Try
                            If debug = "Y" Then mydebuglog.Debug("   > Found: " & FST_NAME & " " & LAST_NAME & ", " & EMAIL_ADDR & ", dmsdocid: " & dmsdocid)
                        Catch ex As Exception
                        End Try

                        ' If the contact has an email address, generate a MESSAGE 0094
                        If NotifyFlg = "Y" And AccessFlg = "Y" And ReplyTo <> "" And EMAIL_ADDR <> "" Then
                            'If debug = "Y" Then EMAIL_ADDR = "bobbittc@gettips.com"
                            SendTo = EMAIL_ADDR

                            MsgXml = "<messages>" & _
                            "<message send_to=""" & Left(SendTo, 128) & """ send_from=""" & Left(ReplyTo, 128) & """ from_name=""" & Left(FROM_NAME, 80) & """ from_id="""" to_id=""" & ConId & """>" & _
                            "<NumFiles>" & Trim(NumFiles.ToString) & "</NumFiles>" & _
                            "<NumRows>" & Trim(NumRows.ToString) & "</NumRows>" & _
                            "<REPORT_NAME>" & Trim(REPORT_NAME) & "</REPORT_NAME>" & _
                            "<FST_NAME>" & Trim(FST_NAME) & "</FST_NAME>" & _
                            "<LAST_NAME>" & Trim(LAST_NAME) & "</LAST_NAME>" & _
                            "<Domain>" & Trim(Domain) & "</Domain>" & _
                            "<SpecialMsg>" & Trim(SpecialMsg) & "</SpecialMsg>" & _
                            "<SPECIAL_NOTICE>" & SecurityElement.Escape(Trim(SPECIAL_NOTICE)) & "</SPECIAL_NOTICE>" & _
                            "<SIGNATURE>" & Trim(SIGNATURE) & "</SIGNATURE>" & _
                            "<AttachFlg>" & Trim(AttachFlg) & "</AttachFlg>"
                            If dmsdocid <> "" And AttachFlg = "Y" Then
                                MsgXml = MsgXml & "<DOC_ID>" & Trim(dmsdocid) & "</DOC_ID>"
                            End If
                            MsgXml = MsgXml & "</message></messages>"
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Sending MESSAGE 0094: ")
                            Try
                                If MsgXml <> "" Then
                                    results = XsltMerge(MsgXml, "683", "683", "EMAIL", LANG_CD, debug, mydebuglog)
                                    'results = http.geturl("http://hciscormsvc.certegrity.com/XSLTMergeWebServices/WebService1.asmx/XSLTMerge?XSL_SOURCE=683&XSL_SOURCE_TYPE=ID&XML_SOURCE=" & Trim(HttpUtility.UrlEncode(MsgXml)) & "&XML_SOURCE_TYPE=ID&OUTPUT_PATH=&OUTPUT_DEST=EMAIL&DEBUG_FLG=N&FTP_SITE_URL=&FTP_SITE_FOLDER=&FTP_USER=&FTP_PASS=&CM_REPORT_ID=&CACHE_KEY=&DELETE_INPUT=Y&DELETE_OUTPUT=Y&LANG_CD=" & LANG_CD, "192.168.7.61", 80, "", "")
                                End If
                            Catch ex As Exception
                                errmsg = errmsg & "Error generating MESSAGE 0094: " & ex.ToString & vbCrLf & " in XML: " & vbCrLf & MsgXml & vbCrLf
                            End Try

                            ' Prepare title of activity
                            If NumFiles = 1 And NumRows > 1 Then
                                Subject = REPORT_NAME.Trim & " were prepared for you"
                            Else
                                Subject = "A" & IIf(IsVowel(REPORT_NAME), "n ", " ") & REPORT_NAME & " was prepared for you"
                            End If

                            ' Prepare description of activity
                            Body2 = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2 Final//EN""> " & _
                           "<HTML><BODY>" & _
                           Body2 = FST_NAME & " " & LAST_NAME & ":<BR>"
                            If NumRows > NumFiles Then
                                Body2 = Body2 & "<P>This document, with " & NumRows.ToString & " pages, was prepared on your behalf and is available within your " & Domain & " portal Documents area.</P>"
                            Else
                                Body2 = Body2 & "<P>This item was prepared on your behalf and is available within your " & Domain & " portal Documents area.</P>"
                            End If
                            If SpecialMsg <> "" Then
                                Body2 = Body2 & "<P>" & SpecialMsg & "</P>"
                            End If
                            If SPECIAL_NOTICE <> "" Then
                                Body2 = Body2 & "<P><B>Please note:</b> " & SPECIAL_NOTICE & "</p>"
                            End If
                            If AttachFlg = "Y" And ACCESS_URL <> "" Then
                                Body2 = Body2 & "<P>To view this item and others prepared for you, click on <a href=" & ACCESS_URL & ">" & ACCESS_URL & "</a> and log in using your previously supplied credentials.</P>"
                            End If
                            Body2 = Body2 & "<P>If you have any questions, please respond to this email message.  Thank you.<P>" & SIGNATURE & "</BODY></HTML>"
                            Body = SqlString(Body2)

                            ' Send message # 0094
                            'SqlS = "INSERT scanner.dbo.MESSAGES (SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT, TO_ID, FROM_ID, SRC_TYPE, SRC_ID, HTML, FROM_NM) " & _
                            '       "VALUES ('" & Left(SendTo, 128) & "', '" & Left(ReplyTo, 128) & "', '" & Subject & "', '" & Body & "','N',GETDATE(),NULL,'" & ConId & "','" & _
                            '       FROM_ID & "','GenCertProd','" & dmsdocs(1, 8) & "','Y','" & Left(FROM_NAME, 80) & "')"
                            'If debug = "Y" Then mydebuglog.Debug("   > Sending email #0094: " & vbCrLf & SqlS)
                            'Try
                            'cmd.CommandText = SqlS
                            'returnv = cmd.ExecuteNonQuery()
                            'Catch ex As Exception
                            'errmsg = errmsg & "The notification email was not created: " & ex.ToString & vbCrLf
                            'End Try

                            ' Log the results
                            Letter = "1-BR52P"
                            SqlS = "INSERT INTO siebeldb.dbo.CX_LETTER_HISTORY_X (ROW_ID, CREATED, CREATED_BY,LAST_UPD,LAST_UPD_BY," & _
                                    "MODIFICATION_NUM,CONFLICT_ID,LIT_ID,OU_ID,CON_ID) " & _
                                    "VALUES ('CP" & dmsdocs(1, 8) & "', getdate(), '0-1', getdate(), '0-1', 0, 0, '" & Letter & "', '" & OrgId & "', '" & ConId & "') "
                            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Logging notification: " & vbCrLf & SqlS & vbCrLf)
                            Try
                                cmd.CommandText = SqlS
                                returnv = cmd.ExecuteNonQuery()
                            Catch ex As Exception
                                errmsg = errmsg & "The notification email was not logged: " & ex.ToString & vbCrLf
                            End Try

                            ' Generate an activity if an employee
                            If eLOGIN <> "" And TypeProd = "P" Then
                                ' Create activity id
                                ActivityId = LoggingService.GenerateRecordId("S_EVT_ACT", "N", debug)
                                If debug = "Y" Then mydebuglog.Debug("  New ActivityId: " & ActivityId)

                                ' Create activity
                                SqlS = "INSERT INTO siebeldb.dbo.S_EVT_ACT " & _
                                "(ACTIVITY_UID,ALARM_FLAG,APPT_REPT_FLG,APPT_START_DT,ASGN_MANL_FLG,ASGN_USR_EXCLD_FLG,BEST_ACTION_FLG,BILLABLE_FLG,CAL_DISP_FLG, " & _
                                "COMMENTS_LONG,CONFLICT_ID,COST_CURCY_CD,COST_EXCH_DT,CREATED,CREATED_BY,CREATOR_LOGIN,DCKING_NUM,DURATION_HRS,EMAIL_ATT_FLG, " & _
                                "EMAIL_FORWARD_FLG,EMAIL_RECIP_ADDR,EVT_PRIORITY_CD,EVT_STAT_CD,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM,NAME,OWNER_LOGIN,OWNER_PER_ID, " & _
                                "PCT_COMPLETE,PRIV_FLG,ROW_ID,ROW_STATUS,TARGET_OU_ID,TARGET_PER_ID,TEMPLATE_FLG,TMSHT_RLTD_FLG,TODO_CD,TODO_PLAN_START_DT, TODO_ACTL_END_DT) " & _
                                "SELECT '" & ActivityId & "','N','N',GETDATE(),'Y','Y','N','N','N','" & _
                                Body & "',0,'USD',GETDATE(),GETDATE(),'" & EmpId & "','" & eLOGIN & "',0,0.00,'N'," & _
                                "'N','" & SendTo & "','2-High','Done', GETDATE(),'" & EmpId & "',0, '" & Subject & "', '" & eLOGIN & "', e.ROW_ID, " & _
                                "100,'N','" & ActivityId & "','N','" & OrgId & "','" & ConId & "','N','N', 'Passports', " & _
                                "GETDATE(), GETDATE() FROM siebeldb.dbo.S_EMPLOYEE e WHERE e.LOGIN='" & eLOGIN & "' AND NOT EXISTS (SELECT ACTIVITY_UID FROM siebeldb.dbo.S_EVT_ACT WHERE ACTIVITY_UID='" & ActivityId & "')"
                                If debug = "Y" Then mydebuglog.Debug("  Generating activity: " & vbCrLf & SqlS & vbCrLf)
                                Try
                                    cmd.CommandText = SqlS
                                    returnv = cmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    errmsg = errmsg & "The activity was not generated: " & ex.ToString & vbCrLf
                                End Try

                                ' Create document association to the activity
                                SqlS = "INSERT INTO DMS.dbo.Document_Associations(created_by, last_upd_by, association_id, doc_id, fkey, pr_flag, access_flag, reqd_flag) " & _
                                    "SELECT TOP 1 1, 1, 9, " & dmsdocs(1, 8) & ", '" & ActivityId & "', '" & AccessFlg & "', '" & AccessFlg & "', 'N' " & _
                                    "FROM DMS.dbo.Document_Associations " & _
                                    "WHERE NOT EXISTS (SELECT doc_id FROM DMS.dbo.Document_Associations WHERE doc_id=" & dmsdocs(1, 8) & " AND fkey='" & ActivityId & "' AND association_id=9)"
                                If debug = "Y" Then mydebuglog.Debug("  Creating Document_Associations record for Activity: " & vbCrLf & SqlS)
                                Try
                                    dcmd.CommandText = SqlS
                                    returnv = dcmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    errmsg = errmsg & "Could not update the Document_Associations table" & vbCrLf & ex.Message
                                    If debug = "Y" Then mydebuglog.Debug(errmsg)
                                    results = "Failure"
                                    GoTo CloseOut
                                End Try
                            End If
                        End If
                    End If
                Else
                    For j = 1 To NumFiles
                        dmsdocid = dmsdocs(j, 8)
                        If dmsdocs(j, 5) <> "" And dmsdocs(j, 8) <> "" Then
                            ' Locate contact information for recipient
                            SqlS = "SELECT FST_NAME, LAST_NAME, EMAIL_ADDR, ROW_ID, X_PR_LANG_CD " & _
                            "FROM siebeldb.dbo.S_CONTACT " & _
                            "WHERE ROW_ID='" & dmsdocs(j, 5) & "'"
                            If debug = "Y" Then mydebuglog.Debug("  Get recipient contact information: " & vbCrLf & SqlS)
                            Try
                                cmd.CommandText = SqlS
                                dr = cmd.ExecuteReader()
                                If Not dr Is Nothing Then
                                    While dr.Read()
                                        Try
                                            FST_NAME = Trim(CheckDBNull(dr(0), enumObjectType.StrType)).ToString
                                            LAST_NAME = Trim(CheckDBNull(dr(1), enumObjectType.StrType)).ToString
                                            EMAIL_ADDR = Trim(CheckDBNull(dr(2), enumObjectType.StrType)).ToString
                                            LANG_CD = Trim(CheckDBNull(dr(4), enumObjectType.StrType)).ToString
                                            If LANG_CD = "" Then LANG_CD = "ENU"

                                            ' Prepare and fix user key link if applicable
                                            temp = Trim(CheckDBNull(dr(3), enumObjectType.StrType)).ToString
                                            UserKey = GenerateUserKey(temp)
                                            If debug = "Y" Then mydebuglog.Debug("   > replacing [USERKEY] with '" & UserKey & "' when applicable")
                                            temp = dmsdocs(j, 18)
                                            If temp.IndexOf("[USERKEY]") > 0 Then dmsdocs(j, 18) = Replace(temp, "[USERKEY]", UserKey) 'Ren Hou; 1-3-2017; Modified to fix issue
                                            'dmsdocs(j, 18) = Replace(temp, "[USERKEY]", UserKey)

                                        Catch ex As Exception
                                            errmsg = errmsg & "Error locating recipient contact. " & ex.ToString
                                        End Try
                                    End While
                                Else
                                    errmsg = errmsg & "The recipient contact was not found."
                                End If
                                Try
                                    dr.Close()
                                    dr = Nothing
                                Catch ex As Exception
                                End Try
                            Catch ex As Exception
                                errmsg = errmsg & "Error locating recipient contact. " & ex.ToString
                            End Try
                            Try
                                If debug = "Y" Then mydebuglog.Debug("   > Found: " & FST_NAME & " " & LAST_NAME & ", " & EMAIL_ADDR)
                            Catch ex As Exception
                            End Try

                            ' If the contact has an email address, generate a MESSAGE 0094
                            If NotifyFlg = "Y" And AccessFlg = "Y" And ReplyTo <> "" And EMAIL_ADDR <> "" Then
                                'If debug = "Y" Then EMAIL_ADDR = "bobbittc@gettips.com"
                                SendTo = EMAIL_ADDR
                                Letter = "1-BR52P"

                                MsgXml = "<messages>" & _
                                "<message send_to=""" & Left(SendTo, 128) & """ send_from=""" & Left(ReplyTo, 128) & """ from_name=""" & Left(FROM_NAME, 80) & """ from_id="""" to_id=""" & dmsdocs(j, 5) & """>" & _
                                "<NumFiles>" & Trim(NumFiles.ToString) & "</NumFiles>" & _
                                "<NumRows>" & Trim(NumRows.ToString) & "</NumRows>" & _
                                "<REPORT_NAME>" & Trim(REPORT_NAME) & "</REPORT_NAME>" & _
                                "<FST_NAME>" & Trim(FST_NAME) & "</FST_NAME>" & _
                                "<LAST_NAME>" & Trim(LAST_NAME) & "</LAST_NAME>" & _
                                "<Domain>" & Trim(Domain) & "</Domain>" & _
                                "<SpecialMsg>" & Trim(SpecialMsg) & "</SpecialMsg>" & _
                                "<SPECIAL_NOTICE>" & Trim(SPECIAL_NOTICE) & "</SPECIAL_NOTICE>" & _
                                "<SIGNATURE>" & Trim(SIGNATURE) & "</SIGNATURE>" & _
                                "<AttachFlg>" & Trim(AttachFlg) & "</AttachFlg>"
                                If dmsdocid <> "" And AttachFlg = "Y" Then
                                    MsgXml = MsgXml & "<DOC_ID>" & Trim(dmsdocid) & "</DOC_ID>"
                                End If
                                MsgXml = MsgXml & "</message></messages>"
                                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "Sending MESSAGE 0094: ")
                                Try
                                    If MsgXml <> "" Then
                                        results = XsltMerge(MsgXml, "683", "683", "EMAIL", LANG_CD, debug, mydebuglog)
                                        'results = http.geturl("http://hciscormsvc.certegrity.com/XSLTMergeWebServices/WebService1.asmx/XSLTMerge?XSL_SOURCE=683&XSL_SOURCE_TYPE=ID&XML_SOURCE=" & Trim(HttpUtility.UrlEncode(MsgXml)) & "&XML_SOURCE_TYPE=ID&OUTPUT_PATH=&OUTPUT_DEST=EMAIL&DEBUG_FLG=N&FTP_SITE_URL=&FTP_SITE_FOLDER=&FTP_USER=&FTP_PASS=&CM_REPORT_ID=&CACHE_KEY=&DELETE_INPUT=Y&DELETE_OUTPUT=Y&LANG_CD=" & LANG_CD, "192.168.7.61", 80, "", "")
                                    End If
                                Catch ex As Exception
                                    errmsg = errmsg & "Error generating MESSAGE 0094: " & ex.ToString & vbCrLf & " in XML: " & vbCrLf & MsgXml & vbCrLf
                                End Try

                                ' Log the results
                                SqlS = "INSERT INTO siebeldb.dbo.CX_LETTER_HISTORY_X (ROW_ID, CREATED, CREATED_BY,LAST_UPD,LAST_UPD_BY," & _
                                        "MODIFICATION_NUM,CONFLICT_ID,LIT_ID,OU_ID,CON_ID) " & _
                                        "VALUES ('CP" & dmsdocs(j, 8) & "', getdate(), '0-1', getdate(), '0-1', 0, 0, '" & Letter & "', '" & dmsdocs(j, 7) & "', '" & dmsdocs(j, 5) & "') "
                                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Logging notification: " & vbCrLf & SqlS & vbCrLf)
                                Try
                                    cmd.CommandText = SqlS
                                    returnv = cmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    errmsg = errmsg & "The notification email was not logged: " & ex.ToString & vbCrLf
                                End Try
                            End If
                        End If
                    Next
                End If

                ' -----
                ' MANAGE PRODUCT QUEUE
                '  Update/Insert CX_CERT_PROD_QUEUE if applicable
                If debug = "Y" Then mydebuglog.Debug(vbCrLf & "-----" & vbCrLf & "Storing product results at " & Now.ToString & vbCrLf)
                If QueueId <> "" Then
                    If NumFiles > 0 Then
                        ' Update queue entry 
                        SqlS = "UPDATE siebeldb.dbo.CX_CERT_PROD_QUEUE " & _
                            "SET EXECUTED=GETDATE(), NO_RESULTS_FLG='N' " & _
                            "WHERE ROW_ID='" & QueueId & "' AND EXECUTED IS NULL"
                    Else
                        SqlS = "UPDATE siebeldb.dbo.CX_CERT_PROD_QUEUE " & _
                            "SET EXECUTED=GETDATE(), NO_RESULTS_FLG='E', ERR_MSG='No records found for product' " & _
                            "WHERE ROW_ID='" & QueueId & "' AND EXECUTED IS NULL"
                    End If
                    If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("  Updating product queue entry: " & vbCrLf & SqlS)
                    Try
                        cmd.CommandText = SqlS
                        returnv = cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        errmsg = errmsg & "Unable to update CX_CERT_PROD_QUEUE record: " & ex.ToString & vbCrLf
                    End Try
                End If

                ' -----
                ' STORE RESULTS TO THE CX_CERT_PROD_RESULTS TABLE
                '  This provides an audit log that allows one to determine what results were generated when
                If NumFiles > 1 Then NumRows = 1 ' If broken into multiple files, then each product is only one record
                If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("    > NumFiles: " & NumFiles)
                For j = 1 To NumFiles
                    If IdentStart = "" Then IdentStart = "0"
                    If IdentEnd = "" Then IdentEnd = "0"
                    If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("    > Document Id to store: " & dmsdocs(j, 8))
                    If dmsdocs(j, 8) <> "" Then
                        If dmsdocs(j, 10) <> "" Then REG_ID = dmsdocs(j, 14)
                        If dmsdocs(j, 11) <> "" Then REG_ID = dmsdocs(j, 15)
                        ' Update or Insert a product result
                        If Trim(dmsdocs(j, 9)) = "Y" Then
                            SqlS = "UPDATE siebeldb.dbo.CX_CERT_PROD_RESULTS " & _
                            "SET LAST_UPD=GETDATE()," & _
                            "PROD_QUEUE_ID='" & QueueId & "'," & _
                            "CERT_CRSE_ID='" & CrseId & "'," & _
                            "SESS_ID='" & dmsdocs(j, 11) & "'," & _
                            "WS_ID='" & dmsdocs(j, 10) & "'," & _
                            "REG_ID='" & dmsdocs(j, 15) & "'," & _
                            "CONTACT_ID='" & ConId & "'," & _
                            "PART_ID='" & dmsdocs(j, 13) & "'," & _
                            "CURRCLM_PER_ID='" & dmsdocs(j, 17) & "'," & _
                            "SESS_PART_ID='" & dmsdocs(j, 13) & "'," & _
                            "CRSE_TSTRUN_ID='" & dmsdocs(j, 16) & "'," & _
                            "DOMAIN='" & Domain & "'," & _
                            "JURIS_ID='" & dmsdocs(j, 6) & "'," & _
                            "OU_ID='" & dmsdocs(j, 7) & "'," & _
                            "FORMAT='" & OutFormat & "'," & _
                            "DESTINATION='" & OutputDest & "'," & _
                            "NOTICE_SENT=NULL," & _
                            "CON_ID='" & dmsdocs(j, 5) & "'," & _
                            "PROD_ID='" & ProdId & "'," & _
                            "PROD_TYPE='" & TypeProd & "'," & _
                            "GENERATED='" & dmsdocs(j, 18) & "'," & _
                            "IDENT_START=" & IdentStart & "," & _
                            "IDENT_END=" & IdentEnd & "," & _
                            "EMP_ID='" & Left(EmpId, 15) & "'," & _
                            "SCHED_SESS_ID='" & dmsdocs(j, 12) & "', " & _
                            "NUM_RECS=" & NumRows.ToString & ", " & _
                            "CERT_POOL_ID='" & ID_POOL_ID & "' " & _
                            "WHERE DOC_ID='" & dmsdocs(j, 8) & "'"
                            If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("    > Updating a product results record: " & vbCrLf & SqlS & vbCrLf)
                        Else
                            SqlS = "INSERT INTO siebeldb.dbo.CX_CERT_PROD_RESULTS " & _
                            "(CONFLICT_ID,CREATED,CREATED_BY,LAST_UPD,LAST_UPD_BY,MODIFICATION_NUM," & _
                            "PROD_QUEUE_ID,CERT_CRSE_ID,SESS_ID,WS_ID,REG_ID,CONTACT_ID,PART_ID," & _
                            "CURRCLM_PER_ID,SESS_PART_ID,CRSE_TSTRUN_ID,DOMAIN,JURIS_ID," & _
                            "OU_ID,DOC_ID,FORMAT,DESTINATION,NOTICE_SENT,ROW_ID," & _
                            "CON_ID,PROD_ID,PROD_TYPE,GENERATED,IDENT_START,IDENT_END,EMP_ID,SCHED_SESS_ID,NUM_RECS,CERT_POOL_ID ) " & _
                            "SELECT TOP 1 0,GETDATE(),'0-1',GETDATE(),'0-1',0, " & _
                            "'" & QueueId & "','" & CrseId & "','" & dmsdocs(j, 11) & "','" & dmsdocs(j, 10) & "','" & dmsdocs(j, 15) & "','" & ConId & "','" & dmsdocs(j, 13) & "'," & _
                            "'" & dmsdocs(j, 17) & "','" & dmsdocs(j, 13) & "','" & dmsdocs(j, 16) & "','" & Domain & "','" & dmsdocs(j, 6) & "'," & _
                            "'" & dmsdocs(j, 7) & "','" & dmsdocs(j, 8) & "','" & OutFormat & "','" & OutputDest & "',NULL,'" & dmsdocs(j, 8) & "'," & _
                            "'" & dmsdocs(j, 5) & "','" & ProdId & "','" & TypeProd & "','" & dmsdocs(j, 18) & "', " & _
                            IdentStart & ", " & IdentEnd & ",'" & Left(EmpId, 15) & "','" & dmsdocs(j, 12) & "'," & NumRows.ToString & ",'" & ID_POOL_ID & "' " & _
                            "FROM siebeldb.dbo.CX_CERT_PROD_RESULTS " & _
                            "WHERE NOT EXISTS (SELECT ROW_ID FROM siebeldb.dbo.CX_CERT_PROD_RESULTS WHERE ROW_ID='" & dmsdocs(j, 8) & "')"
                            If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("    > Inserting a product results record: " & vbCrLf & SqlS)
                        End If
                        Try
                            cmd.CommandText = SqlS
                            returnv = cmd.ExecuteNonQuery()
                        Catch ex As Exception
                            errmsg = errmsg & "Unable to update/create CX_CERT_PROD_RESULTS record: " & ex.ToString & vbCrLf
                        End Try

                        ' If a Certification Number pool product, then update the pool
                        If JURIS_CERT_ID_FLG = "Y" And ID_POOL_ID <> "" And CERT_NUM <> "" Then
                            SqlS = "UPDATE siebeldb.dbo.CX_CERT_PROD_ID_POOL " & _
                                "SET PROD_RESULT_ID='" & dmsdocs(j, 8) & "', USED_DT=GETDATE() " & _
                                "WHERE ROW_ID='" & ID_POOL_ID & "'"
                            If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("    > Updating CX_CERT_PROD_ID_POOL record: " & vbCrLf & SqlS)
                            Try
                                cmd.CommandText = SqlS
                                returnv = cmd.ExecuteNonQuery()
                            Catch ex As Exception
                                errmsg = errmsg & "Unable to update CX_CERT_PROD_ID_POOL record: " & ex.ToString & vbCrLf
                            End Try
                        End If

                        ' Update CX_PART_CURRCLM or CX_CONTACT_CURRCLM when this is a
                        ' certification card product in order to link it to the card generated
                        ' 
                        If TypeProd = "C" Then
                            ' // REMOVED CODE TO DO THIS ON 12-4-2012 AS IT ERRORS ALL THE TIME
                            ' // PROCESS THE _CARD TABLE WITH AN AGENT INSTEAD
                            'Select Case SkillLevel
                            '    Case "Participant"
                            '        SqlS = "UPDATE siebeldb.dbo.CX_PART_CURRCLM " & _
                            '            "SET CURRENT_CARD_ID='" & dmsdocs(j, 8) & "', LAST_UPD=GETDATE() " & _
                            '            "WHERE ROW_ID IN (SELECT ROW_ID FROM siebeldb.dbo.CX_PART_CURRCLM WHERE CURRENT_SPART_ID='" & dmsdocs(j, 13) & "')"
                            '    Case "Trainer"
                            '        SqlS = "UPDATE siebeldb.dbo.CX_CONTACT_CURRCLM " & _
                            '            "SET CURRENT_CARD_ID='" & dmsdocs(j, 8) & "', LAST_UPD=GETDATE() " & _
                            '            "WHERE CURRENT_CERT_ID='" & dmsdocs(j, 17) & "'"
                            'End Select
                            'If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("    > Updating curriculum records: " & vbCrLf & SqlS & vbCrLf)
                            'Try
                            '    cmd.CommandText = SqlS
                            '    returnv = cmd.ExecuteNonQuery()
                            'Catch ex As Exception
                            '    If debug = "Y" Then mydebuglog.Debug("  Unable to update _CURRCLM record: " & vbCrLf & ex.ToString & vbCrLf)
                            'End Try

                            ' Store to the siebeltran.CX_PART_CURRCLM_CARD table for updating
                            Select Case SkillLevel
                                Case "Participant"
                                    SqlS = "INSERT siebeldb.dbo.CX_PART_CURRCLM_CARD (SESS_PART_ID, CURRENT_CARD_ID, CREATED) " & _
                                        "SELECT TOP 1 ROW_ID, '" & dmsdocs(j, 8) & "', GETDATE() " & _
                                        "FROM siebeldb.dbo.CX_SESS_PART_X WHERE ROW_ID='" & dmsdocs(j, 13) & "'"
                                Case "Trainer"
                                    SqlS = "INSERT siebeldb.dbo.CX_PART_CURRCLM_CARD (TRAINER_CERT_ID, CURRENT_CARD_ID, CREATED) " & _
                                        "SELECT TOP 1 ROW_ID, '" & dmsdocs(j, 8) & "', GETDATE() " & _
                                        "FROM siebeldb.dbo.S_CURRCLM_PER WHERE ROW_ID='" & dmsdocs(j, 17) & "'"
                            End Select
                            If debug = "Y" Or logging = "Y" Then mydebuglog.Debug("    > Inserting CX_PART_CURRCLM_CARD record as backup: " & vbCrLf & SqlS & vbCrLf)
                            Try
                                cmd.CommandText = SqlS
                                returnv = cmd.ExecuteNonQuery()
                            Catch ex2 As Exception
                                If debug = "Y" Then mydebuglog.Debug("  Unable to insert CX_PART_CURRCLM_CARD record: " & vbCrLf & ex2.ToString & vbCrLf)
                            End Try
                        End If

                    End If
                Next
            Next
        Catch oBug As Exception
            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Error: " & oBug.ToString)
            results = "Failure"
        End Try
        GoTo CloseOut

ErrorQueue:
        ' ============================================
        ' Error out a queue entry in CX_CERT_PROD_QUEUE if applicable
        If QueueId <> "" Then
            Try
                dr.Close()
                dr = Nothing
            Catch ex As Exception
            End Try

            SqlS = "UPDATE siebeldb.dbo.CX_CERT_PROD_QUEUE " & _
                "SET EXECUTED=GETDATE(), NO_RESULTS_FLG='E', ERR_MSG='" & Trim(errmsg) & "' " & _
                "WHERE ROW_ID='" & QueueId & "' AND EXECUTED IS NULL"
            If debug = "Y" Then mydebuglog.Debug(vbCrLf & "  Updating product queue with error: " & errmsg & vbCrLf & SqlS)
            Try
                cmd.CommandText = SqlS
                returnv = cmd.ExecuteNonQuery()
                If returnv = 0 Then
                    errmsg = errmsg & "Problem with error update"
                End If
            Catch ex2 As Exception
                errmsg = errmsg & "Problem with error update: " & ex2.Message
                results = "Failure"
            End Try
        End If

CloseOut:
        ' ============================================
        ' Close database connections and objects
        Try
            ' hcidb1
            dr = Nothing
            con.Dispose()
            con = Nothing
            cmd.Dispose()
            cmd = Nothing
            'dms
            ddr = Nothing
            dcon.Dispose()
            dcon = Nothing
            dcmd.Dispose()
            dcmd = Nothing
            ' others
            tempstream = Nothing
            myparam = Nothing
            MyCB = Nothing
            'da = Nothing
            ds = Nothing
        Catch ex As Exception
            errmsg = errmsg & vbCrLf & "Unable to close the database connection. "
        End Try

        Try
            Release(Database, True)
            Database = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Database)
        Catch ex As Exception
        End Try

        Try
            Release(myExportOptions, True)
            myExportOptions = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExportOptions)
        Catch ex As Exception
        End Try

        Try
            Release(myDiskFileDestinationOptions, True)
            myDiskFileDestinationOptions = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myDiskFileDestinationOptions)
        Catch ex As Exception
        End Try

        Try
            Release(Report, True)
            Report = Nothing
            Report.Dispose()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Report)
        Catch ex As Exception
        End Try


CloseOut2:
        ' ============================================
        ' RETURN THE RESULTS TO THE CALLING SYSTEM
        '  The results are stored in an XML document as follows
        '	<?xml version="1.0" encoding="utf-8"?>
        '	<products number="2" queueid="">
        '	  <document id="123456" length="45" /><![CDATA[ ... ]]></document>
        '	  <document id="123457" length="48" /><![CDATA[ ... ]]></document>
        '     <error />
        '	</product>
        Dim odoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim resultsDeclare As System.Xml.XmlDeclaration
        Dim resultsRoot As System.Xml.XmlElement
        Dim resultsItem As System.Xml.XmlElement

        ' Create container with results
        resultsDeclare = odoc.CreateXmlDeclaration("1.0", Nothing, String.Empty)
        odoc.InsertBefore(resultsDeclare, odoc.DocumentElement)

        ' Create root node
        resultsRoot = odoc.CreateElement("products")
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        AddXMLAttribute(odoc, resultsRoot, "number", NumFiles.ToString)
        AddXMLAttribute(odoc, resultsRoot, "queueid", QueueId)
        odoc.InsertAfter(resultsRoot, resultsDeclare)
        Try
            If NumFiles > 0 And errmsg = "" Then results = "Success" Else results = "Failure"
            If debug <> "T" Then
                For j = 1 To NumFiles
                    resultsItem = odoc.CreateElement("document")
                    Select Case OutputDest
                        Case "image"
                            lFileLength = Val(dmsdocs(j, 18))
                            Dim MyData(lFileLength) As Byte
                            MyData = dmsBytes(j - 1) 'Ren Hou; 1-3-2017; modified to fix error
                            mydebuglog.Debug(vbCrLf & " MyData Upper bound: " & j & ": " & MyData.GetUpperBound(0))
                            'For k = 0 To lFileLength
                            '    MyData(k) = dmsbin(j, k)
                            'Next
                            'mydebuglog.Debug(vbCrLf & " k MyData: " & k & "; " & MyData(k).ToString)
                            resultsItem.InnerText = "<![CDATA[" & ToBase64(MyData) & "]]>"
                            AddXMLAttribute(odoc, resultsItem, "length", Len(dmsdocs(j, 18).Trim).ToString)
                        Case "web"
                            resultsItem.InnerText = "<![CDATA[" & dmsdocs(j, 18) & "]]>"
                        Case "link"
                            resultsItem.InnerText = "<![CDATA[" & dmsdocs(j, 18) & "]]>"
                        Case "file"
                            resultsItem.InnerText = ""
                        Case "mobile"
                            resultsItem.InnerText = ""
                    End Select
                    AddXMLAttribute(odoc, resultsItem, "id", dmsdocs(j, 8))
                    resultsRoot.AppendChild(resultsItem)
                Next
            End If
            AddXMLChild(odoc, resultsRoot, "results", Trim(results))
            ' Error message
            If errmsg <> "" Then AddXMLChild(odoc, resultsRoot, "error", Trim(errmsg))
        Catch ex As Exception
            mydebuglog.Debug(vbCrLf & "  ex: " & ex.Message & j & " : " & k)
            AddXMLChild(odoc, resultsRoot, "error", "Unable to create proper XML return document")
        End Try

        ' ============================================
        ' RELEASE VARIABLES
        dmsdocs = Nothing
        dmsbin = Nothing

        ' ============================================
        ' CLOSE THE LOG FILE
        If Trim(errmsg) <> "" Then myeventlog.Error("GenCertProd : Error: " & Trim(errmsg))
        'myeventlog.Info("GenCertProd : Results: " & results & " for QueueId: " & QueueId)
        myeventlog.Info("GenCertProd : Results: " & results & " for QueueId: " & QueueId & ", SrcId: " & SrcId)

        If debug = "Y" Or (logging = "Y" And debug <> "T") Then
            Try
                If Trim(errmsg) <> "" Then mydebuglog.Debug(vbCrLf & "Error: " & Trim(errmsg))
                If debug = "Y" Then
                    mydebuglog.Debug("Trace Log Ended " & Now.ToString)
                    mydebuglog.Debug("----------------------------------")
                End If
            Catch ex As Exception
            End Try
        End If

        Try
            fs.Flush()
            fs.Close()
            fs.Dispose()
            fs = Nothing
        Catch ex As Exception
        End Try

        ' ============================================
        ' Log Performance Data
        If debug <> "T" Then
            ' ============================================
            ' Log performance
            Try
                'LoggingService.LogPerformanceDataAsync(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, debug)
                Dim VersionNum As String = "102"
                LoggingService.LogPerformanceData2(System.Environment.MachineName.ToString, System.Reflection.MethodBase.GetCurrentMethod.Name.ToString, LogStartTime, VersionNum, debug)
            Catch ex As Exception
            End Try
        End If

        ' ============================================
        ' Return results
        Return odoc

    End Function

    ' ============================================
    ' SUPPORT FUNCTIONS
    Public Function XsltMerge(ByVal MergeData As String, ByVal TemplateId As String, ByVal DataId As String, _
     ByVal Destination As String, ByVal LangCd As String, ByVal Debug As String, ByRef mydebuglog As ILog) As Boolean
        ' This function is used to generate a service call to the XsltMerge service for sending email

        ' Declarations
        Dim Service, IpAddress As String
        Dim DataToSend, TransformedData, ResultString As String

        ' Prepare link
        If Debug = "Y" Then
            Service = "http://scormsvc4.certegrity.com/XSLTMergeWebServices/WebService1.asmx/XSLTMergeAsync?"
            'IpAddress = "192.168.7.51"
            IpAddress = System.Configuration.ConfigurationManager.AppSettings.Get("scormsvc4_ip")
        Else
            Service = "http://hciscormsvc.certegrity.com/XSLTMergeWebServices/WebService1.asmx/XSLTMergeAsync?"
            'IpAddress = "192.168.7.61"
            IpAddress = System.Configuration.ConfigurationManager.AppSettings.Get("hciscormsvc_farm")
        End If

        ' Prepare data   
        ResultString = ""
        If Destination = "" Then Destination = "EMAIL"
        TransformedData = HttpUtility.UrlEncode(MergeData)
        If DataId.Trim <> "" Then
            DataToSend = "XSL_SOURCE=" & TemplateId.Trim & "&XSL_SOURCE_TYPE=ID&XML_SOURCE=" & TransformedData.Trim & _
            "&XML_SOURCE_TYPE=DATA&OUTPUT_PATH=C:\Program%20Files%20(x86)\XSLTMerge\Files\" & DataId.Trim & "&OUTPUT_DEST=" & Destination.Trim.ToUpper & "&DEBUG_FLG=" & Debug & _
            "&FTP_SITE_URL=&FTP_SITE_FOLDER=&FTP_USER=&FTP_PASS=&CM_REPORT_ID=&DELETE_FILE=Y&CACHE_KEY=&DELETE_INPUT=Y&DELETE_OUTPUT=Y&LANG_CD=" & LangCd
        Else
            DataToSend = "XSL_SOURCE=" & TemplateId.Trim & "&XSL_SOURCE_TYPE=ID&XML_SOURCE=" & TransformedData.Trim & _
            "&XML_SOURCE_TYPE=DATA&OUTPUT_PATH=&OUTPUT_DEST=" & Destination.Trim.ToUpper & "&DEBUG_FLG=" & Debug & _
            "&FTP_SITE_URL=&FTP_SITE_FOLDER=&FTP_USER=&FTP_PASS=&CM_REPORT_ID=&DELETE_FILE=Y&CACHE_KEY=&DELETE_INPUT=Y&DELETE_OUTPUT=Y&LANG_CD=" & LangCd
        End If

        If Debug = "Y" Then
            mydebuglog.Debug(vbCrLf & " Function XsltMerge===========" & vbCrLf)
            mydebuglog.Debug("  MergeData: " & MergeData)
            mydebuglog.Debug("  TemplateId: " & TemplateId)
            mydebuglog.Debug("  Destination: " & Destination)
            mydebuglog.Debug("  LangCd: " & LangCd)
            mydebuglog.Debug("  DataId: " & DataId)
            mydebuglog.Debug("  DataToSend: " & DataToSend)
            mydebuglog.Debug("  TransformedData: " & TransformedData & vbCrLf)
        End If

        Try
            Dim http As New simplehttp()
            ResultString = http.geturl(Service & DataToSend, IpAddress, 80, "", "")
        Catch ex As Exception
            If Debug = "Y" Then mydebuglog.Debug("  Error: " & ex.Message)
        End Try
        If Debug = "Y" Then
            mydebuglog.Debug("  ResultString: " & ResultString)
            mydebuglog.Debug(" Function XsltMerge===========" & vbCrLf)
        End If

        If InStr(LCase(ResultString), "success") > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function Release(ByRef ComObject As Object, ByVal collect As Boolean) As Integer
        Dim Result As Integer
        Try
            Result = System.Runtime.InteropServices.Marshal.ReleaseComObject(ComObject)
            ComObject = Nothing

            If (collect) Then
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If
        Catch
        End Try
        Return Result
    End Function

    Public Function ReleaseObject(ByVal dllname As String) As String
        Try
            Dim hAddress As Long = GetModuleHandle(dllname)
            While hAddress >= 1
                FreeLibrary(hAddress)
                hAddress = GetModuleHandle(dllname)
            End While
            System.IO.File.Delete(dllname)

        Catch ex As Exception
        End Try
        Return "true"
    End Function

    Public Function TranslateDataType(ByVal formatcode As String) As Integer
        Select Case formatcode
            Case "pdf"
                TranslateDataType = 1
            Case "doc"
                TranslateDataType = 2
            Case "xls"
                TranslateDataType = 3
            Case "tif"
                TranslateDataType = 4
            Case "jpg"
                TranslateDataType = 5
            Case "png"
                TranslateDataType = 6
            Case "swf"
                TranslateDataType = 7
            Case Else
                TranslateDataType = 0
        End Select
    End Function

    Public Function GetCategory(ByVal SkillLevel As String, ByVal TypeProd As String) As String
        ' Given a skill level and a product, determines the DMS category to use
        '  Returns nothing if there is no category for this
        '  If no category is returned, the Document_Category record is not created
        GetCategory = ""
        Select Case SkillLevel
            Case "Participant"
                Select Case TypeProd
                    Case "R"    ' Course Completion Certificates
                        GetCategory = "128"         ' Participant Documents
                    Case "C"    ' Certificate Card
                        GetCategory = "112"         ' Certification Documents
                    Case "W"    ' Retail Establishment Certificate
                        GetCategory = "112"         ' Certification Documents
                    Case "P"    ' Passport
                        GetCategory = "109"         ' Order Documents
                    Case "V"    ' Exam Voucher
                        GetCategory = "109"         ' Order Documents
                    Case "F"    ' Regulatory Form
                        GetCategory = "128"         ' Participant Documents
                    Case "T"    ' Exam Completion Certificate
                        GetCategory = "112"         ' Certification Documents
                End Select
            Case "Trainer"
                Select Case TypeProd
                    Case "R"
                        GetCategory = "127"         ' Trainer Documents
                    Case "C"
                        GetCategory = "112"         ' Certification Documents
                    Case "W"
                        GetCategory = "112"         ' Certification Documents
                    Case "P"
                        GetCategory = "109"         ' Order Documents
                    Case "V"
                        GetCategory = "109"         ' Order Documents
                    Case "F"
                        GetCategory = "112"         ' Certification Documents
                    Case "T"
                        GetCategory = "112"         ' Certification Documents
                End Select
            Case "Master Trainer"
                Select Case TypeProd
                    Case "R"
                        GetCategory = "96"          ' Master Trainer Documents
                    Case "C"
                        GetCategory = "112"         ' Certification Documents
                    Case "W"
                        GetCategory = "112"         ' Certification Documents
                    Case "P"
                        GetCategory = "109"         ' Order Documents
                    Case "V"
                        GetCategory = "109"         ' Order Documents
                    Case "T"
                        GetCategory = "112"         ' Certification Documents
                End Select
        End Select
    End Function

    Public Function GetKeyword(ByVal SkillLevel As String, ByVal TypeProd As String) As String
        ' Given a skill level and a product, determines the DMS keyword to assign to a product
        '  Returns nothing if there is no keyword for this
        '  If no category is returned, the Document_Keyword record is not created
        GetKeyword = ""
        Select Case SkillLevel
            Case "Participant"
                Select Case TypeProd
                    Case "R"    ' Course Completion Certificates
                        GetKeyword = "7"         ' Participant Documents
                    Case "C"    ' Certificate Card
                        GetKeyword = "7"         ' Certification Documents
                    Case "W"    ' Retail Establishment Certificate
                        GetKeyword = "8"         ' Certification Documents
                    Case "P"    ' Passport
                        GetKeyword = "8"         ' Order Documents
                    Case "V"    ' Exam Voucher
                        GetKeyword = "8"         ' Order Documents
                    Case "F"    ' Regulatory Form
                        GetKeyword = "8"         ' Participant Documents
                    Case "T"    ' Exam Completion Certificate
                        GetKeyword = "7"         ' Certification Documents
                End Select
            Case "Trainer"
                Select Case TypeProd
                    Case "R"
                        GetKeyword = "3"         ' Trainer Documents
                    Case "C"
                        GetKeyword = "3"         ' Certification Documents
                    Case "W"
                        GetKeyword = "8"         ' Certification Documents
                    Case "P"
                        GetKeyword = "8"         ' Order Documents
                    Case "V"
                        GetKeyword = "8"         ' Order Documents
                    Case "F"
                        GetKeyword = "8"         ' Certification Documents
                    Case "T"
                        GetKeyword = "3"         ' Certification Documents
                End Select
            Case "Master Trainer"
                Select Case TypeProd
                    Case "R"
                        GetKeyword = "5"          ' Master Trainer Documents
                    Case "C"
                        GetKeyword = "5"         ' Certification Documents
                    Case "W"
                        GetKeyword = "8"         ' Certification Documents
                    Case "P"
                        GetKeyword = "8"         ' Order Documents
                    Case "V"
                        GetKeyword = "8"         ' Order Documents
                    Case "T"
                        GetKeyword = "5"         ' Certification Documents
                End Select
        End Select
    End Function

    ' =================================================
    ' EMAIL
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
    ' PDF FUNCTIONS
    Public Function ConvertPDFToImage(ByVal strInputPDF As String, ByVal strOutputImage As String, _
        ByVal strUsername As String, ByVal strPassword As String, _
        ByVal lngXDPI As Long, ByVal lngYDPI As Long, ByVal lngBitCount As Long, _
        ByVal enumCompression As CompressionType, ByVal lngQuality As Long, _
        ByVal blnGreyscale As Boolean, ByVal blnMultipage As Boolean, _
        ByVal lngFirstPage As Long, ByVal lngLastPage As Long) As Long

        ' Register PDF to Image Converter Library
        PDFToImageSetCode("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")

        ConvertPDFToImage = PDFToImageConverter(strInputPDF, strOutputImage, strUsername, strPassword, _
                            lngXDPI, lngYDPI, lngBitCount, enumCompression, lngQuality, blnGreyscale, _
                            blnMultipage, lngFirstPage, lngLastPage)

    End Function

    Public Function ConvertPDFToImageEx(ByVal strInputPDF As String, ByVal strOutputImage As String, _
        ByVal strUsername As String, ByVal strPassword As String, ByVal intPageSizeMode As Long, _
        ByVal lngXDPI As Long, ByVal lngYDPI As Long, ByVal lngBitCount As Long, _
        ByVal enumCompression As CompressionType, ByVal lngQuality As Long, _
        ByVal blnGreyscale As Boolean, ByVal blnMultipage As Boolean, _
        ByVal lngFirstPage As Long, ByVal lngLastPage As Long) As Long

        ' Register PDF to Image Converter Library
        PDFToImageSetCode("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")

        ConvertPDFToImageEx = PDFToImageConverterEx(strInputPDF, strOutputImage, strUsername, strPassword, _
                            intPageSizeMode, lngXDPI, lngYDPI, lngBitCount, enumCompression, lngQuality, blnGreyscale, _
                            blnMultipage, lngFirstPage, lngLastPage)
    End Function

    Public Function GetPageWidth(ByVal strInputPDF As String, ByVal intPage As Long) As Long

        ' Register PDF to Image Converter Library
        PDFToImageSetCode("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")

        GetPageWidth = PDFToImageGetPageWidth(strInputPDF, intPage)

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

    Private Sub CreateXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
    End Sub

    Private Sub AddXMLAttribute(ByVal xmldoc As XmlDocument, _
        ByVal xmlnode As XmlElement, ByVal attribute As String, _
        ByVal attributevalue As String)
        ' Used to add an attribute to a specified node

        Dim newAtt As XmlAttribute

        newAtt = xmldoc.CreateAttribute(attribute)
        newAtt.Value = attributevalue
        xmlnode.Attributes.Append(newAtt)
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
    ' DATABASE FUNCTIONS
    Public Function convertToADODB(ByRef table As Data.DataTable, ByRef mydebuglog As ILog, ByVal debug As String) As ADODB.Recordset
        Dim result As New ADODB.Recordset
        result.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Dim resultFields As ADODB.Fields = result.Fields
        Dim col As DataColumn

        For Each col In table.Columns
            resultFields.Append(col.ColumnName, TranslateType(col.DataType), col.MaxLength, col.AllowDBNull = ADODB.FieldAttributeEnum.adFldIsNullable)
        Next

        result.Open(System.Reflection.Missing.Value, System.Reflection.Missing.Value, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic, 0)

        If debug = "Y" Then mydebuglog.Debug("    .. Converting to ADODB recordset ..")
        For Each row As DataRow In table.Rows
            result.AddNew(System.Reflection.Missing.Value, System.Reflection.Missing.Value)
            For i As Integer = 0 To table.Columns.Count - 1
                Try
                    If debug = "Y" Then mydebuglog.Debug("    .. converting value: " & row(i).ToString)
                    If IsDBNull(row(i)) Then
                        If resultFields(i).Type = ADODB.DataTypeEnum.adNumeric Or resultFields(i).Type = ADODB.DataTypeEnum.adInteger Or resultFields(i).Type = ADODB.DataTypeEnum.adCurrency Then
                            resultFields(i).Value = 0
                        Else
                            resultFields(i).Value = ""
                        End If
                    Else
                        resultFields(i).Value = row(i).ToString
                    End If
                Catch ex As Exception
                    If debug = "Y" Then mydebuglog.Debug("        > conversion error error: " & ex.ToString)
                    resultFields(i).Value = ""
                End Try
            Next
        Next
        If debug = "Y" Then mydebuglog.Debug("    .. Converted to ADODB recordset ..")
        Return result
    End Function

    Private Function TranslateType(ByVal columnType As Type) As ADODB.DataTypeEnum
        Select Case columnType.UnderlyingSystemType.ToString()
            Case "System.Boolean"
                Return ADODB.DataTypeEnum.adBoolean
            Case "System.Byte"
                Return ADODB.DataTypeEnum.adUnsignedTinyInt
            Case "System.Char"
                Return ADODB.DataTypeEnum.adChar
            Case "System.DateTime"
                Return ADODB.DataTypeEnum.adDate
            Case "System.Decimal"
                Return ADODB.DataTypeEnum.adCurrency
            Case "System.Double"
                Return ADODB.DataTypeEnum.adDouble
            Case "System.Int16"
                Return ADODB.DataTypeEnum.adSmallInt
            Case "System.Int32"
                Return ADODB.DataTypeEnum.adInteger
            Case "System.Int64"
                Return ADODB.DataTypeEnum.adBigInt
            Case "System.SByte"
                Return ADODB.DataTypeEnum.adTinyInt
            Case "System.Single"
                Return ADODB.DataTypeEnum.adSingle
            Case "System.UInt16"
                Return ADODB.DataTypeEnum.adUnsignedSmallInt
            Case "System.UInt32"
                Return ADODB.DataTypeEnum.adUnsignedInt
            Case "System.UInt64"
                Return ADODB.DataTypeEnum.adUnsignedBigInt
            Case "System.String"
                Return ADODB.DataTypeEnum.adVarChar
            Case Else
                Return ADODB.DataTypeEnum.adVarChar
        End Select
    End Function

    Public Function convertDataReaderToDataSet(ByVal reader As SqlDataReader) As System.Data.DataSet
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet
        Dim dataRow As DataRow
        Dim columnName As String
        Dim column As DataColumn
        Dim schemaTable As DataTable
        Dim dataTable As DataTable

        Do
            ' Create new data table
            schemaTable = reader.GetSchemaTable
            dataTable = New DataTable
            If Not IsDBNull(schemaTable) Then
                ' A query returning records was executed
                Dim i As Integer
                For i = 0 To schemaTable.Rows.Count - 1
                    dataRow = schemaTable.Rows(i)
                    ' Create a column name that is unique in the data table
                    columnName = dataRow("ColumnName")
                    'Add the column definition to the data table
                    column = New DataColumn(columnName, CType(dataRow("DataType"), Type))
                    dataTable.Columns.Add(column)
                Next
                dataSet.Tables.Add(dataTable)

                'Fill the data table we just created
                While reader.Read()
                    dataRow = dataTable.NewRow()
                    For i = 0 To reader.FieldCount - 1
                        dataRow(i) = reader(i)
                    Next
                    dataTable.Rows.Add(dataRow)
                End While
            Else
                'No records were returned
                column = New DataColumn("RowsAffected")
                dataTable.Columns.Add(column)
                dataSet.Tables.Add(dataTable)
                dataRow = dataTable.NewRow()
                dataRow(0) = reader.RecordsAffected
                dataTable.Rows.Add(dataRow)
            End If
        Loop While reader.NextResult()
        Return dataSet
    End Function

    Public Function convertDRToDS(ByVal reader As SqlDataReader) As System.Data.DataSet
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet
        Dim dataRow As DataRow
        Dim columnName As String
        Dim column As DataColumn
        Dim schemaTable As DataTable
        Dim dataTable As DataTable

        Do
            ' Create new data table
            schemaTable = reader.GetSchemaTable
            dataTable = New DataTable
            If Not IsDBNull(schemaTable) Then
                ' A query returning records was executed
                Dim i As Integer
                For i = 0 To schemaTable.Rows.Count - 1
                    dataRow = schemaTable.Rows(i)
                    ' Create a column name that is unique in the data table
                    columnName = dataRow("ColumnName")
                    'Add the column definition to the data table
                    column = New DataColumn(columnName, CType(dataRow("DataType"), Type))
                    dataTable.Columns.Add(column)
                Next
                dataSet.Tables.Add(dataTable)

                'Fill the data table we just created
                While reader.Read()
                    dataRow = dataTable.NewRow()
                    For i = 0 To reader.FieldCount - 1
                        dataRow(i) = reader(i)
                    Next
                    dataTable.Rows.Add(dataRow)
                End While
            Else
                'No records were returned
                column = New DataColumn("RowsAffected")
                dataTable.Columns.Add(column)
                dataSet.Tables.Add(dataTable)
                dataRow = dataTable.NewRow()
                dataRow(0) = reader.RecordsAffected
                dataTable.Rows.Add(dataRow)
            End If
        Loop While reader.NextResult()
        Return dataSet
    End Function

    ' =================================================
    ' STRING FUNCTIONS
    Public Function GenerateUserKey(ByVal UserKey As String) As String
        ' Generates an encoded string from a provided string
        '  Converts the string to Base64 and then reverses it
        If Len(UserKey.Trim) > 0 Then
            Dim oEncoder As New System.Text.ASCIIEncoding()
            Dim byteu As Byte() = oEncoder.GetBytes(UserKey)
            GenerateUserKey = ReverseString(ToBase64(byteu))
            byteu = Nothing
            oEncoder = Nothing
        Else
            GenerateUserKey = ""
        End If
    End Function

    Public Function IsVowel(ByVal InputString As String) As Boolean
        ' Determines if the first letter in the supplied string is a vowel or not
        Dim temp As String
        IsVowel = False
        If Len(InputString.Trim) > 0 Then
            temp = UCase(Left(InputString, 1))
            If temp = "A" Or temp = "E" Or temp = "I" Or temp = "O" Or temp = "U" Then
                IsVowel = True
            End If
        End If
    End Function

    Public Function ReverseString(ByVal InputString As String) As String
        ' Reverses a string
        Dim lLen As Long, lCtr As Long
        Dim sChar As String
        Dim sAns As String
        sAns = ""
        lLen = Len(InputString)
        For lCtr = lLen To 1 Step -1
            sChar = Mid(InputString, lCtr, 1)
            sAns = sAns & sChar
        Next
        ReverseString = sAns
    End Function

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
        ' Make a string safe for use in a SQL query
        Dim temp As String
        Dim outstring As String
        Dim i As Integer
        Dim tchar, ochar As String

        If Len(Instring) = 0 Or Instring Is Nothing Then
            SqlString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp)
            tchar = Mid(temp, i, 1)
            ochar = ""
            Select Case Asc(tchar)
                Case 39
                    ochar = "''"
                Case 145
                    ochar = "''"
                Case 146
                    ochar = "''"
                Case 147
                    ochar = Chr(34)
                Case 148
                    ochar = Chr(34)
                Case Else
                    If Asc(tchar) > 31 And Asc(tchar) < 127 Then ochar = tchar
            End Select
            outstring = outstring & ochar
        Next
        SqlString = outstring
    End Function

    Function CleanString(ByVal Instring As String) As String
        ' Remove non-standard chars from a string
        Dim outstring As String
        Dim i As Integer
        Dim tchar, ochar As String
        outstring = ""
        For i = 1 To Len(Instring)
            tchar = Mid(Instring, i, 1)
            ochar = ""
            Select Case Asc(tchar)
                Case 39
                    ochar = "''"
                Case 145
                    ochar = "''"
                Case 146
                    ochar = "''"
                Case 147
                    ochar = Chr(34)
                Case 148
                    ochar = Chr(34)
                Case Else
                    If Asc(tchar) > 31 And Asc(tchar) < 127 Then ochar = tchar
            End Select
            outstring = outstring & ochar
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
        ElseIf ObjectType = enumObjectType.DteType And IsDBNull(obj) Then
            objReturn = Now
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

    Public Function FormatPhone(ByVal PhoneNum As String) As String
        ' Format a phone number in the form 9999999999 or 9999999999999
        Dim i As Integer
        FormatPhone = PhoneNum
        i = Len(PhoneNum)
        If IsNumeric(PhoneNum) Then
            If i = 10 Then
                FormatPhone = "(" & Left(PhoneNum, 3) & ") " & Mid(PhoneNum, 4, 3) & "-" & Right(PhoneNum, 4)
            End If
            If i = 13 Then
                FormatPhone = "(" & Left(PhoneNum, 3) & ") " & Mid(PhoneNum, 4, 3) & "-" & Mid(PhoneNum, 7, 4) & " x" & Right(PhoneNum, 3)
            End If
        End If
    End Function

    Public Function ToBase64(ByVal data() As Byte) As String
        ' Encode a Base64 string
        If data Is Nothing Then Throw New ArgumentNullException("data")
        Return Convert.ToBase64String(data)
    End Function

    Public Function FromBase64(ByVal base64 As String) As String
        ' Decode a Base64 string
        Dim results As String
        If base64 Is Nothing Then Throw New ArgumentNullException("base64")
        results = System.Text.Encoding.ASCII.GetString(Convert.FromBase64String(base64))
        Return results
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

    Public Function StringToBytes(ByVal str As String) As Byte()
        ' Convert a random string to a byte array
        ' e.g. "abcdefg" to {a,b,c,d,e,f,g}
        Dim s As Char()
        s = str.ToCharArray
        Dim b(s.Length - 1) As Byte
        Dim i As Integer
        Try
            For i = 0 To s.Length - 1
                If Asc(s(i)) > 0 And Asc(s(i)) < 128 Then
                    b(i) = Convert.ToByte(s(i))
                Else
                    b(i) = Convert.ToByte(" ")
                End If
            Next
        Catch ex As Exception
        End Try
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
    ' IMAGE FUNCTIONS
    Public Sub SetImageProperty(ByVal Img As Image, ByVal PID As Int32, ByVal Data() As Byte, ByVal Type As ExifDataTypes)
        Dim P As System.Drawing.Imaging.PropertyItem = Img.PropertyItems(0)
        P.Id = PID
        P.Value = Data
        P.Type = Type
        P.Len = Data.Length
        Img.SetPropertyItem(P)
    End Sub

    ' =================================================
    ' SHELL FUNCTION
    Public Function ShellandWait(ByVal ProcessPath As String, ByRef mydebuglog As ILog, ByVal Debug As String) As String
        Dim objProcess As System.Diagnostics.Process
        Dim errorflg, myprocess, mycmd As String
        Dim argloc As Integer
        argloc = InStr(ProcessPath, " ") - 1
        myprocess = Left(ProcessPath, argloc)
        myprocess = Replace(myprocess, "\", "\\")
        mycmd = Right(ProcessPath, Len(ProcessPath) - argloc)
        If Debug = "Y" Then mydebuglog.Debug("   .. executing: " & myprocess & " " & mycmd)
        errorflg = "Y"
        Try
            objProcess = New System.Diagnostics.Process()
            objProcess.StartInfo.WindowStyle = Diagnostics.ProcessWindowStyle.Normal
            objProcess.StartInfo.FileName = myprocess
            objProcess.StartInfo.UseShellExecute = True
            objProcess.StartInfo.RedirectStandardOutput = False
            objProcess.StartInfo.RedirectStandardError = False
            objProcess.StartInfo.Arguments = mycmd
            Try
                objProcess.Start()
                objProcess.WaitForExit()
                If Debug = "Y" Then mydebuglog.Debug("   .. exit code: " & objProcess.ExitCode.ToString)
                If objProcess.ExitCode.ToString = "0" Then errorflg = "N"
            Catch ex As Exception
                If Debug = "Y" Then mydebuglog.Debug("   .. unable to execute shell: " & ex.ToString)
            End Try

            'Free resources associated with this process
            objProcess.Close()
        Catch
        End Try
        Return errorflg
    End Function

    ' Function to pause for the seconds specified
    Sub Delay(ByVal dblSecs As Double)

        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
        Loop
    End Sub

    ' =================================================
    ' HTTP PROXY CLASS
    Class simplehttp
        Public Function geturl(ByVal url As String, ByVal proxyip As String, ByVal port As Integer, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Referer = ""
            req.Headers.[Set]("Accept-Language", "en,en-us")
            Dim stream_in As StreamReader

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don’t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim response As String = ""
            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                stream_in = New StreamReader(resp.GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function

        Public Function getposturl(ByVal url As String, ByVal postdata As String, ByVal proxyip As String, ByVal port As Short, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 25000
            req.CookieContainer = New CookieContainer()
            req.Method = "POST"
            req.ContentType = "application/x-www-form-urlencoded"
            req.ContentLength = postdata.Length
            req.Referer = ""

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don’t use proxy credentials (open proxy)
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

