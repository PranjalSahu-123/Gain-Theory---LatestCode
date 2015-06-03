Imports System.Web
Imports System.Web.Services
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Threading
Imports System.Threading.Tasks
Imports System.ComponentModel
Imports System.Xml
Imports System.Net.NetworkInformation
Imports System.Configuration

Module InvokeWebService
    Friend wRequest, request As HttpWebRequest
    Friend wResponse, ws As HttpWebResponse
    Friend stream As Stream
    Friend oStream As Stream
    Friend inputstring, postData As String
    Friend data As Byte()
    Friend bWorkerXMLWS, bwrokerLogWS, bworkerMasterWS As BackgroundWorker
    Friend result As IAsyncResult
    Public Function GetOpXMLFromWS(ByVal inputXML As XElement, ByVal webServiceURL As String) As XElement
        Dim outputXML As XElement
        Try
            wRequest = WebRequest.Create(webServiceURL)
            wRequest.Method = "POST"
            wRequest.ContentType = "application/x-www-form-urlencoded"
            ' wRequest.Timeout = 9000000
            '  wRequest.ServicePoint.MaxIdleTime = 9000000
            wRequest.Timeout = Timeout.Infinite
            wRequest.ServicePoint.MaxIdleTime = Timeout.Infinite
            '  request.KeepAlive = True
            inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(inputXML))

            stream = wRequest.GetRequestStream()
            '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            postData = "inputXML=" + inputstring

            data = encoding.GetBytes(postData)
            'input.Save(stream)
            ' request.ContentLength = data.Length
            stream.Write(data, 0, data.Length)

            ' request.Proxy = Nothing
            ' wResponse = wRequest.GetResponse()
            result = wRequest.BeginGetResponse((New AsyncCallback(AddressOf FinishedXMLWS)), Nothing)
            oStream = wResponse.GetResponseStream()
            '  Dim xmlReader As XmlReader = xmlReader.Create(oStream)
            'outputXML = XElement.Load(xmlReader)
            ' oStream = wResponse.GetResponseStream()
            Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

            '' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New StreamReader(oStream, encode)
            ''Dim xmldoc As XmlDocument = New XmlDocument()
            ''xmldoc.Load(readStream)
            ''xmldoc.Save("C:\\ASR\rnflaargeopxml.xml")
            ''outputXML = XElement.Load(xmldoc.OuterXml)
            Dim str As String = readStream.ReadToEnd()
            '' Dim reader As XmlReader
            ''  Dim reader As XmlReader = XmlReader.Create(readStream)
            ''   outputXML = XElement.ReadFrom(reader)
            outputXML = XElement.Parse(str, Xml.Linq.LoadOptions.None)
            readStream.Close()

        Catch ex As Exception
            LogMpsrintExException("Exception occured while retreiving output from Web service. Message : " + ex.Message)
        Finally
            If Not (wResponse Is Nothing) Then
                wResponse.Close()
            End If
            If Not (oStream Is Nothing) Then
                oStream.Close()
            End If
            If Not (stream Is Nothing) Then
                stream.Close()
            End If

        End Try
        Return outputXML
    End Function
    Public Function CreateTempTable(ByVal inputXML As XElement, ByVal webServiceURL As String) As String
        Dim outputLog As String = String.Empty
        Try

            wRequest = WebRequest.Create(webServiceURL)
            wRequest.Method = "POST"
            wRequest.ContentType = "application/x-www-form-urlencoded"
            'wRequest.Timeout = 900000
            'wRequest.ServicePoint.MaxIdleTime = 900000
            wRequest.Timeout = Timeout.Infinite
            wRequest.ServicePoint.MaxIdleTime = Timeout.Infinite
            '  request.KeepAlive = True
            inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(inputXML))

            stream = wRequest.GetRequestStream()
            '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            postData = "inputXML=" + inputstring

            data = encoding.GetBytes(postData)
            'input.Save(stream)
            ' request.ContentLength = data.Length
            stream.Write(data, 0, data.Length)

            ' request.Proxy = Nothing
            wResponse = wRequest.GetResponse()
            oStream = wResponse.GetResponseStream()
            Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New StreamReader(oStream, encode)
            ' outputXML = XElement.Parse(readStream.ReadToEnd, Xml.Linq.LoadOptions.None)
            outputLog = readStream.ReadToEnd()
            readStream.Close()
            '  result = request.BeginGetResponse((New AsyncCallback(AddressOf FinishedLogWS)), Nothing)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while invoking CreateTempTable webservice.Message :" + ex.Message)
        Finally
            If Not (wResponse Is Nothing) Then
                wResponse.Close()
            End If
            If Not (oStream Is Nothing) Then
                oStream.Close()
            End If
            If Not (stream Is Nothing) Then
                stream.Close()
            End If

        End Try
        Return outputLog
    End Function
    Public Function GetOpForLogWS(ByVal inputXML As XElement, ByVal webServiceURL As String) As String
        Dim outputLog As String = String.Empty
        Try

            wRequest = WebRequest.Create(webServiceURL)
            wRequest.Method = "POST"
            wRequest.ContentType = "application/x-www-form-urlencoded"
            'wRequest.Timeout = 900000
            'wRequest.ServicePoint.MaxIdleTime = 900000
            wRequest.Timeout = Timeout.Infinite
            wRequest.ServicePoint.MaxIdleTime = Timeout.Infinite
            '  request.KeepAlive = True
            inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(inputXML))

            stream = wRequest.GetRequestStream()
            '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            postData = "inputXML=" + inputstring

            data = encoding.GetBytes(postData)
            'input.Save(stream)
            ' request.ContentLength = data.Length
            stream.Write(data, 0, data.Length)

            ' request.Proxy = Nothing
            wResponse = wRequest.GetResponse()
            oStream = wResponse.GetResponseStream()
            Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New StreamReader(oStream, encode)
            ' outputXML = XElement.Parse(readStream.ReadToEnd, Xml.Linq.LoadOptions.None)
            outputLog = readStream.ReadToEnd()
            readStream.Close()
            '  result = request.BeginGetResponse((New AsyncCallback(AddressOf FinishedLogWS)), Nothing)
          
        Catch ex As Exception
        Finally
            If Not (wResponse Is Nothing) Then
                wResponse.Close()
            End If
            If Not (oStream Is Nothing) Then
                oStream.Close()
            End If
            If Not (stream Is Nothing) Then
                stream.Close()
            End If

        End Try
        Return outputLog
    End Function
    Public Function MachineConnectedToInternet() As Boolean
        Dim url As New System.Uri("http://www.google.com/")
        Dim req As System.Net.WebRequest
        req = System.Net.WebRequest.Create(url)
        Dim resp As System.Net.WebResponse
        Try
            resp = req.GetResponse()
            resp.Close()
            req = Nothing
            Return True
        Catch ex As Exception
            req = Nothing
            Return False
            LogMpsrintExException("Internet Connectivity Check failed." + ex.Message)

        End Try
    End Function
    Public Function FormatFileSize(ByVal Size As Long) As String
        Try
            Dim KB As Integer = 1024
            Dim MB As Integer = KB * KB
            Dim GB As Integer = MB * 1024
            ' Return size of file in kilobytes.
            If Size < KB Then
                Return (Size.ToString("D") & " bytes")
            Else
                Select Case Size / KB
                    Case Is < 100
                        Return (Size / KB).ToString("N") & "KB"
                    Case Is < 1000000
                        Return (Size / MB).ToString("N") & "MB"
                    Case Is < 10000000
                        Return (Size / MB / KB).ToString("N") & "GB"
                    Case Is < 10000000
                        Return (Size / GB / MB / KB).ToString("N") & "TB"
                    Case Else
                        Return Size.ToString & "bytes"

                End Select
            End If
        Catch ex As Exception
            Return Size.ToString
        End Try
    End Function
    Public Function GetInternetSpeed() As String
        Try
            Dim myNA() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces

            'Me.Text = myNA(0).Speed
        Catch ex As Exception

        End Try
       

    End Function
    Public Function GetVariantMasterDetails() As XElement
        Dim bslmaster As XElement = New XElement("Master")
        Try
            '
            request = Net.HttpWebRequest.Create("http://54.255.217.55:8080/GroupM/brandlistingdetails/getdetails")
            '
            Dim strr As String = String.Empty
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"
            'request.Timeout = 900000
            'request.ServicePoint.MaxIdleTime = 900000
            request.Timeout = Timeout.Infinite
            request.ServicePoint.MaxIdleTime = Timeout.Infinite
            '  request.KeepAlive = True
            '   InputString = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(Input))

            '  stream = request.GetRequestStream()
            '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            '  Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            ' postData = "inputXML=" + InputString()

            ' Data = encoding.GetBytes(postData)
            'input.Save(stream)
            ' request.ContentLength = data.Length
            'stream.Write(Data, 0, Data.Length)

            ' request.Proxy = Nothing
            ws = request.GetResponse()
            ' result = request.BeginGetResponse((New AsyncCallback(AddressOf FinishedMasterWS)), Nothing)
            oStream = ws.GetResponseStream()
            ' oStream = ws.GetResponseStream()
            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New IO.StreamReader(oStream, encode)
            '   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
            bslmaster = XElement.Parse(readStream.ReadToEnd())
            PopulateVariantMaster(bslmaster)
            ' Dim Globals.Ribbons.MSprintExRibbon.dtchannels As DataTable = New DataTable("Channels")

       
            '   Dim channels As XElement = bslmaster.Element("channel")
            'Parallel.ForEach(channels.Elements.AsEnumerable(), Sub(channel)


            '                                                       'clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                       ' The more computational work you do here, the greater  
            '                                                       ' the speedup compared to a sequential foreach loop. 
            '                                                       'Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                       'Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                       'bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                       'bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                       '' Peek behind the scenes to see how work is parallelized. 
            '                                                       '' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                       'Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                       'close lambda expression and method invocation 
            '                                                   End Sub)

            'For Each channel As XElement In channels.Elements
            '    Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            'Next

            ''genres
            ''Dim Globals.Ribbons.MSprintExRibbon.dtGenres As DataTable = New DataTable("Genres")
            'Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("ChannelCode")
            'Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("Name")
            'Dim keys As DataColumn() = New DataColumn(1) {}
            'keys(0) = Globals.Ribbons.MSprintExRibbon.dtGenres.Columns(0)
            'Globals.Ribbons.MSprintExRibbon.dtGenres.PrimaryKey = keys
            '' Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("Name")
            'Dim genres As XElement = bslmaster.Element("genre")
            'For Each genre As XElement In genres.Elements

            '    'If Not (Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Contains(genre.Attribute("genre_name").Value)) Then
            '    Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Add(genre.Attribute("channel_code").Value, genre.Attribute("genre_name").Value)
            '    '  End If
            'Next
            ''Parallel.ForEach(genres.Elements.AsEnumerable(), Sub(genre)
            ''                                                     ' Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            ''                                                     If Not (Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Contains(genre.Attribute("genre_name").Value)) Then
            ''                                                         Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Add(genre.Attribute("genre_name").Value)
            ''                                                     End If
            ''                                                     'clbSelectGenres.Items.Add(dr(0).ToString())
            ''                                                     ' The more computational work you do here, the greater  
            ''                                                     ' the speedup compared to a sequential foreach loop. 
            ''                                                     'Dim filename As String = System.IO.Path.GetFileName(currentFile)
            ''                                                     'Dim bitmap As New System.Drawing.Bitmap(currentFile)

            ''                                                     'bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            ''                                                     'bitmap.Save(System.IO.Path.Combine(newDir, filename))

            ''                                                     '' Peek behind the scenes to see how work is parallelized. 
            ''                                                     '' But be aware: Thread contention for the Console slows down parallel loops!!!

            ''                                                     'Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            ''                                                     'close lambda expression and method invocation 
            ''                                                 End Sub)
            ''markets
            '' Dim Globals.Ribbons.MSprintExRibbon.dtMarkets1 As DataTable = New DataTable("Markets")
            'Globals.Ribbons.MSprintExRibbon.dtMarkets1.Columns.Add("ID")
            'Globals.Ribbons.MSprintExRibbon.dtMarkets1.Columns.Add("Name")
            'Dim markets As XElement = bslmaster.Element("market")
            'For Each market As XElement In markets.Elements
            '    Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows.Add(market.Attribute("code").Value, market.Attribute("name").Value)
            'Next
            'Parallel.ForEach(markets.Elements.AsEnumerable(), Sub(market)
            '                                                       Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            '                                                     Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows.Add(market.Attribute("code").Value, market.Attribute("name").Value)
            '                                                      clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                       The more computational work you do here, the greater  
            '                                                       the speedup compared to a sequential foreach loop. 
            '                                                      Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                      Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                      bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                      bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                      ' Peek behind the scenes to see how work is parallelized. 
            '                                                      ' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                      Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                      close lambda expression and method invocation 
            '      
        Catch webex As WebException
            LogMpsrintExException("Exception occured while parsing Master WS response." + webex.Message)
            Throw webex
        Catch ex As Exception
            LogMpsrintExException("Exception occured while communicating with Master Webservice")
            Throw ex
        End Try
    End Function
    Public Function PopulateVariantMaster(ByVal master As XElement) As Data.DataTable
        Try
            Globals.Ribbons.MSprintExRibbon.bslMasterTable = New Data.DataTable()
            Globals.Ribbons.MSprintExRibbon.bslMasterTable.Columns.Add("Variant")
            Globals.Ribbons.MSprintExRibbon.bslMasterTable.Columns.Add("Category")
            Globals.Ribbons.MSprintExRibbon.bslMasterTable.Columns.Add("Brand")
            Globals.Ribbons.MSprintExRibbon.bslMasterTable.Columns.Add("Advertiser")
            For Each blisting As XElement In master.Elements
                Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.bslMasterTable.NewRow()
                row("Variant") = blisting.Attribute("variant").Value
                row("Category") = blisting.Attribute("category").Value
                row("Brand") = blisting.Attribute("brand").Value
                row("Advertiser") = blisting.Attribute("advertiser").Value
                Globals.Ribbons.MSprintExRibbon.bslMasterTable.Rows.Add(row)
            Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while populating datatable from variant XML.Message :" + ex.Message)
        End Try
        Return Globals.Ribbons.MSprintExRibbon.bslMasterTable
    End Function
    Public Function GetTempTableList(ByVal URL As String, ByVal requestType As String) As XElement
        Dim latestWeek As XElement = New XElement("mediaplan")
        Try
            ' request = Net.HttpWebRequest.Create("http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/get_status/")
            ' ConfigurationManager.AppSettings("CGMMasterWSURL")
            '  request = Net.HttpWebRequest.Create(Globals.Ribbons.MSprintExRibbon.GetURLForWS("LatestWeekWSURL"))
            request = Net.HttpWebRequest.Create(URL)
            Dim strr As String = String.Empty
            request.Method = requestType
            request.ContentType = "application/x-www-form-urlencoded"
            'request.Timeout = 900000
            'request.ServicePoint.MaxIdleTime = 900000
            request.Timeout = Timeout.Infinite
            request.ServicePoint.MaxIdleTime = Timeout.Infinite
            '  request.KeepAlive = True
            '   InputString = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(Input))

            '  stream = request.GetRequestStream()
            '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            '  Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            ' postData = "inputXML=" + InputString()

            ' Data = encoding.GetBytes(postData)
            'input.Save(stream)
            ' request.ContentLength = data.Length
            'stream.Write(Data, 0, Data.Length)

            ' request.Proxy = Nothing
            ws = request.GetResponse()
            ' result = request.BeginGetResponse((New AsyncCallback(AddressOf FinishedMasterWS)), Nothing)
            oStream = ws.GetResponseStream()
            ' oStream = ws.GetResponseStream()
            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New IO.StreamReader(oStream, encode)
            '   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
            latestWeek = XElement.Parse(readStream.ReadToEnd())

        Catch ex As Exception
            LogMpsrintExException("Exception occured while communicating with GetTempTableList Webservice." + ex.Message)
            Throw ex
        Finally
            ws.Close()
        End Try
        Return latestWeek
    End Function
    Public Function GetLatestWeekDetails(ByVal URL As String, ByVal requestType As String) As XElement
        Dim latestWeek As XElement = New XElement("latest")
        Try
            ' request = Net.HttpWebRequest.Create("http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/get_status/")
            ' ConfigurationManager.AppSettings("CGMMasterWSURL")
            '  request = Net.HttpWebRequest.Create(Globals.Ribbons.MSprintExRibbon.GetURLForWS("LatestWeekWSURL"))
            request = Net.HttpWebRequest.Create(URL)
            Dim strr As String = String.Empty
            request.Method = requestType
            request.ContentType = "application/x-www-form-urlencoded"
            'request.Timeout = 900000
            'request.ServicePoint.MaxIdleTime = 900000
            request.Timeout = Timeout.Infinite
            request.ServicePoint.MaxIdleTime = Timeout.Infinite
            '  request.KeepAlive = True
            '   InputString = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(Input))

            '  stream = request.GetRequestStream()
            '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            '  Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            ' postData = "inputXML=" + InputString()

            ' Data = encoding.GetBytes(postData)
            'input.Save(stream)
            ' request.ContentLength = data.Length
            'stream.Write(Data, 0, Data.Length)

            ' request.Proxy = Nothing
            ws = request.GetResponse()
            ' result = request.BeginGetResponse((New AsyncCallback(AddressOf FinishedMasterWS)), Nothing)
            oStream = ws.GetResponseStream()
            ' oStream = ws.GetResponseStream()
            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New IO.StreamReader(oStream, encode)
            '   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
            latestWeek = XElement.Parse(readStream.ReadToEnd())

        Catch ex As Exception
            LogMpsrintExException("Exception occured while communicating with GetlatestStatus Webservice")
            Throw ex
        Finally
            ws.Close()
        End Try
        Return latestWeek
    End Function
    Public Function PopulateTabsFromMasterWs() As XElement
        Dim master As XElement = New XElement("Master")
        Try
            '
            '  request = Net.HttpWebRequest.Create("http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/master/")
            '
            request = Net.HttpWebRequest.Create(Globals.Ribbons.MSprintExRibbon.GetURLForWS("CGMMasterWSURL_New"))
            Dim strr As String = String.Empty
            request.Method = "GET"
            request.ContentType = "application/x-www-form-urlencoded"
            request.Timeout = Timeout.Infinite
            request.ServicePoint.MaxIdleTime = Timeout.Infinite

            '  request.KeepAlive = True
            '   InputString = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(Input))

            '  stream = request.GetRequestStream()
            '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            '  Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            ' postData = "inputXML=" + InputString()

            ' Data = encoding.GetBytes(postData)
            'input.Save(stream)
            ' request.ContentLength = data.Length
            'stream.Write(Data, 0, Data.Length)

            ' request.Proxy = Nothing
            ws = request.GetResponse()
            ' result = request.BeginGetResponse((New AsyncCallback(AddressOf FinishedMasterWS)), Nothing)
            oStream = ws.GetResponseStream()
            ' oStream = ws.GetResponseStream()
            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New IO.StreamReader(oStream, encode)
            '   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
            master = XElement.Parse(readStream.ReadToEnd())
            ' Dim Globals.Ribbons.MSprintExRibbon.dtchannels As DataTable = New DataTable("Channels")

            Globals.Ribbons.MSprintExRibbon.dtchannels.Columns.Add("ID", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.dtchannels.Columns.Add("Name")
            Dim channels As XElement = master.Element("channel")
            'Parallel.ForEach(channels.Elements.AsEnumerable(), Sub(channel)


            '                                                       'clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                       ' The more computational work you do here, the greater  
            '                                                       ' the speedup compared to a sequential foreach loop. 
            '                                                       'Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                       'Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                       'bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                       'bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                       '' Peek behind the scenes to see how work is parallelized. 
            '                                                       '' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                       'Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                       'close lambda expression and method invocation 
            '                                                   End Sub)

            For Each channel As XElement In channels.Elements
                Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            Next

            'genres
            'Dim Globals.Ribbons.MSprintExRibbon.dtGenres As DataTable = New DataTable("Genres")
            Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("ChannelCode")
            Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("Name")
            Dim keys As DataColumn() = New DataColumn(1) {}
            keys(0) = Globals.Ribbons.MSprintExRibbon.dtGenres.Columns(0)
            Globals.Ribbons.MSprintExRibbon.dtGenres.PrimaryKey = keys
            ' Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("Name")
            Dim genres As XElement = master.Element("genre")
            For Each genre As XElement In genres.Elements

                'If Not (Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Contains(genre.Attribute("genre_name").Value)) Then
                Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Add(genre.Attribute("channel_code").Value, genre.Attribute("genre_name").Value)
                '  End If
            Next
            'Parallel.ForEach(genres.Elements.AsEnumerable(), Sub(genre)
            '                                                     ' Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            '                                                     If Not (Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Contains(genre.Attribute("genre_name").Value)) Then
            '                                                         Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Add(genre.Attribute("genre_name").Value)
            '                                                     End If
            '                                                     'clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                     ' The more computational work you do here, the greater  
            '                                                     ' the speedup compared to a sequential foreach loop. 
            '                                                     'Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                     'Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                     'bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                     'bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                     '' Peek behind the scenes to see how work is parallelized. 
            '                                                     '' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                     'Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                     'close lambda expression and method invocation 
            '                                                 End Sub)
            'markets
            ' Dim Globals.Ribbons.MSprintExRibbon.dtMarkets1 As DataTable = New DataTable("Markets")
            Globals.Ribbons.MSprintExRibbon.dtMarkets1.Columns.Add("ID")
            Globals.Ribbons.MSprintExRibbon.dtMarkets1.Columns.Add("Name")
            Dim markets As XElement = master.Element("market")
            For Each market As XElement In markets.Elements
                Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows.Add(market.Attribute("code").Value, market.Attribute("name").Value)
            Next
            'Parallel.ForEach(markets.Elements.AsEnumerable(), Sub(market)
            '                                                       Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            '                                                     Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows.Add(market.Attribute("code").Value, market.Attribute("name").Value)
            '                                                      clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                       The more computational work you do here, the greater  
            '                                                       the speedup compared to a sequential foreach loop. 
            '                                                      Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                      Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                      bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                      bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                      ' Peek behind the scenes to see how work is parallelized. 
            '                                                      ' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                      Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                      close lambda expression and method invocation 
            '      
        Catch webex As WebException
            LogMpsrintExException("Exception occured while parsing Master WS response." + webex.Message)
            Throw webex
        Catch ex As Exception
            LogMpsrintExException("Exception occured while communicating with Master Webservice")
            Throw ex
        End Try
    End Function
    Public Sub FinishedLogWS(ByVal result As IAsyncResult)
        Dim outputLog As String = String.Empty
        Try
            ws = request.EndGetResponse(result)
            oStream = ws.GetResponseStream()
            ' oStream = wResponse.GetResponseStream()
            Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New StreamReader(oStream, encode)
            ' outputXML = XElement.Parse(readStream.ReadToEnd, Xml.Linq.LoadOptions.None)
            outputLog = readStream.ReadToEnd()
            readStream.Close()
        Catch ex As Exception

        End Try
        Return
    End Sub
    Public Sub FinishedXMLWS(ByVal result As IAsyncResult)
        Dim outputXML As XElement = New XElement("input")
        Try
            wResponse = wRequest.EndGetResponse(result)
           
        Catch ex As Exception

        End Try
        Return
    End Sub
    Public Sub FinishedMasterWS(ByVal result As IAsyncResult)
        Dim master As XElement = New XElement("Master")
        Try
            ws = request.EndGetResponse(result)
            oStream = ws.GetResponseStream()
            ' oStream = ws.GetResponseStream()
            Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

            ' Pipe the stream to a higher level stream reader with the required encoding format.
            Dim readStream As New IO.StreamReader(oStream, encode)
            '   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
            master = XElement.Parse(readStream.ReadToEnd())
            ' Dim Globals.Ribbons.MSprintExRibbon.dtchannels As DataTable = New DataTable("Channels")

            Globals.Ribbons.MSprintExRibbon.dtchannels.Columns.Add("ID", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.dtchannels.Columns.Add("Name")
            Dim channels As XElement = master.Element("channel")
            'Parallel.ForEach(channels.Elements.AsEnumerable(), Sub(channel)


            '                                                       'clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                       ' The more computational work you do here, the greater  
            '                                                       ' the speedup compared to a sequential foreach loop. 
            '                                                       'Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                       'Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                       'bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                       'bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                       '' Peek behind the scenes to see how work is parallelized. 
            '                                                       '' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                       'Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                       'close lambda expression and method invocation 
            '                                                   End Sub)

            For Each channel As XElement In channels.Elements
                Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            Next

            'genres
            'Dim Globals.Ribbons.MSprintExRibbon.dtGenres As DataTable = New DataTable("Genres")
            Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("ChannelCode")
            Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("Name")
            Dim keys As DataColumn() = New DataColumn(1) {}
            keys(0) = Globals.Ribbons.MSprintExRibbon.dtGenres.Columns(0)
            Globals.Ribbons.MSprintExRibbon.dtGenres.PrimaryKey = keys
            ' Globals.Ribbons.MSprintExRibbon.dtGenres.Columns.Add("Name")
            Dim genres As XElement = master.Element("genre")
            For Each genre As XElement In genres.Elements

                'If Not (Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Contains(genre.Attribute("genre_name").Value)) Then
                Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Add(genre.Attribute("channel_code").Value, genre.Attribute("genre_name").Value)
                '  End If

            Next
            'Parallel.ForEach(genres.Elements.AsEnumerable(), Sub(genre)
            '                                                     ' Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            '                                                     If Not (Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Contains(genre.Attribute("genre_name").Value)) Then
            '                                                         Globals.Ribbons.MSprintExRibbon.dtGenres.Rows.Add(genre.Attribute("genre_name").Value)
            '                                                     End If
            '                                                     'clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                     ' The more computational work you do here, the greater  
            '                                                     ' the speedup compared to a sequential foreach loop. 
            '                                                     'Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                     'Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                     'bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                     'bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                     '' Peek behind the scenes to see how work is parallelized. 
            '                                                     '' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                     'Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                     'close lambda expression and method invocation 
            '                                                 End Sub)
            'markets
            ' Dim Globals.Ribbons.MSprintExRibbon.dtMarkets1 As DataTable = New DataTable("Markets")
            Globals.Ribbons.MSprintExRibbon.dtMarkets1.Columns.Add("ID")
            Globals.Ribbons.MSprintExRibbon.dtMarkets1.Columns.Add("Name")
            Dim markets As XElement = master.Element("market")
            For Each market As XElement In markets.Elements
                Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows.Add(market.Attribute("code").Value, market.Attribute("name").Value)
            Next
            'Parallel.ForEach(markets.Elements.AsEnumerable(), Sub(market)
            '                                                       Globals.Ribbons.MSprintExRibbon.dtchannels.Rows.Add(channel.Attribute("code").Value, channel.Attribute("name").Value)
            '                                                     Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows.Add(market.Attribute("code").Value, market.Attribute("name").Value)
            '                                                      clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                       The more computational work you do here, the greater  
            '                                                       the speedup compared to a sequential foreach loop. 
            '                                                      Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                      Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                      bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                      bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                      ' Peek behind the scenes to see how work is parallelized. 
            '                                                      ' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                      Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                      close lambda expression and method invocation 
            '      
        Catch webex As WebException
            LogMpsrintExException("Exception occured while parsing Master WS response." + webex.Message)
            Throw webex
        Catch ex As Exception
            LogMpsrintExException("Exception occured while parsing Master WS response." + ex.Message)
            Throw ex
        End Try

        Return
    End Sub
End Module
