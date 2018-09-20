Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Math
Imports System.WeakReference
Imports System.Xml
Imports System.Text
Module Procesa_XML
    Dim path As String = "\\server-nas\CFDI_CP_Factoraje\"
    'Dim path As String = "C:\Users\vicente-apolonio\Desktop\temp\"
    Sub Main()
        Dim dtF100 As New XML_FactorajeDSTableAdapters.Vw_ChequesDetalleTableAdapter
        Dim dtWebXML As New XML_FactorajeDSTableAdapters.WEB_FacturasXMLTableAdapter
        Dim D As System.IO.DirectoryInfo
        D = New System.IO.DirectoryInfo(path)

        For Each Archivo As String In My.Computer.FileSystem.GetFiles(
                                path,
                                FileIO.SearchOption.SearchTopLevelOnly,
                                "*.xml")

            Dim nombre() As String = My.Computer.FileSystem.GetName(Archivo).Split(".")


            Dim cadXML As String
            Dim cadena As StreamReader
            cadena = New StreamReader(Archivo)
            cadXML = cadena.ReadToEnd
            cadena.Close()

            Dim var As String = ""
            Dim folio As String = leeXMLF(cadXML, "Folio")
            Dim serie As String = leeXMLF(cadXML, "Serie")
            Dim rfc As String = leeXMLF(cadXML, "RFCR")
            Dim importe As String = leeXMLF(cadXML, "Total")
            Dim ffactura As String = CDate(leeXMLF(cadXML, "Fecha")).ToShortDateString
            Dim uuid As String = leeXMLF(cadXML, "UUID")
            Dim tcambio As String = leeXMLF(cadXML, "TipoCambio")
            Dim mpago As String = leeXMLF(cadXML, "MetodoPago")
            Dim moneda As String = leeXMLF(cadXML, "Moneda")
            Dim tcomprobante As String = leeXMLF(cadXML, "TipoDeComprobante")



            If tcomprobante = "I" Then
                'Validación de información
                If rfc = "DME061031H27" Or rfc = "CVN140812CQ9" Or rfc = "DIM061230LN8" Then
                    var = dtF100.Existe_Solo_Folio_ScalarQuery(rfc, importe, folio)
                Else
                    var = dtF100.Existe_ScalarQuery(rfc, serie & folio, importe)
                End If
                Dim var_xml As String = dtWebXML.Existe_ScalarQuery(uuid)

                If moneda <> "MXN" Then
                    var = dtF100.Existe_DifMXN_ScalarQuery(rfc, folio)
                    tcambio = CDbl(var) / CDbl(importe)
                End If
                If var <> "0" Then
                    If var_xml = "NE" Then
                        If rfc = "DME061031H27" Or rfc = "CVN140812CQ9" Or rfc = "DIM061230LN8" Then
                            dtWebXML.Insert(folio, folio, rfc, CDbl(importe), 0, ffactura, Nothing, False, Nothing, Nothing, uuid, "", CInt(folio), tcambio, mpago, moneda)
                        Else
                            dtWebXML.Insert(folio, folio, rfc, CDbl(importe), 0, ffactura, Nothing, False, Nothing, Nothing, uuid, serie, CInt(folio), tcambio, mpago, moneda)
                        End If
                    End If
                End If

                System.IO.File.Copy(Archivo, path & "I_Procesados\" & uuid & ".xml", True)
                'System.IO.File.Copy(path & nombre(0) & ".pdf", path & "I_Procesados\" & uuid & ".pdf", True)
            ElseIf tcomprobante = "P" Then
                If System.IO.File.Exists(path & nombre(0) & ".pdf") Then
                    envia_mail("REDCOFIDI|DIVISION:|CODIGO:", nombre(0))
                    System.IO.File.Copy(Archivo, path & "P_Procesados\" & uuid & ".xml", True)
                    System.IO.File.Copy(path & nombre(0) & ".pdf", path & "P_Procesados\" & uuid & ".pdf", True)
                End If
            Else
                cadena.Close()
                System.IO.File.Copy(Archivo, path & "O_Procesados\" & uuid & ".xml", True)
                System.IO.File.Copy(path & nombre(0) & ".pdf", path & "O_Procesados\" & uuid & ".pdf", True)
            End If
            File.Delete(Archivo)
            File.Delete(path & nombre(0) & ".pdf")
        Next
    End Sub

    Public Sub envia_mail(asunto As String, att_archivo As String)
        Dim Servidor As New Mail.SmtpClient
        Dim Mensaje As Mail.MailMessage
        Dim Adjunto1 As Mail.Attachment
        Dim Adjunto2 As Mail.Attachment

        Servidor.Host = "smtp01.cmoderna.com"
        Servidor.Port = "26"
        Try
            Mensaje = New Mail.MailMessage
            Mensaje.IsBodyHtml = True
            Mensaje.From = New Mail.MailAddress("jdelgado@finagil.com.mx")
            Mensaje.To.Add("red.cofidi.inbox@ateb.com.mx")
            Mensaje.To.Add("jdelgado@finagil.com.mx")

            Mensaje.Subject = asunto
            Adjunto1 = New Mail.Attachment(path & att_archivo & ".xml")
            Adjunto2 = New Mail.Attachment(path & att_archivo & ".pdf")

            Mensaje.Attachments.Add(Adjunto1)
            Mensaje.Attachments.Add(Adjunto2)

            Servidor.Send(Mensaje)

            Adjunto1.Dispose()
            Adjunto2.Dispose()

        Catch ex As Exception
        End Try
    End Sub

    Public Function leeXMLF(docXML As String, nodo As String)

        Dim doc As XmlDataDocument
        doc = New XmlDataDocument
        doc.LoadXml(docXML)
        Dim CFDI As XmlNode
        Dim retorno As String = ""

        CFDI = doc.DocumentElement
        If nodo = "RFCR" Or nodo = "RFCE" Then
            For Each comprobante As XmlNode In CFDI.ChildNodes
                If comprobante.Name = "cfdi:Receptor" And nodo = "RFCR" Then
                    For Each receptor As XmlNode In comprobante.Attributes
                        If receptor.Name = "Rfc" Then
                            retorno = receptor.Value.ToString
                            Return retorno
                        End If
                    Next
                End If
                If comprobante.Name = "cfdi:Emisor" And nodo = "RFCE" Then
                    For Each receptor As XmlNode In comprobante.Attributes
                        If receptor.Name = "Rfc" Then
                            retorno = receptor.Value.ToString
                            Return retorno
                        End If
                    Next
                End If
            Next
        End If

        If nodo <> "UUID" Then
            For Each Comprobante As XmlNode In CFDI.Attributes
                If Comprobante.Name = "Moneda" And nodo = "Moneda" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf Comprobante.Name = "TipoCambio" And nodo = "TipoCambio" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf Comprobante.Name = "TipoDeComprobante" And nodo = "TipoDeComprobante" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf (Comprobante.Name = "Total" Or Comprobante.Name = "total") And nodo = "Total" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf (Comprobante.Name = "MetodoPago" Or Comprobante.Name = "metodoDePago") And nodo = "MetodoPago" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf Comprobante.Name = "FormaPago" And nodo = "FormaPago" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf (Comprobante.Name = "Serie" Or Comprobante.Name = "serie") And nodo = "Serie" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf (Comprobante.Name = "Folio" Or Comprobante.Name = "folio") And nodo = "Folio" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                ElseIf Comprobante.Name = "Fecha" And nodo = "Fecha" Then
                    retorno = Comprobante.Value.ToString
                    Return retorno
                    Exit Function
                End If
            Next
        Else
            For Each Comprobante As XmlNode In CFDI.ChildNodes
                For Each Complemento As XmlNode In Comprobante
                    If Complemento.Name = "tfd:TimbreFiscalDigital" Then
                        For Each atributos As XmlNode In Complemento.Attributes
                            If atributos.Name = "UUID" Then
                                retorno = atributos.Value.ToString
                                Return retorno
                                Exit Function
                            End If
                        Next
                    End If
                Next
            Next
        End If
    End Function
End Module
