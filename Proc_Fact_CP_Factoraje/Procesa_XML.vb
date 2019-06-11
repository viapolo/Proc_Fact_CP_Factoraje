Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Math
Imports System.WeakReference
Imports System.Xml
Imports System.Text
Imports System
Module Procesa_XML
    Dim path As String = "\\server-nas\CFDI_CP_Factoraje\"
    Dim pathCxpA As String = "\\server-nas\Contabilidad CFDI\ARCHIVOS ADD CONTPAQi\CFDI_PROV\ARFIN\Todos\"
    Dim pathCxpF As String = "\\server-nas\Contabilidad CFDI\ARCHIVOS ADD CONTPAQi\CFDI_PROV\FINAGIL\Todos\"
    'Dim path As String = "C:\Users\vicente-apolonio\Desktop\temp\"
    Sub Main()
        Dim dtF100 As New XML_FactorajeDSTableAdapters.Vw_ChequesDetalleTableAdapter
        Dim dtWebXML As New XML_FactorajeDSTableAdapters.WEB_FacturasXMLTableAdapter
        Dim D As System.IO.DirectoryInfo
        D = New System.IO.DirectoryInfo(path)

        procesaXmlCxpA()
        procesaXmlCxpF()

        For Each Archivo As String In My.Computer.FileSystem.GetFiles(path, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")

            Dim nombre() As String = My.Computer.FileSystem.GetName(Archivo).Split(".")


            Dim cadXML As String
            Dim cadena As StreamReader
            cadena = New StreamReader(Archivo)
            cadXML = cadena.ReadToEnd
            cadena.Close()
            'cadena.Dispose()


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
                If rfc = "DME061031H27" Or rfc = "CVN140812CQ9" Or rfc = "DIM061230LN8" Or rfc = "GTC980421R4A" Then
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
                        If rfc = "DME061031H27" Or rfc = "CVN140812CQ9" Or rfc = "DIM061230LN8" Or rfc = "GTC980421R4A" Then
                            dtWebXML.Insert(folio, folio, rfc, CDbl(importe), 0, ffactura, Nothing, False, Nothing, Nothing, uuid, "", CInt(folio), tcambio, mpago, moneda)
                        Else
                            dtWebXML.Insert(folio, folio, rfc, CDbl(importe), 0, ffactura, Nothing, False, Nothing, Nothing, uuid, serie, CInt(folio), tcambio, mpago, moneda)
                        End If
                    End If
                End If

                System.IO.File.Copy(Archivo, path & "I_Procesados\" & uuid & ".xml", True)
                cadena.Close()
                cadena.Dispose()
                'System.IO.File.Copy(path & nombre(0) & ".pdf", path & "I_Procesados\" & uuid & ".pdf", True)
            ElseIf tcomprobante = "P" Then
                If System.IO.File.Exists(path & nombre(0) & ".pdf") Then
                    envia_mail("REDCOFIDI|DIVISION:|CODIGO:", nombre(0))

                    System.IO.File.Copy(Archivo, path & "P_Procesados\" & uuid & ".xml", True)
                    System.IO.File.Copy(path & nombre(0) & ".pdf", path & "P_Procesados\" & uuid & ".pdf", True)
                    cadena.Close()
                    cadena.Dispose()
                End If
            Else
                System.IO.File.Copy(Archivo, path & "O_Procesados\" & uuid & ".xml", True)
                System.IO.File.Copy(path & nombre(0) & ".pdf", path & "O_Procesados\" & uuid & ".pdf", True)
                cadena.Close()
                cadena.Dispose()

            End If

            cadena.Close()
            cadena.Dispose()

            'File.Delete(Archivo)
            File.Delete(path & nombre(0) & ".xml")
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
            Mensaje.Dispose()
            Servidor.Dispose()

        Catch ex As Exception
            Adjunto1.Dispose()
            Adjunto2.Dispose()
            Mensaje.Dispose()
            Servidor.Dispose()
        End Try
    End Sub

    Public Function leeXMLF(docXML As String, nodo As String)

        Dim doc As XmlDataDocument
        doc = New XmlDataDocument
        doc.LoadXml(docXML)
        Dim CFDI As XmlNode
        Dim retorno As String = ""

        CFDI = doc.DocumentElement
        If nodo = "RFCR" Or nodo = "RFCE" Or nodo = "NombreR" Or nodo = "NombreE" Then
            For Each comprobante As XmlNode In CFDI.ChildNodes
                If comprobante.Name = "cfdi:Receptor" And (nodo = "RFCR" Or nodo = "NombreR") Then
                    For Each receptor As XmlNode In comprobante.Attributes
                        If receptor.Name = "Rfc" Then
                            retorno = receptor.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                        If receptor.Name = "Nombre" And nodo = "NombreR" Then
                            retorno = receptor.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                    Next
                End If
                If comprobante.Name = "cfdi:Emisor" And (nodo = "RFCE" Or nodo = "NombreE") Then
                    For Each receptor As XmlNode In comprobante.Attributes
                        If receptor.Name = "Rfc" And nodo = "RFCE" Then
                            retorno = receptor.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                        If receptor.Name = "Nombre" And nodo = "NombreE" Then
                            retorno = receptor.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If

        If nodo = "TIR" Or nodo = "TIT" Then
            For Each comprobante As XmlNode In CFDI.ChildNodes
                If comprobante.Name = "cfdi:Impuestos" Then
                    For Each comprobanteC As XmlNode In comprobante.Attributes
                        If comprobanteC.Name = "TotalImpuestosTrasladados" And nodo = "TIT" Then
                            retorno = comprobanteC.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                        If comprobanteC.Name = "TotalImpuestosRetenidos" And nodo = "TIR" Then
                            retorno = comprobanteC.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                    Next
                End If
            Next
        End If

        If nodo <> "UUID" And nodo <> "FechaT" Then
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
                    If Comprobante.Value.ToString.Length > 19 Then
                        retorno = (Comprobante.Value.ToString).Substring(0, 20)
                    Else
                        retorno = Comprobante.Value.ToString
                    End If

                    If Comprobante.Value.ToString = "" Then
                        retorno = 0
                    End If
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
                            If atributos.Name = "UUID" And nodo = "UUID" Then
                                retorno = atributos.Value.ToString
                                Return retorno
                                Exit Function
                            End If
                            If atributos.Name = "FechaTimbrado" And nodo = "FechaT" Then
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

    Public Sub procesaXmlCxpA()
        Dim dtCxp As New XML_CXPDSTableAdapters.CXP_XmlCfdi2TableAdapter
        Dim dtProveedores As New XML_CXPDSTableAdapters.CXP_ProveedoresTableAdapter

        Dim D As System.IO.DirectoryInfo
        D = New System.IO.DirectoryInfo(pathCxpA)
        Dim res As readXML_CFDI_class = New readXML_CFDI_class

        For Each Archivo As String In My.Computer.FileSystem.GetFiles(pathCxpA, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")

            Dim nombre() As String = My.Computer.FileSystem.GetName(Archivo).Split(".")

            Dim impLocRet As Decimal
            Dim impLocTra As Decimal
            Dim cadXML As String
            Dim cadena As StreamReader
            cadena = New StreamReader(Archivo)
            cadXML = cadena.ReadToEnd
            cadena.Close()


            Try

                If dtProveedores.ExisteProv_ScalarQuery(res.LeeXML(Archivo, "RFCE")) = "NE" Then
                    dtProveedores.Insert(res.LeeXML(Archivo, "RFCE"), Nothing, Nothing, res.LeeXML(Archivo, "NombreE"), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, System.Data.SqlTypes.SqlDateTime.Null, Nothing, Nothing, Nothing, Nothing, Nothing)
                End If


                If dtCxp.Existe_ScalarQuery(leeXMLF(cadXML, "UUID")).ToString = "NE" Then
                    impLocTra = res.LeeXML(Archivo, "implocalT")
                    impLocRet = res.LeeXML(Archivo, "implocalR")

                    'dtCxp.Insert(leeXMLF(cadXML, "RFCE"), leeXMLF(cadXML, "RFCR"), CDbl(leeXMLF(cadXML, "Total")), CDbl(leeXMLF(cadXML, "TIR")), CDbl(leeXMLF(cadXML, "TIT")), leeXMLF(cadXML, "UUID"), leeXMLF(cadXML, "NombreE"), leeXMLF(cadXML, "Moneda"), leeXMLF(cadXML, "MetodoPago"), leeXMLF(cadXML, "FormaPago"), CDbl(leeXMLF(cadXML, "TipoCambio")), leeXMLF(cadXML, "TipoDeComprobante"), leeXMLF(cadXML, "Serie"), leeXMLF(cadXML, "Folio"), leeXMLF(cadXML, "Fecha"), leeXMLF(cadXML, "FechaT"), False, Date.Now)
                    Dim conceptos As XmlNode = res.LeeXML_Conceptos(Archivo, "Concepto")
                    Dim contDetalle As Integer = 0
                    For Each detalle_conceptos As XmlNode In conceptos.ChildNodes
                        If detalle_conceptos.Name = "cfdi:Concepto" Then
                            If detalle_conceptos.ChildNodes.Count = 0 Then
                                contDetalle += 1
                                dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")), Nothing, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", CDec(res.LeeXML(Archivo, "Total")), contDetalle.ToString, Nothing, Nothing, impLocRet, impLocTra, Nothing)
                            End If
                            For Each concepto_hijos As XmlNode In detalle_conceptos.ChildNodes
                                If concepto_hijos.Name = "cfdi:Impuestos" Then
                                    For Each impuestos_detalle As XmlNode In concepto_hijos.ChildNodes
                                        If impuestos_detalle.Name = "cfdi:Traslados" Then
                                            Dim Base As String = ""
                                            Dim Impuesto As String = ""
                                            Dim Tipofactor As String = ""
                                            Dim TasaOCuota As String = "0"
                                            Dim ImporteImpuesto As String = "0"
                                            For Each impuestos_traslado As XmlNode In impuestos_detalle.ChildNodes
                                                If impuestos_traslado.Name = "cfdi:Traslado" Then
                                                    For Each impuestos_traslado_atributos As XmlNode In impuestos_traslado.Attributes
                                                        If impuestos_traslado_atributos.Name = "Base" Then
                                                            Base = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Impuesto" Then
                                                            Impuesto = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TipoFactor" Then
                                                            Tipofactor = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TasaOCuota" Then
                                                            TasaOCuota = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Importe" Then
                                                            ImporteImpuesto = impuestos_traslado_atributos.Value.ToString
                                                        End If
                                                    Next
                                                    'Insert
                                                    contDetalle += 1
                                                    dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")), Impuesto, CDec(ImporteImpuesto), res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", CDec(res.LeeXML(Archivo, "Total")), contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, Nothing)
                                                End If
                                            Next
                                        End If
                                        If impuestos_detalle.Name = "cfdi:Retenciones" Then
                                            Dim Base As String = ""
                                            Dim Impuesto As String = ""
                                            Dim Tipofactor As String = ""
                                            Dim TasaOCuota As String = ""
                                            Dim ImporteImpuesto As String = ""
                                            For Each impuestos_traslado As XmlNode In impuestos_detalle.ChildNodes
                                                If impuestos_traslado.Name = "cfdi:Retencion" Then
                                                    For Each impuestos_traslado_atributos As XmlNode In impuestos_traslado.Attributes
                                                        If impuestos_traslado_atributos.Name = "Base" Then
                                                            Base = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Impuesto" Then
                                                            Impuesto = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TipoFactor" Then
                                                            Tipofactor = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TasaOCuota" Then
                                                            TasaOCuota = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Importe" Then
                                                            ImporteImpuesto = impuestos_traslado_atributos.Value.ToString
                                                        End If
                                                    Next
                                                    'Insert
                                                    contDetalle += 1
                                                    dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")), Impuesto, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", CDec(res.LeeXML(Archivo, "Total")), contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, CDec(ImporteImpuesto))
                                                End If
                                            Next

                                        End If
                                    Next
                                End If


                            Next
                        End If
                    Next

                    'dtCxp.Insert(res.LeeXML(cadXML, "RFCE"), res.LeeXML(cadXML, "RFCR"), res.LeeXML(cadXML, "SubTotal"), res.LeeXML(cadXML, "SubTotal"))
                    'System.IO.File.Move(Archivo, pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    'System.IO.File.Move(pathCxpA & nombre(0) & ".pdf", pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                    'WriteLine("Se insertó el UUID: " & leeXMLF(cadXML, "UUID").ToString)
                    System.IO.File.Move(Archivo, pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    System.IO.File.Move(pathCxpA & nombre(0) & ".pdf", pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                    'Else
                    '    System.IO.File.Move(Archivo, pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    '    System.IO.File.Move(pathCxpA & nombre(0) & ".pdf", pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                End If
            Catch ex As Exception
                'WriteLine(ex.ToString)
            End Try
            File.Delete(Archivo)
            File.Delete(pathCxpA & nombre(0) & ".pdf")
        Next
    End Sub

    Public Sub procesaXmlCxpF()
        Dim dtCxp As New XML_CXPDSTableAdapters.CXP_XmlCfdi2TableAdapter
        Dim dtProveedores As New XML_CXPDSTableAdapters.CXP_ProveedoresTableAdapter

        Dim D As System.IO.DirectoryInfo
        D = New System.IO.DirectoryInfo(pathCxpF)
        Dim res As readXML_CFDI_class = New readXML_CFDI_class

        For Each Archivo As String In My.Computer.FileSystem.GetFiles(
                                pathCxpF,
                                FileIO.SearchOption.SearchTopLevelOnly,
                                "*.xml")

            Dim nombre() As String = My.Computer.FileSystem.GetName(Archivo).Split(".")

            Dim impLocRet As Decimal
            Dim impLocTra As Decimal
            Dim cadXML As String
            Dim cadena As StreamReader
            cadena = New StreamReader(Archivo)
            cadXML = cadena.ReadToEnd
            cadena.Close()

            Try

                If dtProveedores.ExisteProv_ScalarQuery(res.LeeXML(Archivo, "RFCE")) = "NE" Then
                    dtProveedores.Insert(res.LeeXML(Archivo, "RFCE"), Nothing, Nothing, res.LeeXML(Archivo, "NombreE"), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, System.Data.SqlTypes.SqlDateTime.Null, Nothing, Nothing, Nothing, Nothing, Nothing)
                End If

                If dtCxp.Existe_ScalarQuery(leeXMLF(cadXML, "UUID")).ToString = "NE" Then
                    impLocTra = res.LeeXML(Archivo, "implocalT")
                    impLocRet = res.LeeXML(Archivo, "implocalR")

                    'dtCxp.Insert(leeXMLF(cadXML, "RFCE"), leeXMLF(cadXML, "RFCR"), CDbl(leeXMLF(cadXML, "Total")), CDbl(leeXMLF(cadXML, "TIR")), CDbl(leeXMLF(cadXML, "TIT")), leeXMLF(cadXML, "UUID"), leeXMLF(cadXML, "NombreE"), leeXMLF(cadXML, "Moneda"), leeXMLF(cadXML, "MetodoPago"), leeXMLF(cadXML, "FormaPago"), CDbl(leeXMLF(cadXML, "TipoCambio")), leeXMLF(cadXML, "TipoDeComprobante"), leeXMLF(cadXML, "Serie"), leeXMLF(cadXML, "Folio"), leeXMLF(cadXML, "Fecha"), leeXMLF(cadXML, "FechaT"), False, Date.Now)
                    Dim conceptos As XmlNode = res.LeeXML_Conceptos(Archivo, "Concepto")
                    Dim contDetalle As Integer = 0
                    For Each detalle_conceptos As XmlNode In conceptos.ChildNodes
                        If detalle_conceptos.Name = "cfdi:Concepto" Then
                            If detalle_conceptos.ChildNodes.Count = 0 Then
                                contDetalle += 1
                                dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")), Nothing, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", CDec(res.LeeXML(Archivo, "Total")), contDetalle.ToString, Nothing, Nothing, impLocRet, impLocTra, Nothing)
                            End If
                            For Each concepto_hijos As XmlNode In detalle_conceptos.ChildNodes
                                If concepto_hijos.Name = "cfdi:Impuestos" Then
                                    For Each impuestos_detalle As XmlNode In concepto_hijos.ChildNodes
                                        If impuestos_detalle.Name = "cfdi:Traslados" Then
                                            Dim Base As String = ""
                                            Dim Impuesto As String = ""
                                            Dim Tipofactor As String = ""
                                            Dim TasaOCuota As String = "0"
                                            Dim ImporteImpuesto As String = "0"
                                            For Each impuestos_traslado As XmlNode In impuestos_detalle.ChildNodes
                                                If impuestos_traslado.Name = "cfdi:Traslado" Then
                                                    For Each impuestos_traslado_atributos As XmlNode In impuestos_traslado.Attributes
                                                        If impuestos_traslado_atributos.Name = "Base" Then
                                                            Base = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Impuesto" Then
                                                            Impuesto = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TipoFactor" Then
                                                            Tipofactor = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TasaOCuota" Then
                                                            TasaOCuota = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Importe" Then
                                                            ImporteImpuesto = impuestos_traslado_atributos.Value.ToString
                                                        End If
                                                    Next
                                                    'Insert
                                                    contDetalle += 1
                                                    dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")), Impuesto, CDec(ImporteImpuesto), res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", CDec(res.LeeXML(Archivo, "Total")), contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, Nothing)
                                                End If
                                            Next
                                        End If
                                        If impuestos_detalle.Name = "cfdi:Retenciones" Then
                                            Dim Base As String = ""
                                            Dim Impuesto As String = ""
                                            Dim Tipofactor As String = ""
                                            Dim TasaOCuota As String = ""
                                            Dim ImporteImpuesto As String = ""
                                            For Each impuestos_traslado As XmlNode In impuestos_detalle.ChildNodes
                                                If impuestos_traslado.Name = "cfdi:Retencion" Then
                                                    For Each impuestos_traslado_atributos As XmlNode In impuestos_traslado.Attributes
                                                        If impuestos_traslado_atributos.Name = "Base" Then
                                                            Base = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Impuesto" Then
                                                            Impuesto = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TipoFactor" Then
                                                            Tipofactor = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "TasaOCuota" Then
                                                            TasaOCuota = impuestos_traslado_atributos.Value.ToString
                                                        ElseIf impuestos_traslado_atributos.Name = "Importe" Then
                                                            ImporteImpuesto = impuestos_traslado_atributos.Value.ToString
                                                        End If
                                                    Next
                                                    'Insert
                                                    contDetalle += 1
                                                    dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")), Impuesto, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", CDec(res.LeeXML(Archivo, "Total")), contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, CDec(ImporteImpuesto))
                                                End If
                                            Next

                                        End If
                                    Next
                                End If


                            Next
                        End If
                    Next

                    'dtCxp.Insert(res.LeeXML(cadXML, "RFCE"), res.LeeXML(cadXML, "RFCR"), res.LeeXML(cadXML, "SubTotal"), res.LeeXML(cadXML, "SubTotal"))
                    'System.IO.File.Move(Archivo, pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    'System.IO.File.Move(pathCxpA & nombre(0) & ".pdf", pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                    'WriteLine("Se insertó el UUID: " & leeXMLF(cadXML, "UUID").ToString)
                    System.IO.File.Move(Archivo, pathCxpF & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    System.IO.File.Move(pathCxpF & nombre(0) & ".pdf", pathCxpF & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                    'Else
                    '    System.IO.File.Move(Archivo, pathCxpF & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    '    System.IO.File.Move(pathCxpF & nombre(0) & ".pdf", pathCxpF & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                End If
            Catch ex As Exception
                'WriteLine(ex.ToString)
                MsgBox(ex.ToString)
            End Try
            File.Delete(Archivo)
            File.Delete(pathCxpF & nombre(0) & ".pdf")
        Next
    End Sub


End Module
