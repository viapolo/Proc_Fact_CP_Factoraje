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
    Dim argumento() As String
    'Dim path As String = "C:\Users\vicente-apolonio\Desktop\temp\"
    Sub Main()
        Dim dtF100 As New XML_FactorajeDSTableAdapters.Vw_ChequesDetalleTableAdapter
        Dim dtWebXML As New XML_FactorajeDSTableAdapters.WEB_FacturasXMLTableAdapter
        Dim D As System.IO.DirectoryInfo
        D = New System.IO.DirectoryInfo(path)

        argumento = Environment.GetCommandLineArgs()

        If argumento.Length > 1 Then
            If UCase(argumento(1)) = "SALDO" Then
                enviaSaldoComprobacion()
            End If
        End If

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

    Public Sub enviaSaldoComprobacion()
        Dim taCorreos As New XML_CXPDSTableAdapters.GEN_Correos_SistemaFinagilTableAdapter

        Dim taSaldo As New XML_CXPDSTableAdapters.Vw_CXP_SaldoComprobacionGastosTableAdapter
        Dim rwSaldo As XML_CXPDS.Vw_CXP_SaldoComprobacionGastosRow
        Dim rwSaldoDetalle As XML_CXPDS.Vw_CXP_SaldoComprobacionGastosRow
        Dim dtSaldo As New XML_CXPDS.Vw_CXP_SaldoComprobacionGastosDataTable
        Dim dtSaldoDetalle As New XML_CXPDS.Vw_CXP_SaldoComprobacionGastosDataTable

        Dim taEmpresas As New XML_CXPDSTableAdapters.CXP_EmpresasTableAdapter
        Dim rowEmpresas As XML_CXPDS.CXP_EmpresasRow
        Dim dtEmpresas As New XML_CXPDS.CXP_EmpresasDataTable
        Dim mensaje As String = ""

        taEmpresas.Fill(dtEmpresas)

        For Each rowEmpresas In dtEmpresas 'recorre empresas
            taSaldo.ObtSaldoPorUsuario_FillBy(dtSaldo, rowEmpresas.idConceptoGastos)

            For Each rwSaldo In dtSaldo 'recorre usuarios
                mensaje = "<html><body><font size=3 face=" & Chr(34) & "Arial" & Chr(34) & ">" &
                    "<h1><font size=3 align" & Chr(34) & "center" & Chr(34) & ">" & "Estimado (a): " & rwSaldo.nombre & vbNewLine & ", le notificamos que cuenta con un saldo pendiente por comprobar:  </font></h1>" &
                    "<table  align=" & Chr(34) & "center" & Chr(34) & " border=1 cellspacing=0 cellpadding=2>" &
                    "<tr>" &
                        "<td>Folio de Solicitud</td>" &
                        "<td>Beneficiario</td>" &
                         "<td>Fecha de Solicitud</td>" &
                        "<td>Importe Solicitado</td>" &
                        "<td>Importe por Comprobar</td>" &
                    "</tr>"
                taSaldo.ObtieneDetalle__FillBy(dtSaldoDetalle, rowEmpresas.idEmpresas, rowEmpresas.idConceptoGastos, rwSaldo.usuario)

                Dim contDetalle As Integer = 0
                For Each rwSaldoDetalle In dtSaldoDetalle
                    If contDetalle < 10 Then
                        mensaje = mensaje &
                        "<tr>" &
                            "<td>" & rwSaldoDetalle.folioSolicitud & "</td>" &
                            "<td>" & rwSaldoDetalle.razonSocial & "</td>" &
                            "<td>" & rwSaldoDetalle.fechaSolicitud & "</td>" &
                            "<td>" & rwSaldoDetalle.totalPagado.ToString("c") & "</td>" &
                            "<td>" & rwSaldoDetalle.saldoSolicitud.ToString("c") & "</td>" &
                        "</tr>"
                    End If
                    contDetalle += 1
                Next
                mensaje = mensaje & "</table>" & vbNewLine &
                "<HR width=20%>" &
                "<tfoot><tr><font align=" & Chr(34) & "center" & Chr(34) & "size=3 face=" & Chr(34) & "Arial" & Chr(34) & ">" & "Atentamente: " & rowEmpresas.razonSocial & vbNewLine & "</font></tr></tfoot>" &
                     "</body></html>"
                taCorreos.Insert("Gastos@finagil.com.mx", rwSaldo.mail, "Saldo pendiente por comprobar", mensaje, 0, Date.Now, "")
                taCorreos.Insert("Gastos@finagil.com.mx", "viapolo@finagil.com.mx", "Saldo pendiente por comprobar", mensaje, 0, Date.Now, "")
                taCorreos.Insert("Gastos@finagil.com.mx", "lgarcia@finagil.com.mx", "Saldo pendiente por comprobar", mensaje, 0, Date.Now, "")
            Next
            dtSaldo.Dispose()
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
            'Mensaje.To.Add("jdelgado@finagil.com.mx")
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
            Dim totalGl As Decimal = 0
            cadena = New StreamReader(Archivo)
            cadXML = cadena.ReadToEnd
            cadena.Close()


            Try
                If res.LeeXML(Archivo, "RFCE") = "ASE930924SS7" Then
                    If res.LeeXML(Archivo, "Edenred") <> "" Then
                        totalGl = CDec(res.LeeXML(Archivo, "Edenred"))
                    Else
                        totalGl = CDec(res.LeeXML(Archivo, "Total"))
                    End If
                Else
                        totalGl = CDec(res.LeeXML(Archivo, "Total"))
                End If

                If dtProveedores.ExisteProv_ScalarQuery(res.LeeXML(Archivo, "RFCE")) = "NE" Then
                    dtProveedores.Insert(res.LeeXML(Archivo, "RFCE"), Nothing, Nothing, res.LeeXML(Archivo, "NombreE"), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, System.Data.SqlTypes.SqlDateTime.Null, Nothing, Nothing, Nothing, Nothing, Nothing)
                End If

                If dtCxp.Existe_ScalarQuery(leeXMLF(cadXML, "UUID")).ToString = "NE" Then
                    impLocTra = res.LeeXML(Archivo, "implocalT")
                    impLocRet = res.LeeXML(Archivo, "implocalR")

                    Dim conceptos As XmlNode = res.LeeXML_Conceptos(Archivo, "Concepto")
                    Dim contDetalle As Integer = 0
                    For Each detalle_conceptos As XmlNode In conceptos.ChildNodes
                        If detalle_conceptos.Name = "cfdi:Concepto" Then
                            Dim importe As String = Nothing
                            Dim claveSat As String = Nothing
                            Dim descuento As String = Nothing
                            Dim importeCuentaDeTerceros As String = Nothing
                            Dim concepto As String = Nothing
                            For Each conceptoDetalle As XmlNode In detalle_conceptos.Attributes
                                If conceptoDetalle.Name = "Descripcion" Then
                                    concepto = conceptoDetalle.Value.ToString
                                    If concepto.Length > 300 Then
                                        concepto = concepto.Substring(0, 299)
                                    End If
                                End If
                            Next
                            For Each complementosConcepto As XmlNode In detalle_conceptos.ChildNodes
                                If complementosConcepto.Name = "cfdi:ComplementoConcepto" Then
                                    For Each porCuentaDeTerceros As XmlNode In complementosConcepto.ChildNodes
                                        If porCuentaDeTerceros.Name = "terceros:PorCuentadeTerceros" Then
                                            For Each hijoPorCuentaDeTerceros As XmlNode In porCuentaDeTerceros.ChildNodes
                                                If hijoPorCuentaDeTerceros.Name = "terceros:Parte" Then
                                                    For Each tercerosParte As XmlNode In hijoPorCuentaDeTerceros.Attributes
                                                        If tercerosParte.Name = "importe" Then
                                                            importeCuentaDeTerceros = tercerosParte.Value.ToString
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                            For Each atributosConceptos As XmlNode In detalle_conceptos.Attributes
                                If atributosConceptos.Name = "Importe" Then
                                    importe = atributosConceptos.Value.ToString
                                ElseIf atributosConceptos.Name = "ClaveProdServ" Then
                                    claveSat = atributosConceptos.Value.ToString
                                ElseIf atributosConceptos.Name = "Descuento" Then
                                    descuento = atributosConceptos.Value.ToString
                                End If
                            Next
                            If detalle_conceptos.ChildNodes.Count = 0 Then
                                contDetalle += 1
                                dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Nothing, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Nothing, Nothing, impLocRet, impLocTra, Nothing, Nothing, Nothing, importe, claveSat, descuento, importeCuentaDeTerceros, concepto)
                            End If


                            For Each concepto_hijos As XmlNode In detalle_conceptos.ChildNodes
                                'If concepto_hijos.Name = "cfdi:ComplementoConcepto" Then
                                '    For Each concepto_hijo_complemento As XmlNode In concepto_hijos.ChildNodes
                                '        If concepto_hijo_complemento.Name = "iedu:instEducativas" Then
                                '            contDetalle += 1
                                '            dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Nothing, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Nothing, Nothing, impLocRet, impLocTra, Nothing, Nothing, Nothing, Nothing, Nothing)
                                '        End If
                                '    Next
                                'End If

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
                                                    dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Impuesto, CDec(ImporteImpuesto), res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, Nothing, Base, Nothing, importe, claveSat, descuento, importeCuentaDeTerceros, concepto)
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
                                                    dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Impuesto, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, CDec(ImporteImpuesto), Nothing, Base, importe, claveSat, descuento, importeCuentaDeTerceros, concepto)
                                                End If
                                            Next

                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next

                    System.IO.File.Move(Archivo, pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    System.IO.File.Move(pathCxpA & nombre(0) & ".pdf", pathCxpA & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                End If
            Catch ex As Exception
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
            Dim totalGl As Decimal = 0
            cadena = New StreamReader(Archivo)
            cadXML = cadena.ReadToEnd
            cadena.Close()

            Try

                If res.LeeXML(Archivo, "RFCE") = "ASE930924SS7" Then
                    If res.LeeXML(Archivo, "Edenred") <> "" Then
                        totalGl = CDec(res.LeeXML(Archivo, "Edenred"))
                    Else
                        totalGl = CDec(res.LeeXML(Archivo, "Total"))
                    End If
                Else
                    totalGl = CDec(res.LeeXML(Archivo, "Total"))
                End If

                If dtProveedores.ExisteProv_ScalarQuery(res.LeeXML(Archivo, "RFCE")) = "NE" Then
                    dtProveedores.Insert(res.LeeXML(Archivo, "RFCE"), Nothing, Nothing, res.LeeXML(Archivo, "NombreE"), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, System.Data.SqlTypes.SqlDateTime.Null, Nothing, Nothing, Nothing, Nothing, Nothing)
                End If

                If dtCxp.Existe_ScalarQuery(leeXMLF(cadXML, "UUID")).ToString = "NE" Then
                    impLocTra = res.LeeXML(Archivo, "implocalT")
                    impLocRet = res.LeeXML(Archivo, "implocalR")

                    Dim conceptos As XmlNode = res.LeeXML_Conceptos(Archivo, "Concepto")
                    Dim contDetalle As Integer = 0
                    For Each detalle_conceptos As XmlNode In conceptos.ChildNodes
                        If detalle_conceptos.Name = "cfdi:Concepto" Then
                            Dim importe As String = Nothing
                            Dim claveSat As String = Nothing
                            Dim descuento As String = Nothing
                            Dim importeCuentaDeTerceros As String = Nothing
                            Dim concepto As String = Nothing
                            For Each conceptoDetalle As XmlNode In detalle_conceptos.Attributes
                                If conceptoDetalle.Name = "Descripcion" Then
                                    concepto = conceptoDetalle.Value.ToString
                                    If concepto.Length > 300 Then
                                        concepto = concepto.Substring(0, 299)
                                    End If
                                End If
                            Next
                            For Each complementosConcepto As XmlNode In detalle_conceptos.ChildNodes
                                If complementosConcepto.Name = "cfdi:ComplementoConcepto" Then
                                    For Each porCuentaDeTerceros As XmlNode In complementosConcepto.ChildNodes
                                        If porCuentaDeTerceros.Name = "terceros:PorCuentadeTerceros" Then
                                            For Each hijoPorCuentaDeTerceros As XmlNode In porCuentaDeTerceros.ChildNodes
                                                If hijoPorCuentaDeTerceros.Name = "terceros:Parte" Then
                                                    For Each tercerosParte As XmlNode In hijoPorCuentaDeTerceros.Attributes
                                                        If tercerosParte.Name = "importe" Then
                                                            importeCuentaDeTerceros = tercerosParte.Value.ToString
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                            For Each atributosConceptos As XmlNode In detalle_conceptos.Attributes
                                If atributosConceptos.Name = "Importe" Then
                                    importe = atributosConceptos.Value.ToString
                                ElseIf atributosConceptos.Name = "ClaveProdServ" Then
                                    claveSat = atributosConceptos.Value.ToString
                                ElseIf atributosConceptos.Name = "Descuento" Then
                                    descuento = atributosConceptos.Value.ToString
                                End If
                            Next
                            If detalle_conceptos.ChildNodes.Count = 0 Then
                                contDetalle += 1
                                dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Nothing, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Nothing, Nothing, impLocRet, impLocTra, Nothing, Nothing, Nothing, importe, claveSat, descuento, importeCuentaDeTerceros, concepto)
                            End If
                            'valida si nodo impuestos existe

                            If (res.LeeXML(Archivo, "ExisteImpuestos")) = Nothing Then
                                'MsgBox("hola")
                                dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Nothing, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Nothing, Nothing, impLocRet, impLocTra, Nothing, Nothing, Nothing, importe, claveSat, descuento, importeCuentaDeTerceros, concepto)
                            End If
                            'termina valida nodo impuestos

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
                                                        dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Impuesto, CDec(ImporteImpuesto), res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, Nothing, Base, Nothing, importe, claveSat, descuento, importeCuentaDeTerceros, concepto)
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
                                                        dtCxp.Insert(res.LeeXML(Archivo, "RFCE"), res.LeeXML(Archivo, "RFCR"), CDec(res.LeeXML(Archivo, "SubTotal")) - CDec(res.LeeXML(Archivo, "Descuento")), Impuesto, Nothing, res.LeeXML(Archivo, "UUID"), res.LeeXML(Archivo, "NombreE"), res.LeeXML(Archivo, "Moneda"), res.LeeXML(Archivo, "MetodoPago"), res.LeeXML(Archivo, "FormaPago"), CDec(res.LeeXML(Archivo, "TipoCambio")), res.LeeXML(Archivo, "TipoDeComprobante"), res.LeeXML(Archivo, "Serie"), res.LeeXML(Archivo, "Folio"), res.LeeXML(Archivo, "Fecha"), res.LeeXML(Archivo, "FechaTimbrado"), System.Data.SqlTypes.SqlDateTime.Null, "PENDIENTE", totalGl, contDetalle.ToString, Tipofactor, CDec(TasaOCuota), impLocRet, impLocTra, CDec(ImporteImpuesto), Nothing, Base, importe, claveSat, descuento, importeCuentaDeTerceros, concepto)
                                                    End If
                                                Next

                                            End If
                                        Next
                                    End If
                                Next
                            End If
                    Next

                    System.IO.File.Move(Archivo, pathCxpF & "Procesados\" & leeXMLF(cadXML, "UUID") & ".xml")
                    System.IO.File.Move(pathCxpF & nombre(0) & ".pdf", pathCxpF & "Procesados\" & leeXMLF(cadXML, "UUID") & ".pdf")
                End If
            Catch ex As Exception
                'MsgBox(ex.ToString)
            End Try
            File.Delete(Archivo)
            File.Delete(pathCxpF & nombre(0) & ".pdf")
        Next
    End Sub


End Module
