<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<%
' Configuração da resposta
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' 1. Conexão e Execução da Query
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
' Garante que a string de conexão 'StrConnSales' esteja definida em conSunSales.asp
conn.Open StrConnSales 

' Query para agrupar por localidade
Dim sql
sql = "SELECT Localidade, SUM(ValorUnidade) as VGV, COUNT(*) as TotalVendas, " & _
      "MIN(Localizacao) as Coordenada " & _
      "FROM Vendas " & _
      "WHERE Localidade IS NOT NULL AND Localidade <> '' " & _
      "AND Localizacao IS NOT NULL AND Localizacao <> '' " & _
      "AND ValorUnidade > 0 " & _
      "GROUP BY Localidade " & _
      "HAVING SUM(ValorUnidade) > 0"

Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html>
<head>
    <title>SGVendas Geo-Mapa de Vendas</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="https://unpkg.com/leaflet/dist/leaflet.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css" />
    <style>
        /* Estilos CSS (inalterados) */
        body {
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            height: 100vh;
            font-family: Arial, sans-serif;
        }

        #header-bar {
            background-color: #2c3e50; 
            color: white;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            z-index: 1000;
            flex-shrink: 0;
        }

        #header-bar h1 {
            margin: 0;
            font-size: 1.25rem;
            font-weight: 400;
        }

        .close-btn {
            background-color: #e74c3c; 
            color: white;
            padding: 5px 15px;
            border-radius: 5px;
            text-decoration: none;
            font-weight: bold;
            transition: background-color 0.2s;
            border: none;
            cursor: pointer;
        }

        .close-btn:hover {
            background-color: #c0392b;
        }

        #map {
            flex-grow: 1;
            width: 100%;
        }
    </style>
</head>
<body>

    <div id="header-bar">
        <h1>Mapa de Vendas por Localidade</h1>
        <small><%=Session("Usuario")%></small>
        <a href="javascript:window.close()" class="close-btn" title="Fechar a aba do navegador">
            <i class="fas fa-times me-1"></i> Fechar
        </a>
    </div>

    <div id="map"></div>

    <script src="https://unpkg.com/leaflet/dist/leaflet.js"></script>
    <script>
        // 2. Geração Dinâmica do Array JavaScript com VBScript
        var localidades = [
            <%
            Dim isFirstRecord : isFirstRecord = True ' Flag para controlar a vírgula
            Dim localidade, VGV, totalVendas, coordenada, VGV_formatado, parts, lat, lng
            
            If Not rs.EOF Then
                Do While Not rs.EOF
                    localidade = Trim(rs("Localidade"))
                    VGV = rs("VGV")
                    VGV = Replace(CStr(VGV), ",", ".")
                    'response.write VGV
                    'Response.END 
                    totalVendas = rs("TotalVendas")
                    coordenada = Trim(rs("Coordenada"))
                    
                    ' CORREÇÃO CRÍTICA #1: Converter VGV para string e usar ponto decimal para JS
                    VGV_formatado = Replace(CStr(VGV), ",", ".") 
                    
                    ' Extrai lat e lng
                    If InStr(coordenada, ",") > 0 Then
                        parts = Split(coordenada, ",")
                        
                        ' Verifica se há pelo menos duas partes (lat e lng)
                        If UBound(parts) >= 1 Then
                            lat = Trim(parts(0))
                            lng = Trim(parts(1))
                            
                            ' CORREÇÃO CRÍTICA #2: Garantir que coordenadas usem ponto decimal para JS
                            lat = Replace(lat, ",", ".")
                            lng = Replace(lng, ",", ".")
                            
                            ' Verifica se todos os valores são numéricos antes de escrever o objeto JS
                            If IsNumeric(lat) And IsNumeric(lng) And IsNumeric(VGV_formatado) Then
                                ' Adiciona a vírgula APENAS antes do segundo registro em diante
                                If Not isFirstRecord Then Response.Write ","
                                
                                Response.Write "{"
                                Response.Write "nome: '" & Replace(localidade, "'", "\'") & "',"
                                Response.Write "vgv: " & VGV_formatado & ","
                                Response.Write "vendas: " & totalVendas & ","
                                Response.Write "lat: " & lat & ","
                                Response.Write "lng: " & lng
                                Response.Write "}"
                                
                                isFirstRecord = False ' Marca que o primeiro registro válido foi escrito
                            End If
                        End If
                    End If
                    rs.MoveNext
                Loop
            End If
            
            ' Fecha o recordset e a conexão
            rs.Close
            Set rs = Nothing
            conn.Close
            Set conn = Nothing
            %>
        ];

        // 3. Inicialização e Renderização do Mapa (Leaflet JS)

        // Coordenada Central Solicitada: -8.506219, -35.000454
        var CENTER_LAT = -8.506219;
        var CENTER_LNG = -35.000454;
        var DEFAULT_ZOOM = 10; 

        // Inicializa o mapa
        var map = L.map('map').setView([CENTER_LAT, CENTER_LNG], DEFAULT_ZOOM);
        
        // Camada do mapa
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '© OpenStreetMap'
        }).addTo(map);

        var latLngs = [];
        var maxVGV = 0;
        
        // Calcula máximo VGV
        for (var i = 0; i < localidades.length; i++) {
            // Usa parseFloat para garantir que o VGV formatado seja tratado como número
            var currentVGV = parseFloat(localidades[i].vgv); 
            if (currentVGV > maxVGV) maxVGV = currentVGV;
        }

        // Array de cores
        var cores = [
            '#3498db', '#2ecc71', '#f1c40f', '#e74c3c', '#9b59b6', '#1abc9c',
            '#f39c12', '#d35400', '#c0392b', '#2980b9', '#27ae60', '#8e44ad',
            '#e67e22', '#34495e', '#7f8c8d', '#bdc3c7', '#ecf0f1', '#95a5a6'
        ];

        // Adiciona círculos e coleta coordenadas
        for (var i = 0; i < localidades.length; i++) {
            var loc = localidades[i];
            
            // Tenta converter lat e lng para números de ponto flutuante
            var lat = parseFloat(loc.lat);
            var lng = parseFloat(loc.lng);
            var vgv = parseFloat(loc.vgv);
            
            // Garante que as coordenadas são válidas
            if (!isNaN(lat) && !isNaN(lng) && !isNaN(vgv)) {
                latLngs.push([lat, lng]);
                
                // Raio escalonado pelo VGV (mínimo 10, máximo 50)
                var raio = Math.max(10, (vgv / maxVGV) * 50);
                
                // Seleciona uma cor
                var cor = cores[i % cores.length];
                
                L.circle([lat, lng], {
                    radius: raio * 100, // Multiplica para ficar visível (em metros)
                    fillColor: cor,
                    color: '#2c3e50', /* Borda escura */
                    weight: 1,
                    opacity: 0.8,
                    fillOpacity: 0.7
                })
                // Formatação do Popup (uso de toLocaleString para formato de moeda brasileira)
                .bindPopup('<b>' + loc.nome + '</b><br>VGV: R$ ' + vgv.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '<br>Vendas: ' + loc.vendas)
                .addTo(map);
            }
        }
        
        // LÓGICA DE ZOOM AUTOMÁTICO (fitBounds)
        if (latLngs.length > 0) {
            var bounds = L.latLngBounds(latLngs);
            map.fitBounds(bounds, {
                padding: [20, 20]
            });
        }

        console.log('Mapa carregado com ' + localidades.length + ' localidades. Centro inicial: ' + CENTER_LAT + ', ' + CENTER_LNG);
    </script>
</body>
</html>