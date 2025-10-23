<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conexao.asp"-->
<!--#include file="conSunSales.asp"-->

<% ' funcional 09:22'
Response.Buffer = True
Response.Expires = -1
Response.CodePage = 65001
Response.Charset = "utf-8"

' Conexão simples
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open StrConnSales

' Query simples para agrupar por localidade
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
    <title>Mapa de Vendas por Localidade</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="https://unpkg.com/leaflet/dist/leaflet.css" />
    <style>
        body { margin: 0; padding: 0; }
        #map { height: 100vh; width: 100%; }
    </style>
</head>
<body>
    <div id="map"></div>

    <script src="https://unpkg.com/leaflet/dist/leaflet.js"></script>
    <script>
        // Dados das localidades
        var localidades = [
            <%
            Dim isFirstRecord : isFirstRecord = True ' Flag para controlar a vírgula
            
            If Not rs.EOF Then
                Do While Not rs.EOF
                    localidade = rs("Localidade")
                    VGV = rs("VGV")
                    totalVendas = rs("TotalVendas")
                    coordenada = rs("Coordenada")
                    
                    ' Extrai lat e lng
                    If InStr(coordenada, ",") > 0 Then
                        parts = Split(coordenada, ",")
                        lat = Trim(parts(0))
                        lng = Trim(parts(1))
                        
                        If IsNumeric(lat) And IsNumeric(lng) Then
                            ' Adiciona a vírgula APENAS antes do segundo registro em diante
                            If Not isFirstRecord Then Response.Write ","
                            
                            Response.Write "{"
                            Response.Write "nome: '" & Replace(localidade, "'", "\'") & "',"
                            Response.Write "vgv: " & VGV & ","
                            Response.Write "vendas: " & totalVendas & ","
                            Response.Write "lat: " & lat & ","
                            Response.Write "lng: " & lng
                            Response.Write "}"
                            
                            isFirstRecord = False ' Marca que o primeiro registro válido foi escrito
                        End If
                    End If
                    rs.MoveNext
                Loop
            End If
            rs.Close
            conn.Close
            %>
        ];

        // Inicializa o mapa (agora sem setView inicial)
        var map = L.map('map');
        
        // Camada do mapa
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '© OpenStreetMap'
        }).addTo(map);

        // Variável para coletar todas as coordenadas
        var latLngs = [];

        // Calcula máximo VGV
        var maxVGV = 0;
        for (var i = 0; i < localidades.length; i++) {
            if (localidades[i].vgv > maxVGV) maxVGV = localidades[i].vgv;
        }

        // Array de cores diferentes para os pontos
        var cores = [
            '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
            '#FFA500', '#800080', '#008000', '#800000', '#008080', '#000080',
            '#FF4500', '#2E8B57', '#DA70D6', '#191970', '#FFD700', '#DC143C',
            '#00CED1', '#FF69B4', '#8A2BE2', '#228B22', '#B22222', '#4682B4',
            '#32CD32', '#9932CC', '#FF6347', '#40E0D0', '#EE82EE', '#F4A460'
        ];

        // Adiciona círculos e coleta coordenadas
        for (var i = 0; i < localidades.length; i++) {
            var loc = localidades[i];
            
            // Garante que as coordenadas são válidas e as adiciona ao array
            if (loc.lat && loc.lng) {
                latLngs.push([loc.lat, loc.lng]);
                
                var raio = Math.max(10, (loc.vgv / maxVGV) * 50);
                
                // Seleciona uma cor do array (usa módulo para repetir cores se necessário)
                var cor = cores[i % cores.length];
            
                L.circle([loc.lat, loc.lng], {
                    radius: raio * 100, // Multiplica para ficar visível
                    fillColor: cor,
                    color: '#000',
                    weight: 1,
                    opacity: 0.8,
                    fillOpacity: 0.6
                })
                .bindPopup('<b>' + loc.nome + '</b><br>VGV: R$ ' + loc.vgv.toLocaleString('pt-BR') + '<br>Vendas: ' + loc.vendas)
                .addTo(map);
            }
        }
        
        // LÓGICA DE ZOOM AUTOMÁTICO (fitBounds)
        if (latLngs.length > 0) {
            // 1. Cria um objeto L.LatLngBounds a partir da matriz de coordenadas.
            var bounds = L.latLngBounds(latLngs);
            
            // 2. Ajusta o mapa para se encaixar nos limites, adicionando um pequeno padding
            // para que os círculos não fiquem colados nas bordas.
            map.fitBounds(bounds, {
                padding: [20, 20]
            });
        } else {
            // Se não houver dados, define uma visualização padrão para evitar mapa vazio.
            map.setView([-15, -55], 4); // Visualização do Brasil Central (ajuste se necessário)
        }

        console.log('Mapa carregado com ' + localidades.length + ' localidades e zoom ajustado.');
    </script>
</body>
</html>