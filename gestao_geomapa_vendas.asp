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
    <!-- Adiciona Font Awesome para o ícone de fechar -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css" />
    <style>
        /* Define o layout principal para usar flexbox em toda a tela */
        body {
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            height: 100vh;
            font-family: Arial, sans-serif;
        }

        /* Estilização da nova barra superior */
        #header-bar {
            background-color: #2c3e50; /* Cor escura elegante */
            color: white;
            padding: 10px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            z-index: 1000; /* Garante que fique acima do mapa */
            flex-shrink: 0; /* Impede que a barra diminua */
        }

        #header-bar h1 {
            margin: 0;
            font-size: 1.25rem;
            font-weight: 400;
        }

        /* Estilização do botão Fechar */
        .close-btn {
            background-color: #e74c3c; /* Vermelho/alerta */
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

        /* O mapa agora ocupa o espaço restante (flex-grow: 1) */
        /* Isso garante que ele preencha a área abaixo da barra superior */
        #map {
            flex-grow: 1;
            width: 100%;
        }
    </style>
    <style>
    body {
        /* Define a escala de 0.8 (80%) */
        transform: scale(0.8); 
        
        /* Define o ponto de origem para o canto superior esquerdo */
        transform-origin: 0 0; 
        
        /* Ajusta a largura para que o conteúdo ocupe 80% da largura original */
        /* Isso ajuda a prevenir barras de rolagem desnecessárias. */
        width: calc(100% / 0.8); 
    }
</style>
</head>
<body>

    <!-- BARRA SUPERIOR ADICIONADA -->
    <div id="header-bar">
        <h1>Mapa de Vendas por Localidade</h1>
        <!-- Botão Fechar que executa window.close() -->
        <a href="javascript:window.close()" class="close-btn" title="Fechar a aba do navegador">
            <i class="fas fa-times me-1"></i> Fechar
        </a>
    </div>

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
                            // Adiciona a vírgula APENAS antes do segundo registro em diante
                            If Not isFirstRecord Then Response.Write ","
                            
                            Response.Write "{"
                            Response.Write "nome: '" & Replace(localidade, "'", "\'") & "',"
                            Response.Write "vgv: " & VGV & ","
                            Response.Write "vendas: " & totalVendas & ","
                            Response.Write "lat: " & lat & ","
                            Response.Write "lng: " & lng
                            Response.Write "}"
                            
                            isFirstRecord = False // Marca que o primeiro registro válido foi escrito
                        End If
                    End If
                    rs.MoveNext
                Loop
            End If
            rs.Close
            conn.Close
            %>
        ];

        // Coordenada Central Solicitada: -8.506219, -35.000454
        var CENTER_LAT = -8.506219;
        var CENTER_LNG = -35.000454;
        var DEFAULT_ZOOM = 10; // Nível de zoom razoável para a localização

        // Inicializa o mapa centralizando na coordenada solicitada como ponto de partida
        var map = L.map('map').setView([CENTER_LAT, CENTER_LNG], DEFAULT_ZOOM);
        
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
            '#3498db', '#2ecc71', '#f1c40f', '#e74c3c', '#9b59b6', '#1abc9c',
            '#f39c12', '#d35400', '#c0392b', '#2980b9', '#27ae60', '#8e44ad',
            '#e67e22', '#34495e', '#7f8c8d', '#bdc3c7', '#ecf0f1', '#95a5a6'
        ];

        // Adiciona círculos e coleta coordenadas
        for (var i = 0; i < localidades.length; i++) {
            var loc = localidades[i];
            
            // Garante que as coordenadas são válidas e as adiciona ao array
            if (loc.lat && loc.lng) {
                latLngs.push([loc.lat, loc.lng]);
                
                // Raio escalonado pelo VGV (mínimo 10, máximo 50)
                var raio = Math.max(10, (loc.vgv / maxVGV) * 50);
                
                // Seleciona uma cor do array (usa módulo para repetir cores se necessário)
                var cor = cores[i % cores.length];
            
                L.circle([loc.lat, loc.lng], {
                    radius: raio * 100, // Multiplica para ficar visível (em metros)
                    fillColor: cor,
                    color: '#2c3e50', /* Borda escura */
                    weight: 1,
                    opacity: 0.8,
                    fillOpacity: 0.7
                })
                .bindPopup('<b>' + loc.nome + '</b><br>VGV: R$ ' + loc.vgv.toLocaleString('pt-BR') + '<br>Vendas: ' + loc.vendas)
                .addTo(map);
            }
        }
        
        // LÓGICA DE ZOOM AUTOMÁTICO (fitBounds)
        // Se houver pontos no mapa, o fitBounds sobrescreve o setView inicial para mostrar todos os marcadores
        if (latLngs.length > 0) {
            // 1. Cria um objeto L.LatLngBounds a partir da matriz de coordenadas.
            var bounds = L.latLngBounds(latLngs);
            
            // 2. Ajusta o mapa para se encaixar nos limites, adicionando um pequeno padding
            map.fitBounds(bounds, {
                padding: [20, 20] // Padding (margem) em pixels [top-left, bottom-right]
            });
        } else {
            // Se não houver dados, o mapa permanece na coordenada inicial e zoom padrão
            // (que já foi definido no L.map('map').setView(...) )
        }

        console.log('Mapa carregado com ' + localidades.length + ' localidades. Centro inicial: ' + CENTER_LAT + ', ' + CENTER_LNG);
    </script>
</body>
</html>