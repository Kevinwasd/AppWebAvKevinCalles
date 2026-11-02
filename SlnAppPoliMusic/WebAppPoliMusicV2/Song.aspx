<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>
<!DOCTYPE html>
<html lang="es">
<head runat="server">
    <title>Lista de Canciones - PoliMusic</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 20px; }
        .player-container { text-align: center; margin-bottom: 30px; padding: 20px; background: #eef2f3; border-radius: 8px; border: 1px solid #dcdcdc; }
        #currentSong { font-weight: 600; margin-bottom: 10px; display: block; color: #e74c3c; font-size: 1.1em; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #e1e1e1; }
        th { background-color: #34495e; color: white; font-weight: 500; }
        tr:hover { background-color: #f9f9f9; }
        .btn-play { background-color: #27ae60; color: white; border: none; padding: 8px 15px; cursor: pointer; border-radius: 4px; transition: background 0.3s; }
        .btn-play:hover { background-color: #219150; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="container">
            <h2>Explorar Canciones</h2>

            <div class="player-container">
                <span id="currentSong">Selecciona una canción de la lista</span>
                <audio id="audioPlayer" controls style="width: 100%; outline: none;">
                    <source id="audioSource" src="" type="audio/mpeg">
                    Tu navegador no soporta el reproductor de audio.
                </audio>
            </div>

            <table>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Nombre</th>
                        <th>Reproducciones</th>
                        <th>Acción</th>
                    </tr>
                </thead>
                <tbody>
                    <asp:Repeater ID="rptSongs" runat="server">
                        <ItemTemplate>
                            <tr>
                                <!-- Consumiendo los datos de la tabla TBL_SONG -->
                                <td><%# Eval("ID_SONG") %></td>
                                <td><%# Eval("SONG_NAME") %></td>
                                <td><%# Eval("PLAYS") %></td>
                                <td>
                                    <!-- Pasa la ruta (resolviendo el ~) y el nombre a la función JS -->
                                    <button type="button" class="btn-play" onclick="playSong('<%# ResolveUrl(Eval("SONG_PATH").ToString()) %>', '<%# Eval("SONG_NAME") %>')">Reproducir</button>
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </tbody>
            </table>
        </div>
    </form>

    <script>
        function playSong(path, name) {
            var player = document.getElementById('audioPlayer');
            var source = document.getElementById('audioSource');
            var title = document.getElementById('currentSong');
            
            source.src = path;
            title.innerText = "Reproduciendo: " + name;
            
            player.load(); // Recarga el elemento de audio con la nueva ruta
            player.play(); // Inicia la reproducción
        }
    </script>
</body>
</html>
