<%@ Page Title="Canciones" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <h2>Welcome <asp:Label ID="lblUserName" runat="server" Text="kevin.calles"></asp:Label> 👨‍🏫</h2>
    <br />
    <asp:LinkButton ID="lnkLogout" runat="server" OnClick="lnkLogout_Click">Logout</asp:LinkButton>
    <br /><br />
    
    <h3>Song list:</h3>
    <br />
    
    <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered table-striped">
        <Columns>
            <asp:BoundField DataField="ID_SONG" HeaderText="ID Canción" />
            <asp:BoundField DataField="SONG_NAME" HeaderText="Nombre" />
            <asp:BoundField DataField="SONG_PATH" HeaderText="Ruta" />
            <asp:BoundField DataField="PLAYS" HeaderText="Reproducciones" />
            
            <asp:TemplateField HeaderText="Reproducir">
                <ItemTemplate>
                    <audio controls>
                        <!-- ResolveUrl es necesario para que el navegador procese el símbolo "~/" -->
                        <source src='<%# ResolveUrl(Eval("SONG_PATH").ToString()) %>' type="audio/mpeg">
                        Tu navegador no soporta el elemento de audio.
                    </audio>
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
    </asp:GridView>
</asp:Content>
