using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        // Se llama a la cadena de conexión definida en Web.config
        private string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadSongs();
            }
        }

        private void LoadSongs()
        {
            using (SqlConnection conn = new SqlConnection(strConnString))
            {
                // Consulta que obtiene el nombre, ruta y reproducciones de TBL_SONG[cite: 3]
                string query = "SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        // Vinculación de los datos de la BD al Repeater
                        rptSongs.DataSource = dt;
                        rptSongs.DataBind();
                    }
                }
            }
        }
    }
}