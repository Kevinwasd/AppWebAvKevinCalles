using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        // Obtenemos la cadena de conexión desde el Web.config
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // Mostramos el nombre del usuario logueado usando la sesión
                if (Session[EntityLayer.PoliMusicV3.Util.Constants.USER] != null)
                {
                    EntityLayer.PoliMusicV3.User currentUser = (EntityLayer.PoliMusicV3.User)Session[EntityLayer.PoliMusicV3.Util.Constants.USER];
                    lblUserName.Text = currentUser.UserName;
                }

                // Cargamos la tabla de canciones
                LoadSongs();
            }
        }

        private void LoadSongs()
        {
            // Instanciamos la clase Song de la capa de Negocio (BLL)
            DataTable dtSongs = new BLL.PoliMusicV3.Song(strConnString).Read();

            // Enlazamos los datos al GridView
            gvSongs.DataSource = dtSongs;
            gvSongs.DataBind();
        }

        protected void lnkLogout_Click(object sender, EventArgs e)
        {
            FormsAuthentication.SignOut();
            Response.Redirect("Login.aspx");
        }
    }
}