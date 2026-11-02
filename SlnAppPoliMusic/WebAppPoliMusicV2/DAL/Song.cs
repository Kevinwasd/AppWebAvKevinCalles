using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace WebAppPoliMusicV2.DAL
{
    public class UserDAL
    {
        // Se obtiene la cadena de conexión desde el Web.config
        private string connectionString = ConfigurationManager.ConnectionStrings["PoliMusicDB"].ConnectionString;

        public DataTable GetUserForLogin(string email, string password)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                // Se consultan los datos del usuario en la tabla TBL_USER
                string query = "SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE EMAIL = @Email AND PASSWORD = @Password";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Email", email);
                    cmd.Parameters.AddWithValue("@Password", password);

                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
    }
}