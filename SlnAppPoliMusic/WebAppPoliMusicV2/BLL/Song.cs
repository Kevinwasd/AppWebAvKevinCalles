using System;
using System.Data;
using WebAppPoliMusicV2.DAL; // Referencia a la capa de datos

namespace WebAppPoliMusicV2.BLL
{
    public class UserBLL
    {
        // Instancia de la clase DAL
        private UserDAL userDal = new UserDAL();

        public DataTable AuthenticateUser(string email, string password)
        {
            // Regla de negocio: Validar que los campos no estén vacíos antes de ir a la base de datos
            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            {
                throw new ArgumentException("El correo y la contraseña son obligatorios.");
            }

            // Llamada al método de la DAL
            return userDal.GetUserForLogin(email, password);
        }
    }
}