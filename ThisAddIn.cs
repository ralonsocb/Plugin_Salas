using System;
using System.Reflection;
using System.IO;
using java.util;
using javax.crypto.spec;
using java.text;
using System.Text;
using java.security.spec;
using javax.crypto;
using java.security;
using java.io;
using System.Windows.Forms;

namespace PluginSalas
{
    public partial class ThisAddIn
    {
       
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {           
         
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Nota: Outlook ya no genera este evento. Si tiene código que 
            //    se debe ejecutar cuando Outlook se apaga, consulte https://go.microsoft.com/fwlink/?LinkId=506785
        }

        /// <summary>
        /// Devuelve la ruta de ejecución del plguin, donde se almacenará al usuario en un fichero xml.
        /// </summary>
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        /// <summary>
        /// Función encargada de encriptar la contraseña
        /// </summary>
        /// <param name="passwordToHash">contraseña sin cifrar</param>
        /// <returns></returns>
        public string encryptPassword(string passwordToHash) 
        {
            string generatedPassword = null;
            string salt = "0uTl@k";
            try
            {
                MessageDigest md = MessageDigest.getInstance("SHA-256");
                md.update(Encoding.UTF8.GetBytes(salt));
                byte[] bytes = md.digest(Encoding.UTF8.GetBytes(passwordToHash));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    sb.Append(java.lang.Integer.toString((bytes[i] & 0xff) + 0x100, 16).Substring(1));
                }
                generatedPassword = sb.ToString();
            }
            catch (NoSuchAlgorithmException e)
            {             
                e.printStackTrace();
            }
            return generatedPassword;
            throw new UnsupportedEncodingException();
        }

        /// <summary>
        /// Función encargada de encriptar los datos del usuario
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public ArrayList encryptOutlook(string password)
        {
            string semilla = "0uTl@k";
            string marcaTiempo = "";
            ArrayList resultado = new ArrayList();
            string encriptado = "";
            try
            {
                // do
                // {              
                    byte[] iv = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                    IvParameterSpec ivspec = new IvParameterSpec(iv);
                    marcaTiempo = (new SimpleDateFormat("ddMMyyyyHHmmss")).format(new Date());

                    KeySpec clave = new PBEKeySpec(marcaTiempo.ToCharArray(), Encoding.Default.GetBytes(semilla), 65536, 256);
                    SecretKey hash = SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256").generateSecret(clave);
                    SecretKeySpec key = new SecretKeySpec(hash.getEncoded(), "AES");

                    Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
                    cipher.init(Cipher.ENCRYPT_MODE, key, ivspec);               
                    encriptado = Base64.getEncoder().encodeToString(cipher.doFinal(Encoding.UTF8.GetBytes(password)));          
                    resultado.add(encriptado);
                    resultado.add(marcaTiempo);
            }
            catch (Exception e)
            {
                System.Console.WriteLine("Error en la encriptacion: " + e.ToString());
                 resultado = new ArrayList();
            }
            return resultado;
        }


        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
