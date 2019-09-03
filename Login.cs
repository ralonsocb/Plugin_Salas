using java.util;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.Drawing.Text;
using System.Runtime.InteropServices;


namespace PluginSalas
{
    public partial class LoginWindow : Form
    {
        private static bool wait = false;
        ///Generamos los textos con la fuente indicada
        [DllImport("gdi32.dll")]
        private static extern IntPtr AddFontMemResourceEx(IntPtr pbfont,uint cbfont,IntPtr pdv,[In] ref uint pcFonts);
        string user = "";
        string password = "";

        FontFamily ff;
        Font font;
        
        public LoginWindow()
        { 
            InitializeComponent();
        }


        /// <summary>
        /// Botón que permite realizar el login del usuario cuando el usuario abre el plugin por primera vez
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(() => StartWaitForm());
            thread.Start();
            string username, password;
            passwordText.PasswordChar = '*';
            username = userText.Text;
            password = passwordText.Text;
            string encryptedPassword = Globals.ThisAddIn.encryptPassword(password);
            string user = "(" + username + "," + encryptedPassword + ")";
            ArrayList userFinal = Globals.ThisAddIn.encryptOutlook(user);
            var resultAutenticar = Globals.Ribbons.Ribbon1.AutenticarUsuarioOutlook("http://88.12.10.158:81/AutenticarUsuarioOutlook", userFinal);
            JObject jsonAutenticar = JObject.Parse(resultAutenticar.Result);
            string errnoAutenticar = (string)jsonAutenticar.SelectToken("errno");
            if (errnoAutenticar.Equals("0"))
            {
                thread.Abort();
                thread.Join();
                UserData data = new UserData();
                data.CreateUserFile();
                data.AddUser(username, encryptedPassword);
                this.Close();
                WaitForm waitForm = new WaitForm();
                Globals.Ribbons.Ribbon1.CreaReunion(waitForm);
            }
            else
            {
                thread.Abort();
                thread.Join();
                this.Close();
                MessageBox.Show((string)jsonAutenticar.SelectToken("error"));
            }
        }

        public void StartWaitForm()
        {
            WaitForm waitForm = new WaitForm();
            waitForm.ShowDialog();
        }

        private void UserText_TextChanged(object sender, EventArgs e)
        {
        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {
        }

        private void Label4_Click(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// Función encargada de cargar la fuente
        /// </summary>
        private void loadFont()
        {
            byte[] fontArray = PluginSalas.Properties.Resources.HelvNeue75_W1G;
            int dataLength = PluginSalas.Properties.Resources.HelvNeue75_W1G.Length;

            IntPtr ptrData = Marshal.AllocCoTaskMem(dataLength);

            Marshal.Copy(fontArray, 0, ptrData, dataLength);

            uint cFonts = 0;

            AddFontMemResourceEx(ptrData, (uint)fontArray.Length, IntPtr.Zero, ref cFonts);

            PrivateFontCollection pfc = new PrivateFontCollection();

            pfc.AddMemoryFont(ptrData, dataLength);

            Marshal.FreeCoTaskMem(ptrData);

            ff = pfc.Families[0];
            font = new Font(ff, 15f, FontStyle.Regular);
        }

        /// <summary>
        /// Función encargada de generar el estilo de la fuente
        /// </summary>
        /// <param name="f"></param>
        /// <param name="c"></param>
        /// <param name="size"></param>
        private void AllocFont(Font f, Control c, float size)
        {
            FontStyle fontStyle = FontStyle.Regular;

            c.Font = new Font(ff, size, fontStyle);
        }

        /// <summary>
        /// Función que genera los elementos de la pantalla de login
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoginWindow_Load(object sender, EventArgs e)
        {
            loadFont();
            AllocFont(font, this.label3, 11);
            AllocFont(font, this.label4, 11);
            
            AllocFont(font, this.button1, 11);
            userText.Text = "Usuario";
            userText.ForeColor = Color.LightGray;
            passwordText.Text = "Contraseña";
            passwordText.ForeColor = Color.LightGray;           
        }

        private void UserText_Enter(object sender, EventArgs e)
        {
            userText.Text = "";
            userText.ForeColor = Color.Black;
        }

        private void UserText_Leave(object sender, EventArgs e)
        {
            user = userText.Text;
            if (user.Equals("Usuario"))
            {
                userText.Text = "Usuario";
                userText.ForeColor = Color.LightGray;
            } else
            {
                if (user.Equals(""))
                {
                    userText.Text = "Usuario";
                    userText.ForeColor = Color.LightGray;
                }
                else
                {
                    userText.Text = user;
                    userText.ForeColor = Color.Black;
                }
            }
        }

        private void PasswordText_Enter(object sender, EventArgs e)
        {
            passwordText.Text = "";
            passwordText.ForeColor = Color.Black;
        }

        private void PasswordText_Leave(object sender, EventArgs e)
        {
            password = passwordText.Text;
            if (password.Equals("Contraseña"))
            {
                passwordText.Text = "Contraseña";
                passwordText.ForeColor = Color.LightGray;
            }
            else
            {
                if (password.Equals(""))
                {
                    passwordText.Text = "Contraseña";
                    passwordText.ForeColor = Color.LightGray;
                }
                else
                {
                    passwordText.Text = password;
                    passwordText.ForeColor = Color.Black;
                }
            }
        }     

    }
}
