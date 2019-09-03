using System;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Windows.Forms;
using java.util;

namespace PluginSalas
{
    class UserData
    {
        public object UpdateAppointmentnewSubject { get; internal set; }

        /// <summary>
        /// Registra al usuario, sólo será necesario la primera vez que se pulse en el plugin
        /// </summary>
        /// <param name="username">El nombre del usuario</param>
        /// <param name="password">La contraseña del usuario</param>
        public void AddUser(string username,string password)
        {
            string path = GetPath();
            XDocument xdoc = XDocument.Load(path);

            xdoc.Root.Element("User").Add(
                new XAttribute("id","01"),
                new XElement("username", username),
                new XElement("password", password));
            xdoc.Save(path);  
        }

        /// <summary>
        /// Actualiza los datos del usuario en el caso de que estos hayan cambiado
        /// </summary>
        /// <param name="username">El nombre del usuario</param>
        /// <param name="password">La contraseña del usuario</param>
        public void UpdateUser(string username,string password)
        {
            string path = GetPath();
            XDocument xdoc = XDocument.Load(path);

            var user = xdoc.Descendants("User").Single(p => p.Attribute("id").Value.Equals("01"));
            user.SetElementValue("username", username);
            user.SetElementValue("password", password);
            xdoc.Save(path);
        }

        /// <summary>
        /// Obtiene el nombre del usuario
        /// </summary>
        /// <returns>string username</returns>
        public string GetUsername()
        {
            string username;
            string path = GetPath();

            XDocument xdoc = XDocument.Load(path);
            var user = xdoc.Descendants("User").Single(p => p.Attribute("id").Value.Equals("01"));
            username = user.Element("username").Value;

            return username;
        }

        /// <summary>
        /// Obtiene la contraseña del usuario
        /// </summary>
        /// <returns>string password</returns>
        public string GetPassword()
        {
            string password;
            string path = GetPath();
            XDocument xdoc = XDocument.Load(path);
            var user = xdoc.Descendants("User").Single(p => p.Attribute("id").Value.Equals("01"));
            password = user.Element("password").Value;

            return password;
        }

        /// <summary>
        /// Obtiene una reunión a partir de su id
        /// </summary>
        /// <param name="id">identificador generado por outlook de la reunión</param>
        /// <returns>arraylist appointment</returns>
        public ArrayList GetAppointment(string id)
        {
            string path = GetPath();
            ArrayList oldAppointment= new ArrayList();

            XDocument xdoc = XDocument.Load(path);
            var appointments = xdoc.Descendants("Appointment");
            foreach (var appointment in appointments)
            {
                if (appointment.Element("id").Value.Equals(id))
                {   
                    oldAppointment.add(appointment.Element("subject").Value);
                    oldAppointment.add(appointment.Element("date").Value);
                    oldAppointment.add(appointment.Element("start").Value);
                    oldAppointment.add(appointment.Element("end").Value);
                }
            }
            return oldAppointment;
        }

        public bool AppointmentExists(string id)
        {
            string path = GetPath();
            bool exists = false;

            XDocument xdoc = XDocument.Load(path);
            var appointments = xdoc.Descendants("Appointment");
            foreach (var appointment in appointments)
            {
                if (appointment.Element("id").Value.Equals(id))
                {
                    exists = true;
                }
            }
            return exists;
        }
        /// <summary>
        /// Permite saber si el usuario ha accedido al plugin y ha almacenado sus datos de login
        /// </summary>
        /// <returns>boolean: Falso->Usuario no existe, Verdadero->usuario existe</returns>
        public bool UserExists()
        {
            bool userExists = true;
            string path = GetPath();

            if (!File.Exists(path))
            {
                userExists = false;               
            }
            return userExists;
        }

        /// <summary>
        /// Permite acceder a la ruta en la que se encuentra el login y donde se creará un archivo xml con los datos
        /// del usuario
        /// </summary>
        /// <returns></returns>
        public string GetPath()
        {
            string path = ThisAddIn.AssemblyDirectory;
            string[] appPath = path.Split(new string[] { "bin" }, StringSplitOptions.None);
            string totalPath = appPath[0] + "\\user.xml";

            return totalPath;
        }

        /// <summary>
        /// Permite registrar en el fichero xml una nueva cita
        /// </summary>
        /// <param name="subject">Asunto de la cita</param>
        /// <param name="day">Día de la cita</param>
        /// <param name="start">Hora de inicio de la cita</param>
        /// <param name="end">Hora de finalización de la cita</param>
        public void AddAppointment(string id, string subject, string day, string start, string end)
        {
            DeletePastAppointments();
            string path = GetPath();
            XDocument xdoc = XDocument.Load(path);
            xdoc.Root.Element("Appointments").Add(
               new XElement("Appointment",
                    new XElement("id", id),
                    new XElement("subject", subject),
                    new XElement("date", day),
                    new XElement("start", start),
                    new XElement("end", end))
               );
           
            xdoc.Save(path);
        }

        /// <summary>
        /// Elimina la reunión más antigua entre las almacenadas en la máquin del usuario si esta tiene una antigüedad de, al menos, 7 días
        /// </summary>
        private void DeletePastAppointments()
        {
            string path = GetPath();
            DateTime date;
            DateTime today = DateTime.Now;
            XDocument xdoc = XDocument.Load(path);

            var appointments = from app in xdoc.Descendants("Appointments") select app;

            foreach(XElement appointment in appointments.Elements("Appointment"))
            {
                date = Convert.ToDateTime(appointment.Element("date").Value);
                TimeSpan difference = today - date;
                int differenceInDays = difference.Days;
                if (differenceInDays >= 7)
                {
                    appointment.Remove();
                    xdoc.Save(path);
                }
            }
        }

        /// <summary>
        /// Actualiza los datos de una reunión
        /// </summary>
        /// <param name="id">identificador de la reunión</param>
        /// <param name="newSubject">nuevo asunto de la reunión</param>
        /// <param name="newDay">nuevo día de la reunión</param>
        /// <param name="newStart">nuevo inició dela reunión</param>
        /// <param name="newEnd">nuevo fin de la reunión</param>
        public void UpdateAppointment(string id, string newSubject, string newDay, string newStart, string newEnd)
        {
            string path = GetPath();
            

            XDocument xdoc = XDocument.Load(path);
            var appointments = xdoc.Descendants("Appointment");
            foreach (var appointment in appointments)
            {
                if (appointment.Element("id").Value.Equals(id))
                {
                       appointment.Element("subject").Value = newSubject;
                       appointment.Element("date").Value = newDay;
                       appointment.Element("start").Value = newStart;
                       appointment.Element("end").Value = newEnd;
                       xdoc.Save(path);
                }
            }
        }

        /// <summary>
        /// Crea el fichero xml en la máquina del usuario donde se almacenarán los datos de este y sus reuniones
        /// </summary>
        public void CreateUserFile()
        {
            string path = GetPath();
            new XDocument(
                    new XElement("UserData",
                        new XElement("User"),
                        new XElement("Appointments"))
                ).Save(path);
        }
    }
}
