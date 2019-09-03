using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Net.Http;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using java.util;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.IO;

namespace PluginSalas
{
    public partial class Ribbon1
    {
        private static bool wait = false;
        

        /// <summary>
        /// Botón que permite la conexión del usuario con la aplicación de reserva de salas
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {            
            UserData data = new UserData();  
            //Comprobamos si el usuario ya está registrado, para mostrar o no la ventana de login
            if (data.UserExists() == false)
            {
                LoginWindow logWin = new LoginWindow();
                logWin.Show();
            }
            else
            {
                WaitForm waitForm = new WaitForm();
                waitForm.Show();
       
                Thread thread = new Thread(() => CreaReunion(waitForm));
                thread.Start();

            }
        }

        /// <summary>
        /// Función encargada de crear una nueva reunión, se conectará con el servidor y mostrará en un navegador dedicado la página de
        /// reserva de salas. A su vez, almacenará la reserva en la máquina cliente
        /// </summary>
        public void CreaReunion(WaitForm waitForm)
        {
            string path = @"C:\Users\cifua\Desktop\salida\salida.txt";
            UserData data = new UserData();
            string username, password, fullStartDate, fullEndDate, day, start, end, subject, location, id;
            bool actualizado = false;
            bool existAppointment = false;

            //Guardamos en variables todos los datos de la reunión creada en outlook
            Outlook.AppointmentItem appointment = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
            fullStartDate = appointment.Start.ToString();
            string[] date = fullStartDate.Split(' ');
            day = date[0];
            start = date[1];
            fullEndDate = appointment.End.ToString();
            date = fullEndDate.Split(' ');
            end = date[1];
            location = appointment.Location;
      
           
            //Se necesita guardar ya que, en caso de que el usuario no haya salido del campo Asunto a la hora de escribirlo, este no será capturado a no ser que se guarde la reunión
            appointment.Save();
            id = appointment.EntryID;
            subject = appointment.Subject;
            username = data.GetUsername();
            password = data.GetPassword();
            string user = "(" + username + "," + password + ")";

            ArrayList userFinal = Globals.ThisAddIn.encryptOutlook(user);
            var resultAutenticar = AutenticarUsuarioOutlook("http://88.12.10.158:81/AutenticarUsuarioOutlook", userFinal);
            JObject jsonAutenticar = JObject.Parse(resultAutenticar.Result);
            string errnoAutenticar = (string)jsonAutenticar.SelectToken("errno");
            if (errnoAutenticar.Equals("0"))
            {
                var resultReserva = GetURLCrearReservaOutlook("http://88.12.10.158:81/GetURLCrearReservaOutlook", userFinal, subject, day, start, end);
                JObject jsonReserva = JObject.Parse(resultReserva.Result);
                string errnoGetURL = (string)jsonReserva.SelectToken("errno");
                if (errnoGetURL.Equals("0"))
                {
                    string url = "";
                    if (subject is null)
                    {
                        subject = " ";
                    }
                    
                    if (data.AppointmentExists(id))
                    {
                        File.AppendAllText(path, "ACTUALIZACION DE REUNION \n");
                        ArrayList oldAppointment = data.GetAppointment(id);

                        url = "http://88.12.10.158:81/CrearReservaOutlook?user=" + userFinal.get(0).ToString() + "&password=" + userFinal.get(1).ToString() +
                              "&accion=" + 2 +"&asuntoAnterior="+oldAppointment.get(0).ToString()+ "&asuntoNuevo=" + subject + "&fechaAnterior="+oldAppointment.get(1).ToString()+
                              "&fechaNuevo=" + day + "&inicioAnterior="+oldAppointment.get(2).ToString()+"&hInicioNuevo=" + start +"&hFinAnterior="+oldAppointment.get(3).ToString()+ "&hFinNuevo=" + end;
                        existAppointment = true;
               
                    }  
                    else
                    {
                        url = "http://88.12.10.158:81/CrearReservaOutlook?user=" + userFinal.get(0).ToString() + "&password=" + userFinal.get(1).ToString() +
                                        "&accion=" + 1 + "&asuntoNuevo=" + subject + "&fechaNuevo=" + day + "&hInicioNuevo=" + start + "&hFinNuevo=" + end;              
                    }  
 
                    waitForm.Close();
                   
                    //Creamos un nuevo hilo donde se abrirá el navegador con la web de reservas de salas
                    Thread thread = new Thread(() => StartBrowser(url));
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    Thread.Sleep(1000);

                    File.AppendAllText(path, "Primer Envio: " + subject + " " + day + " " + start + " " + end + "\n");
                    var err = GetDatosReserva(userFinal, subject, day, start, end);
                    Thread.Sleep(1000);
                   File.AppendAllText(path, "Primera Respuesta: "+JObject.Parse(err.Result)+"\n"); 
                   
                    if (err.Result.Contains("\"errno\":\"1\"")) { err = GetDatosReserva(userFinal, subject, day, start, end);  }
                    File.AppendAllText(path, "Segunda Respuesta: " + JObject.Parse(err.Result) + "\n");
                    //err = GetDatosReserva(userFinal, subject, day, start, end);
                    // File.AppendAllText(path, "RECIBO: " + err.Result+"\n");
                    //Mientras se nos devuelva que no existen datos de la reserva, se preguntará continuamente
                    if (err.Result.Contains("\"errno\":\"1\""))
                    {
                        while (err.Result.Contains("\"errno\":\"1\""))
                        {
                            if (!thread.IsAlive)
                            {
                                break;
                            }                         
                            err = GetDatosReserva(userFinal, subject, day, start, end);
                            //File.AppendAllText(path, "RECIBO: " + err.Result + "\n");
                        }
                    }
                    File.AppendAllText(path, "Respuesta FINAL: " + JObject.Parse(err.Result) + "\n");
                    //Si se nos devuelven datos de una reserva, primero tendremos que comprobar que, en el caso de estar actualizando una
                    //reunión ya existente, estos datos de respuesta son distintos
                    if (err.Result.Contains("\"errno\":\"0\""))
                    {
                        if (data.AppointmentExists(id))
                        {
                            File.AppendAllText(path, "REUNION EXISTE \n");
                            err = GetDatosReserva(userFinal, subject, day, start, end);
                            Thread.Sleep(1000);
                            err = GetDatosReserva(userFinal, subject, day, start, end);
                            File.AppendAllText(path, "PRIMEROS DATOS: "+JObject.Parse(err.Result)+"\n");


                            ArrayList oldAppointment = data.GetAppointment(id);
                            string firstLocation = location;
                            JObject r = JObject.Parse(err.Result);
                            string fsede = (string)r.SelectToken("error[0].sede");
                            string fedificio = (string)r.SelectToken("error[0].edificio");
                            string fplanta = (string)r.SelectToken("error[0].planta");
                            string fsala = (string)r.SelectToken("error[0].sala");

                            while (oldAppointment.get(0).ToString().Equals((string)r.SelectToken("error[0].asunto")) & oldAppointment.get(1).ToString().Equals((string)r.SelectToken("error[0].fecha")) &
                               oldAppointment.get(2).ToString().Equals((string)r.SelectToken("error[0].hInicio")) & oldAppointment.get(3).ToString().Equals((string)r.SelectToken("error[0].hFin")) &
                                fsede.Equals((string)r.SelectToken("error[0].sede")) & fedificio.Equals((string)r.SelectToken("error[0].edificio"))
                                & fplanta.Equals((string)r.SelectToken("error[0].planta")) & fsala.Equals((string)r.SelectToken("error[0].sala")))
                            {
                                if (!thread.IsAlive)
                                {
                                    break;
                                }
                                Thread.Sleep(2000);
                                err = GetDatosReserva(userFinal, subject, day, start, end);

                                r = JObject.Parse(err.Result);
                                File.AppendAllText(path, "RECIBO: " + r + "\n");

                            }
                        }
                        File.AppendAllText(path, "HE SALIDO: " + JObject.Parse(err.Result) + "\n");
                        //Actualizamos los datos de la reunión con lo generado en la web de reservas
                        err = GetDatosReserva(userFinal, subject, day, start, end);
                        File.AppendAllText(path, "HE SALIDO Y REACTUALIZADO: " + JObject.Parse(err.Result) + "\n");
                        JObject result = JObject.Parse(err.Result);
                        string newSubject = (string)result.SelectToken("error[0].asunto");
                        string newDay = (string)result.SelectToken("error[0].fecha");
                        string newStart = (string)result.SelectToken("error[0].hInicio");
                        string newEnd = (string)result.SelectToken("error[0].hFin");
                        string sede = (string)result.SelectToken("error[0].sede");
                        string edificio = (string)result.SelectToken("error[0].edificio");
                        string planta = (string)result.SelectToken("error[0].planta");
                        string sala = (string)result.SelectToken("error[0].sala");
                        appointment.Location = "Sede: " + sede + ", " + edificio + ", " + planta + ", sala: " + sala;
                        appointment.Subject = newSubject;
                        string sd = newDay + " " + newStart;
                        DateTime startDate = DateTime.Parse(sd);
                        appointment.Start = startDate;  
                        string sd2 = newDay + " " + newEnd;
                        DateTime endDate = DateTime.Parse(sd2);
                        appointment.End = endDate;
                        appointment.Save();

                        thread.Abort();
                        thread.Join();
                        if (!subject.Equals(newSubject) | !newStart.Equals(start) | !day.Equals(newDay) | !newEnd.Equals(end))
                        {
                             MessageBox.Show("Los datos de la reunión han sido actualizados");
                        }
                        if (existAppointment == true)
                        {
                            data.UpdateAppointment(id, newSubject, newDay, newStart, newEnd);
                        }
                        else
                        {
                            //appointment.Save();
                            id = appointment.EntryID;
                            data.AddAppointment(id, newSubject, newDay, newStart, newEnd);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se han podido recopilar los datos de la reunion"); 
                        if (!existAppointment)
                        {
                            appointment.Delete();
                        }
                    }
                }
                else
                {
                    MessageBox.Show((string)jsonReserva.SelectToken("error"));
                }
            }
            else
            {
                MessageBox.Show((string)jsonAutenticar.SelectToken("error"));
            }
        }

        /// <summary>
        /// Función encargada de iniciar el navegador
        /// </summary>
        /// <param name="url"></param>
        public void StartBrowser(string url)
        {
            Browser webBrowser = new Browser(url);
            Application.Run(webBrowser);
        }

        /// <summary>
        /// Función encargada de recopilar los datos de una reserva almacenada en el servidor
        /// </summary>
        /// <param name="user">Username y password del usuario</param>
        /// <param name="subject">Asunto de la reunión</param>
        /// <param name="date">Día de la reunión</param>
        /// <param name="start">Inicio de la reunión</param>
        /// <param name="end">Finalización de la reunión</param>
        /// <returns></returns>
          async static Task<string> GetDatosReserva(ArrayList user, string subject, string date, string start, string end)
          {
            string path = @"C:\Users\cifua\Desktop\salida\salida.txt";
            using (HttpClient client = new HttpClient())
              {
              //  File.AppendAllText(path,"ENVIO: " + "http://88.12.10.158:81/GetDatosReservaOutlook?user=" + user.get(0).ToString() + "&password=" + user.get(1).ToString() + "&asunto=" + subject + "&fecha=" + date + "&hInicio=" + start + "&hFin=" + end+"\n");
                  using (HttpResponseMessage response = await client.GetAsync("http://88.12.10.158:81/GetDatosReservaOutlook?user=" + user.get(0).ToString()+"&password="+user.get(1).ToString()+"&asunto="+subject+"&fecha="+date+"&hInicio="+start+"&hFin="+end))
                  {
                      using (HttpContent content = response.Content)
                      {
                          string mycontent = await content.ReadAsStringAsync();
                          return mycontent;               
                      }
                  }
              }
          }
    
        /// <summary>
        /// Función encargada de buscar al usuario entre los registrados en la web de reservas
        /// </summary>
        /// <param name="url">URL a la que conectarse</param>
        /// <param name="user">Username y password del usuario</param>
        /// <returns></returns>
          public async Task<string> AutenticarUsuarioOutlook(string url, ArrayList user)
          {
              IEnumerable<KeyValuePair<string, string>> queries = new List<KeyValuePair<string, string>>()
              {
                  new KeyValuePair<string, string>("user",user.get(0).ToString()),
                  new KeyValuePair<string, string>("password", user.get(1).ToString())
              };
              HttpContent q = new FormUrlEncodedContent(queries);
              using (HttpClient client = new HttpClient())
              {
                  using (HttpResponseMessage response = await client.PostAsync(new Uri(url), q))
                  {
                      using (HttpContent content = response.Content)
                      {
                          string mycontent = await content.ReadAsStringAsync();
                          return mycontent;                       
                      }
                  }
              }
          }

        /// <summary>
        /// Función encargada de comprobar si los datos de una nueva reserva son los adecuados para la creación de la misma
        /// </summary>
        /// <param name="url">URL a la que conectarse</param>
        /// <param name="user">Username y password del usuario</param>
        /// <param name="subject">Asunto de la reunion</param>
        /// <param name="date">Fecha de la reunión</param>
        /// <param name="start">Hora de inicio de la reunión</param>
        /// <param name="end">Hora de finalización de la reunión</param>
        /// <returns></returns>
          async static Task <string> GetURLCrearReservaOutlook(string url, ArrayList user, string subject, string date, string start, string end)
          {
              IEnumerable<KeyValuePair<string, string>> queries = new List<KeyValuePair<string, string>>()
              {
                  new KeyValuePair<string, string>("user",user.get(0).ToString()),
                  new KeyValuePair<string, string>("password", user.get(1).ToString()),
                  new KeyValuePair<string, string>("asunto",subject),
                  new KeyValuePair<string, string>("fecha",date),
                  new KeyValuePair<string, string>("hInicio",start),
                  new KeyValuePair<string, string>("hFin", end)
              };
              HttpContent q = new FormUrlEncodedContent(queries);
              using (HttpClient client = new HttpClient())
              {
                  using (HttpResponseMessage response = await client.PostAsync(new Uri(url), q))
                  {
                      using (HttpContent content = response.Content)
                      {
                          string mycontent = await content.ReadAsStringAsync();
                          return mycontent;
                      }
                  }
              }
          }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
          {
          }
      }
  }
