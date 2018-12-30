using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Xynthesis.Utilidades
{
    public class LogXynthesis
    {
        public void EscribaLog(string modulo, string error, string user)
        {
            String path = ConfigurationManager.AppSettings["LogErrores"];
            using (StreamWriter sw = File.AppendText(path + "LOG_" + modulo + "_" + System.DateTime.Now.ToString("dd-MM-yyyy")))
            {
                sw.WriteLine("");
                sw.WriteLine("Se ha generado el siguiente Error: " + error);
                sw.WriteLine("");
                sw.WriteLine("registrado el : " + System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + " con el usuario " + user);
                sw.WriteLine("");
                sw.WriteLine("=================================================================================================");
            }
        }

        public void EscribaLog(string modulo, string log)
        {
            string path = ConfigurationManager.AppSettings["LogInformacion"];

            using (StreamWriter sw = File.AppendText(path + "LOG_" + modulo.ToUpper() + "_" + System.DateTime.Now.ToString("dd-MM-yyyy") + ".txt"))
            {
                sw.WriteLine("");
                sw.WriteLine("Se ha generado el siguiente LOG : \n" + log);
                sw.WriteLine("\n");
                sw.WriteLine("Registrado el : " + System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"));
                sw.WriteLine("");
                sw.WriteLine("=================================================================================================");
            }
        }
    }
}
