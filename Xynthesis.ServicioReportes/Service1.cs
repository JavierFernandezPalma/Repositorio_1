using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xynthesis.AccesoDatos;
using Xynthesis.Modelo;
using Xynthesis.Utilidades;

namespace Xynthesis.ServicioReportes
{
    public partial class Service1 : ServiceBase
    {
        xynthesisEntities xyt = new xynthesisEntities();
        public DateTime fechaEjecucion = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " 00:00:00");
        public string diaEjecucion = ConfigurationManager.AppSettings["diaEjecucion"] as string;
        LogXynthesis log = new LogXynthesis();
        Timer Schedular;


        public Service1()
        {
            InitializeComponent();
        }


        private void SchedularCallback(object e)
        {
            this.verificarReportes();

        }

        protected override void OnStart(string[] args)  //Inicio del servicio Metodo OnStart()
        {
            try
            {
                this.verificarReportes();
                log.EscribaLog("OnStart()", "Inicia el metodo OnStart()");
            }
            catch (Exception ex)
            {
                log.EscribaLog("OnStart()_Error", "El error es : " + ex.ToString());
            }
        }

        protected override void OnStop()
        {
            try
            {
                this.Schedular.Dispose();
                log.EscribaLog("OnStop()", "Finalizo el servicio Windows Reportes programado.");
                EventLog.WriteEntry("Finalizo el servicio Windows Reportes programado.");
            }
            catch (Exception ex)
            {
                log.EscribaLog("OnStop()_Error", "El error es : " + ex.ToString());
            }
        }
        private void LogService(string content)
        {
            string RutaLog = ConfigurationManager.AppSettings["RutaLogServiceRpt"] as string;
            FileStream fs = new FileStream(RutaLog, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(content);
            sw.Flush();
            sw.Close();
        }

        public void ServiceWindowsTest()
        {
            try
            {
                log.EscribaLog("OnStart()", "Inicia el metodo OnStart()");
                this.verificarReportes();
            }
            catch (Exception ex)
            {
                log.EscribaLog("OnStart()_Error", "El error es : " + ex.ToString());
            }
        }


        public void validarEjecRepDiarioSemanal(long idReporte)
        {
            DateTime fechaActual = DateTime.Now;
            //DateTime fechaActual = Convert.ToDateTime("2019-01-19 12:00:00"); // Para hacer pruebas
            DateTime fechPrimerDiaDelMes = new DateTime(fechaActual.Year, fechaActual.Month, 1);
            var diasDelMesactual = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);

            var reporteProgramado = (from filas in xyt.xy_configuracionrptprogramado where filas.ConfiguracionId == idReporte select filas).First();
            var horaProgramada = (from filas in xyt.xy_configuracionrptprogramado where filas.ConfiguracionId == idReporte select filas.HoraEjecucion).First();


            if (reporteProgramado.programacion == 1 && Convert.ToDateTime(fechaActual.ToString("H:mm") + ":" + "00").TimeOfDay < horaProgramada.TimeOfDay)
            {
                var a = fechaActual.ToString("yyyy-MM-dd");
                DateTime fechaProxReport = Convert.ToDateTime(fechaActual.ToString("yyyy-MM-dd") + " " + horaProgramada.ToString("H:mm:ss")); // Convierte y compacta fecha con hora para el proximo reporte
                xyt.xyp_ProxEjecucionReporte(fechaProxReport.ToString("yyyy-MM-dd H:mm:ss"), reporteProgramado.ConfiguracionId.ToString()); // Guarda fecha para proximo reporte en tabla xy_configuracionrptprogramado

            }
            else if (reporteProgramado.programacion == 1 && Convert.ToDateTime(fechaActual.ToString("H:mm") + ":" + "00").TimeOfDay > horaProgramada.TimeOfDay)
            {
                DateTime fecha = DateTime.Now.AddDays(Convert.ToInt32(1)); // Adiciona dias a la fecha actual
                DateTime hora = reporteProgramado.HoraEjecucion;
                DateTime fechaProxReport = Convert.ToDateTime(fecha.ToString("yyyy-MM-dd") + " " + hora.ToString("H:mm:ss")); // Convierte y compacta fecha con hora para el proximo reporte

                xyt.xyp_ProxEjecucionReporte(fechaProxReport.ToString("yyyy-MM-dd H:mm:ss"), reporteProgramado.ConfiguracionId.ToString()); // Guarda fecha para proximo reporte en tabla xy_configuracionrptprogramado
            }
            else
            {
                var nomPrimerDiaMes = fechPrimerDiaDelMes.ToString("dddd"); // String del nombre del día
                string nomDiaComp;
                int contadorDias = 0;

                switch (Convert.ToInt32(reporteProgramado.programacion))
                {
                    case 1:
                        nomDiaComp = "lunes";
                        break;
                    case 2:
                        nomDiaComp = "martes";
                        break;
                    case 3:
                        nomDiaComp = "miércoles";
                        break;
                    case 4:
                        nomDiaComp = "jueves";
                        break;
                    case 5:
                        nomDiaComp = "viernes";
                        break;
                    case 6:
                        nomDiaComp = "sábado";
                        break;
                    default:
                        nomDiaComp = "domingo";
                        break;

                }

                while (nomPrimerDiaMes != nomDiaComp)
                {
                    contadorDias++;
                    nomPrimerDiaMes = fechPrimerDiaDelMes.AddDays(contadorDias).ToString("dddd");
                }

                DateTime fechaComienso = Convert.ToDateTime(fechPrimerDiaDelMes.AddDays(contadorDias).ToString("yyyy-MM-dd") + " " + fechaActual.ToString("H:mm:ss"));


                while (fechaActual.Date > fechaComienso.Date)
                {
                    fechaComienso = fechaComienso.AddDays(7);
                }

                var diasParaReporte = (fechaComienso.Day - fechaActual.Day); // Diferencia de dias para el proximo reporte

                DateTime fecha = DateTime.Now.AddDays(Convert.ToInt32(diasParaReporte)); // Adiciona dias a la fecha actual
                DateTime hora = reporteProgramado.HoraEjecucion;
                DateTime fechaProxReport = Convert.ToDateTime(fecha.ToString("yyyy-MM-dd") + " " + hora.ToString("H:mm:ss")); // Convierte y compacta fecha con hora para el proximo reporte

                if (fechaActual < fechaProxReport)
                {
                    xyt.xyp_ProxEjecucionReporte(fechaProxReport.ToString("yyyy-MM-dd H:mm:ss"), reporteProgramado.ConfiguracionId.ToString()); // Guarda fecha para proximo reporte en tabla xy_configuracionrptprogramado
                }
            }

            reporteProgramado = null;
        }


        public void validarEjecucionReporteMensual(long idReporte)
        {

            DateTime fechaActual = DateTime.Now;
            //DateTime fechaActual = Convert.ToDateTime("2019-01-19 12:00:00"); // Para hacer pruebas
            DateTime fechPrimerDiaDelMes = new DateTime(fechaActual.Year, fechaActual.Month, 1);
            var diasDelMesactual = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);

            var listaRepProg = (from filas in xyt.xy_configuracionrptprogramado where filas.ConfiguracionId == idReporte select filas).ToList();


            foreach (var reporteProgramado in listaRepProg)
            {

                var diaActual = fechaActual.Day;

                if (diaActual < reporteProgramado.programacion && fechaActual.Day != diasDelMesactual)
                {
                    var diasParaReporte = reporteProgramado.programacion - diaActual;
                    DateTime fecha = DateTime.Now.AddDays(Convert.ToInt32(diasParaReporte)); // Adiciona dias a la fecha actual
                    DateTime hora = reporteProgramado.HoraEjecucion;
                    DateTime fechaProxReport = Convert.ToDateTime(fecha.ToString("yyyy-MM-dd") + " " + hora.ToString("H:mm:ss")); // Convierte y compacta fecha con hora para el proximo reporte

                    xyt.xyp_ProxEjecucionReporte(fechaProxReport.ToString("yyyy-MM-dd H:mm:ss"), reporteProgramado.ConfiguracionId.ToString()); // Guarda fecha para proximo reporte en tabla xy_configuracionrptprogramado
                }
                else if (diaActual > reporteProgramado.programacion)
                {
                    DateTime fecha = Convert.ToDateTime(DateTime.Now.AddMonths(1).ToString("yyyy-MM") + "-" + reporteProgramado.programacion); // Adiciona dias a la fecha actual
                    DateTime hora = reporteProgramado.HoraEjecucion;
                    DateTime fechaProxReport = Convert.ToDateTime(fecha.ToString("yyyy-MM-dd") + " " + hora.ToString("H:mm:ss")); // Convierte y compacta fecha con hora para el proximo reporte

                    xyt.xyp_ProxEjecucionReporte(fechaProxReport.ToString("yyyy-MM-dd H:mm:ss"), reporteProgramado.ConfiguracionId.ToString()); // Guarda fecha para proximo reporte en tabla xy_configuracionrptprogramado

                }


            }
        }

        private void verificarReportes()
        {

            try
            {
                log.EscribaLog("ejecutarReporteProgramado()", "Iniciando el metodo ejecutarReporteProgramado()");
                this.ejecutarReporteProgramado();
                log.EscribaLog("ejecutarReporteProgramado()", "Finalizo el metodo ejecutarReporteProgramado()");

                var verificarProxEjecu = (from filas in xyt.xy_configuracionrptprogramado where filas.ProxEjecucion == null select filas).ToList();

                if (verificarProxEjecu.Count != 0)
                {
                    log.EscribaLog("verificarReportes()", "Iniciando componente verificarReportes()");
                    foreach (var reporteP in verificarProxEjecu)
                    {

                        if (reporteP.TipoFrecuencia == "mensual")
                        {
                            this.validarEjecucionReporteMensual(reporteP.ConfiguracionId);
                        }
                        else
                        {
                            this.validarEjecRepDiarioSemanal(reporteP.ConfiguracionId);
                        }

                        log.EscribaLog("verificarReportes()", "Finalizo componente verificarReportes() para reporteId: " + reporteP.ConfiguracionId);
                    }

                }

                verificarProxEjecu = null;
            }

            catch (Exception ex)
            {
                log.EscribaLog("verificarReportes()_Error", "El error es : " + ex.ToString());
            }

        }

        private void ejecutarReporteProgramado()
        {

            DateTime fechaActual = DateTime.Now;
            //DateTime fechaActual = Convert.ToDateTime("2018-12-21 13:14:00"); // Para hacer pruebas
            Schedular = new Timer(new TimerCallback(SchedularCallback));
            Int64 dueTime = 0;
            TimeSpan timeSpan;


            //this.LogService("dia");
            ADBase contexto = new ADBase();

            List<xyp_SelReports_Result> lstReportes = new List<xyp_SelReports_Result>();

            var parametroUno = Convert.ToDateTime(fechaActual.ToString("yyyy-MM-dd") + " " + "00:00:00");
            var parametroDos = Convert.ToDateTime(fechaActual.ToString("yyyy-MM-dd") + " " + "23:59:59");

            List<xy_configuracionrptprogramado> listaReportesHoy = (from filasRepHoy in xyt.xy_configuracionrptprogramado where filasRepHoy.ProxEjecucion >= parametroUno && filasRepHoy.ProxEjecucion <= parametroDos orderby filasRepHoy.ProxEjecucion ascending select filasRepHoy).ToList();

            if (listaReportesHoy.Count != 0)
            {
                foreach (var reporteVerificar in listaReportesHoy)
                {
                    var idReporte = Convert.ToInt32(reporteVerificar.ConfiguracionId);
                    var horaProgramada = (from filas in xyt.xy_configuracionrptprogramado where filas.ConfiguracionId == idReporte select filas.ProxEjecucion).First();


                    if (Convert.ToDateTime(horaProgramada).TimeOfDay == Convert.ToDateTime(fechaActual.ToString("H:mm") + ":" + "00").TimeOfDay)
                    {
                        lstReportes = contexto.ObtenerProgramacionReportes(Convert.ToInt32(reporteVerificar.ConfiguracionId));

                        if (lstReportes.Count() != 0)
                        {
                            try
                            {
                                var reportesProgramados = lstReportes;
                                var intervaloFechIni = reporteVerificar.fechaInicial.ToString();
                                var intervaloFechFin = reporteVerificar.fechaFinal.ToString();

                                Xynthesis.Reportes.ExportacionReportes.GenerarReporteProgramado(reportesProgramados, intervaloFechIni, intervaloFechFin);
                                xyt.xyp_ProxEjecucionReporte(null, reporteVerificar.ConfiguracionId.ToString());
                            }
                            catch (Exception ex)
                            {
                                this.LogService("Error en ejecucion de reporte ConfiguracionID:" + reporteVerificar.ConfiguracionId + " Error: " + ex.Message + " StackTrace:" + ex.StackTrace);
                            }

                            //if (rpt != HorMin) // Se puede utilizar para validar un reporte ya evidenciado, queda pendiente si no realiza algun reporte
                            //{

                            //    DateTime fecProxRecorrido = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " " + Convert.ToDateTime(Horas[1]).ToString("H:mm:ss"));


                            //    DateTime fecProxRecorrido2 = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " " + Convert.ToDateTime("00:00:01").ToString("H:mm:ss"));
                            //    DateTime fecProxRecorrido3 = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd") + " " + Convert.ToDateTime("00:00:02").ToString("H:mm:ss"));
                            //    timeSpan = fecProxRecorrido3.Subtract(fecProxRecorrido2);
                            //    dueTime = Convert.ToInt64(timeSpan.TotalMilliseconds);
                            //    this.LogService("Proxima Fecha de ejecucion: " + fecProxRecorrido.ToString());
                            //    Schedular.Change(Convert.ToInt64(1000), Timeout.Infinite);
                            //    return;
                            //}

                        }

                    }

                }

            }

            Schedular = new Timer(new TimerCallback(SchedularCallback));
            Schedular.Change(Convert.ToInt64(10000), Timeout.Infinite);
        }

    }

}
