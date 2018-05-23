using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Data;
using System.Collections;
using ContractSender.Clases;
using System.IO;
using System.Configuration;
using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ContractSender
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(string.Format("Inicia el envío de contratos {0}", DateTime.Now.ToString()));
            
            ReportDocument rpt = new ReportDocument();
            ReportDocument rptHC = new ReportDocument();
            rpt.Load(ConfigurationManager.AppSettings["RutaReporte"]);
            rptHC.Load(ConfigurationManager.AppSettings["RutaReporteHC"]);
            
            DBAccessWin db = new DBAccessWin();
            
            DataTable dtInscripciones = new DataTable();            
            ArrayList parametros = new ArrayList();

            DataTable dtHojaCondiciones = new DataTable();

            Console.WriteLine("Obtiene lsita de inscripciones para enviar contrato");
            dtInscripciones = db.EjecutarSQLStoredProcedure("sp_ListaInscripcionesContrato", parametros);
            parametros.Clear();


            foreach (DataRow row in dtInscripciones.Rows)
            {
                DataTable dtReporte = new DataTable();
                DataTable dtSubReporte = new DataTable();
                dtReporte.Columns.Add("Nombre");
                dtReporte.Columns.Add("Representante");
                string respuesta = string.Empty;

                if (!File.Exists(string.Format("{0}Contrato{1}.pdf", ConfigurationManager.AppSettings["RutaContratos"],row["NoFamilia"].ToString())))
                {
                    DataRow reporteRow = dtReporte.NewRow();
                    reporteRow["Nombre"] = row["Nombre"].ToString();
                    reporteRow["Representante"] = row["Representante"].ToString();
                    dtReporte.Rows.Add(reporteRow);

                    rpt.SetDataSource(dtReporte);
                    Console.WriteLine(string.Format("Creando contrato: {0}", row["NoFamilia"]));
                    if (dtReporte.Rows.Count > 0)
                    {
                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, string.Format("{0}Contrato{1}.pdf", ConfigurationManager.AppSettings["RutaContratos"], row["NoFamilia"].ToString()));    
                    }                   

                    parametros.Clear();

                }

                if (!File.Exists(string.Format("{0}Caratula{1}.pdf", ConfigurationManager.AppSettings["RutaCaratulas"], row["NoFamilia"])))
                {
                    parametros.Add(row["IdMember"]);
                    dtHojaCondiciones = db.EjecutarSQLStoredProcedure("sp_BuscaDatosContrato", parametros);
                    dtSubReporte = db.EjecutarSQLStoredProcedure("sp_BuscaDatosContratoDetalle", parametros);
                    parametros.Clear();
                    rptHC.SetDataSource(dtHojaCondiciones);
                    rptHC.Subreports[0].SetDataSource(dtSubReporte);
                    if (dtHojaCondiciones.Rows.Count > 0 && dtSubReporte.Rows.Count > 0)
                    {
                        rptHC.ExportToDisk(ExportFormatType.PortableDocFormat, string.Format("{0}Caratula{1}.pdf", ConfigurationManager.AppSettings["RutaCaratulas"], row["NoFamilia"]));
                    }                    
                }


                if (File.Exists(string.Format("{0}Contrato{1}.pdf", ConfigurationManager.AppSettings["RutaContratos"], row["NoFamilia"].ToString())) && File.Exists(string.Format("{0}Caratula{1}.pdf", ConfigurationManager.AppSettings["RutaCaratulas"], row["NoFamilia"])))
                {
                    List<string> pdfs = new List<string>();
                    pdfs.Add(string.Format("{0}Caratula{1}.pdf", ConfigurationManager.AppSettings["RutaCaratulas"], row["NoFamilia"]));
                    pdfs.Add(string.Format("{0}Contrato{1}.pdf", ConfigurationManager.AppSettings["RutaContratos"], row["NoFamilia"].ToString()));

                    Merge(string.Format("{0}Contrato{1}.pdf", ConfigurationManager.AppSettings["RutaMerge"].ToString(), row["NoFamilia"]), pdfs.ToArray());

                    //Arma mensaje desde archivo

                    StreamReader objReader = new StreamReader(ConfigurationManager.AppSettings["Mensaje"]);
                    string sLine = "";
                    string msg = string.Empty;

                    while (sLine != null)
                    {
                        sLine = objReader.ReadLine();
                        if (sLine != null)
                            msg += sLine;
                    }
                    objReader.Close();

                        if (ValidacionMail(row["Email"].ToString()))
                        {
                            Console.WriteLine(string.Format("->  Enviando contrato a  {0}", row["Email"]));
                            respuesta = Message(msg, row["Email"].ToString(), string.Format("{0}Contrato{1}.pdf", ConfigurationManager.AppSettings["RutaContratos"], row["NoFamilia"].ToString())); 
                            //respuesta = Message(msg, "evaldes@sportium.com.mx", string.Format("{0}Contrato{1}.pdf", ConfigurationManager.AppSettings["RutaMerge"], row["NoFamilia"].ToString()));
                            if (respuesta == "Enviado")
                            {
                                parametros.Add(row["IdMember"]);
                                parametros.Add(1);
                                parametros.Add(string.Empty);
                            }
                            else
                            {
                                parametros.Add(row["IdMember"]);
                                parametros.Add(0);
                                parametros.Add(respuesta);
                            }
                    
                        }
                        
                        else
                        {
                            parametros.Add(row["IdMember"]);
                            parametros.Add(0);
                            parametros.Add("Email no válido");
                        }
                }   
                 else
                        {
                            parametros.Add(row["IdMember"]);
                            parametros.Add(0);
                            parametros.Add("Algunos archivos no fueron creados por falta de información");
                        }
                
                db.EjecutarSQLStoredProcedure("sp_ActualizaEnvioContrato", parametros);
                parametros.Clear();

                dtReporte.Clear();
            }

            
        }
        //Manda correos por medio de  Smtp Amazon
        public static string Message(string message, string mailTo, string attachFile)
        {
            var appSettings = ConfigurationManager.AppSettings;
            

            MailAddress fromAddress = new MailAddress(appSettings["MailFrom"], appSettings["MailFromName"]);
            MailAddress toAddress = new MailAddress(mailTo);
            

            string subject = appSettings["Subject"];
            string body = message;
            
            var smtp = new SmtpClient
            {
                Host = appSettings["Host"],
                Port = Convert.ToInt32(appSettings["Port"]),
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(appSettings["AmazonUser"], appSettings["PassMailFrom"])
            };

            var msg = new MailMessage(fromAddress, toAddress);

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(attachFile);
            msg.Attachments.Add(attachment);

            msg.Subject = subject;
            msg.Body = body;
            msg.IsBodyHtml = true;
            msg.Bcc.Add(new MailAddress(appSettings["BCCMail"]));

            try
            {
                smtp.Send(msg);
                return "Enviado";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            
            
        }

        private static Boolean ValidacionMail(String email)
        {
            string expresion;
            expresion = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
            if (email.Contains("notiene"))
            {
                return false;
            }
            if (Regex.IsMatch(email, expresion))
            {
                if (Regex.Replace(email, expresion, String.Empty).Length == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        internal static bool Merge(string strFileTarget, string[] arrStrFilesSource)
        {
            bool blnMerged = false;

            // Crea el PDF de salida
            try
            {
                using (System.IO.FileStream stmFile = new System.IO.FileStream(strFileTarget, System.IO.FileMode.Create))
                {
                    Document objDocument = null;
                    PdfWriter objWriter = null;

                    // Recorre los archivos
                    for (int intIndexFile = 0; intIndexFile < arrStrFilesSource.Length; intIndexFile++)
                    {
                        PdfReader objReader = new PdfReader(arrStrFilesSource[intIndexFile]);
                        int intNumberOfPages = objReader.NumberOfPages;

                        // La primera vez, inicializa el documento y el escritor
                        if (intIndexFile == 0)
                        { // Asigna el documento y el generador
                            objDocument = new Document(objReader.GetPageSizeWithRotation(1));
                            objWriter = PdfWriter.GetInstance(objDocument, stmFile);
                            // Abre el documento
                            objDocument.Open();
                        }
                        // Añade las páginas
                        for (int intPage = 0; intPage < intNumberOfPages; intPage++)
                        {
                            int intRotation = objReader.GetPageRotation(intPage + 1);
                            PdfImportedPage objPage = objWriter.GetImportedPage(objReader, intPage + 1);

                            // Asigna el tamaño de la página
                            objDocument.SetPageSize(objReader.GetPageSizeWithRotation(intPage + 1));
                            // Crea una nueva página
                            objDocument.NewPage();
                            // Añade la página leída
                            if (intRotation == 90 || intRotation == 270)
                                objWriter.DirectContent.AddTemplate(objPage, 0, -1f, 1f, 0, 0,
                                                                    objReader.GetPageSizeWithRotation(intPage + 1).Height);
                            else
                                objWriter.DirectContent.AddTemplate(objPage, 1f, 0, 0, 1f, 0, 0);
                        }
                    }
                    // Cierra el documento
                    if (objDocument != null)
                        objDocument.Close();
                    // Cierra el stream del archivo
                    stmFile.Close();
                }
                // Indica que se ha creado el documento
                blnMerged = true;
            }
            catch (Exception objException)
            {
                System.Diagnostics.Debug.WriteLine(objException.Message);
            }
            // Devuelve el valor que indica si se han mezclado los archivos
            return blnMerged;
        }

    }
}
