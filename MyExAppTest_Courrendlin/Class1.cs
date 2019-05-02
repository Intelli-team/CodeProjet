using System;
using System.Text;
using Micros.Ops;
using Micros.Ops.Extensibility;
using Micros.PosCore.Extensibility;
using Micros.PosCore.Extensibility.Ops;
using System.Data.SqlClient;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO.Ports;
using System.Threading;
//using MyExAppTest.WebReference;
using System.ComponentModel;
using System.Data;
//using Windows.Devices.PointOfService;

namespace MyExtensionApplication
{
    /// <summary>
    /// Implements the extension application
    /// </summary>
    public class Application : OpsExtensibilityApplication
    {
        /// <summary>
        /// Extension application Attributes
        /// </summary>
        public static int Taux_Or;	              		//Taux de Rabais pour les Cartes OR
        public static int Taux_Platine;	       		//Taux de Rabais pour les Cartes PLATINE
        public static int Taux_Premium;
        public static double Valeur_Invit_Platine;		//Valeur Invitation 2 pour les Cartes PLATINE
        public static double Valeur_Invit_Premium;		//Valeur Invitation 1 pour les Cartes PREMIUM
        public int Rabais_Or;
        public int Rabais_Platine;
        public int Rabais_Premium;
        public int Rabais_all;
        public int Rabais_Invit_Platine;
        public int Rabais_Invit_Premium;
        public double montant_Invit;
        public string cardID;
        public int PointsAcquis;
        public string saveFile;
        public int PlayerID;
        //public string CardID;
        public string Nom;
        public string Prenom;
        public int Niveau;
        public string DateValidite;                                   //Variable pour la date de validité.
        public short JourValidite;                                 //Variable pour le jour de la date de validité.
        public short MoisValidite;
        public int AnneeValidite;
        public double Coefficient;
        public string Commentaire;
        public int Invitations;
        //public int Invitations2;
        public short JourValiditeM1;
        public short MoisValiditeM1;
        public int AnneeValiditeM1;
        public int UtilInvitation;
        //public int UtilInvitation2;
        public long NouveauSolde;
        public long AncienSolde;
        public int Valeur_Point;
        public string saveRef;
        public string NomNiveau;
        public int CasId;
        public bool UseInvit;
        public bool IdentPlayer;
        /********* Tpe *******/
        static bool _continue;
        static SerialPort _serialPort;
        /******** V2 ********/
        public DateTime Deadline;
        public byte[] Picture;
        public event EventHandler FinalTender;
        // public delegate void OpsTmedEventDelegate(object sender, OpsTmedEventArgs e);
        //public event PropertyChangedEventHandler PropertyChanged;



        /// <summary>
        /// Extension application constructor
        /// </summary>
        /// <param name="context">the execution context for the application</param>
        public Application(IExecutionContext context)
            : base(context)
        {
            // TODO: Add initialization code and hook up event handlers here
            Taux_Or = 10;
            Taux_Platine = 13;
            Taux_Premium = 15;
            Valeur_Invit_Platine = 20.00;
            Valeur_Invit_Premium = 35.00;
            Rabais_Or = 5030;
            Rabais_Platine = 5050;
            Rabais_Premium = 5060;
            Rabais_all = 5061;
            montant_Invit = 0;
            cardID = null;
            PointsAcquis = 0;
            Nom = null;
            Prenom = null;
            Niveau = 0;
            DateValidite = null;                                   //Variable pour la date de validité.
            JourValidite = 0;                                 //Variable pour le jour de la date de validité.
            MoisValidite = 0;
            AnneeValidite = 0;
            Coefficient = 5; // initialiser parce que on le recupere plus par la getplayerinfo v2
            Commentaire = null;
            Invitations = 0;
            //Invitations2 = 0;
            JourValiditeM1 = 0;
            MoisValiditeM1 = 0;
            AnneeValiditeM1 = 0;
            UtilInvitation = 0;
            //UtilInvitation2 = 0;
            NouveauSolde = 0;
            AncienSolde = 0;
            Valeur_Point = 0;
            saveFile = null;
            saveRef = null;
            NomNiveau = null;
            UseInvit = false;
            IdentPlayer = false;
            Deadline = default(DateTime);
            Picture = default(byte[]);



        }

        [ExtensibilityMethod]
        public void MyExtensionMethod()
        {
            // string res=null;
            // TODO: rename method and implement
            //OpsContext.ShowMessage(string.Format("invtation Sx : {0} invitation Ent : {1} taux Or : {2}  Rabais prem nbr : {3}",Valeur_Invit_2_Platine,Valeur_Invit_1_Platine,Taux_Or, Rabais_Premium));
            //Micros.PosCore.Extensibility.Ops.ExtensibilityCardPaymentData cardPaymentData = new ExtensibilityCardPaymentData { };

            try
            {
                //PlayerAccountServiceService service = new PlayerAccountServiceService();
                //PlayerInfoDTO playerInfo = new PlayerInfoDTO();
                //PlayerPointsDTO playerPoints = new PlayerPointsDTO();
                //playerInfo = service.GetPlayerInfoByMagneticCard(cardID);
                //bool acctresspecified = false;
                //service.CreditExternalPlayerPrimeAccount(playerInfo.PlayerId, playerInfo.PlayerIdSpecified, 100, true, 1, true, 2, true, out int acctresult, out acctresspecified);
                /*service.CreditExternalPlayerPrimeAccount(PlayerID, true, 100, true, 1, true, 2, true, out int acctresult, out acctresspecified);
                OpsContext.ShowMessage(string.Format("result : {0} , result Spec : {1}", acctresult, acctresspecified));
                playerPoints = service.GetPlayerAccountBalanceByPlayerId(PlayerID, true);
                OpsContext.ShowMessage(string.Format("points : {0} , playerID : {1} , Error : {2}",playerPoints.Points,playerPoints.PlayerId,playerPoints.ErrorCode));
                service.DebitExternalPlayerPrimeAccount(PlayerID, true, 100, false, 1, false, 2, false, out int abc, out acctresspecified);
                OpsContext.ShowMessage(string.Format("result : {0} , result Spec : {1}", abc, acctresspecified));*/
                //OpsTmedEventDelegate eventDelegate;
                //this.FinalTender += testEventMethode;
                //OpsContext.PropertyChanged += TestEventMethode;
                //OpsTmedEvent += TestEventMethode;

                /* this.FinalTender += TestEventMethode;
                 if (OpsContext.LastAction == "1 Plat 15.00")
                 {
                     OnFinalTender(EventArgs.Empty);
                 }
                 */

                /*OpsContext.ShowMessage(string.Format("TransactionOrderTypeName {0}  TransactionOrderTypeNumber {1}   {2}", OpsContext.TransactionOrderTypeName,OpsContext.TransactionOrderTypeNumber,OpsContext.ConnectionStatusText));
                OpsContext.ShowMessage(string.Format("checkeventname : {0} / checkeventshortCode {1} operatoreventname : {2} / operatoreventshortCode {3}", OpsContext.CheckCurrentEventName,OpsContext.CheckCurrentEventShortCode,OpsContext.OperatorCurrentEventName,OpsContext.OperatorCurrentEventShortCode));*/

                /*OpsContext.ShowMessage("Paiement Espece");
                opsCommand.Command = OpsCommandType.TransactionOrderType;
                opsCommand.Arguments = OpsContext.CheckTotalDue;
                opsCommand.Number = 101;
                opsCommand.Index = 101;
                opsCommand.Text = "Paiement Espece";
                string res = opsCommand.ToAutoTestString();
                command = OpsContext.ProcessCommand(opsCommand);

                Micros.PosCore.Common.ResultType result = command.ResultType;
                OpsContext.ShowMessage(string.Format("result is {0}  res is : {1}", result,res));*/

                //var state = OpsContext.DataStore.GetState();

                //var tmed = OpsContext.DataStore.ReadTenderMediaByNum(3);
                //OpsContext.ShowMessage(string.Format("current local id {0} ----- getState : {1}", OpsContext.CurrentLocaleId, state));

                // OpsContext.ShowMessage(string.Format("name tender: {0} ----- number tender : {1} ---- Last Action : {2}", tmed.GetName(), tmed.GetObjectNumber(), OpsContext.LastAction));
                // PaymentEspece();
                OpsContext.ShowMessage(string.Format("Last Action is : {0}", OpsContext.LastAction));
                /*var alo = OpsContext.CurrentLocaleInfoPrint;
                OpsContext.ShowMessage(string.Format("CurrentLocaleInfoPrint : {0}",alo));
                System.IO.File.WriteAllText(@"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\PhotoExe.txt", cardID);
                Process.Start(@"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\PhotoExe");*/
                GetPhotoClient();

            }
            catch (Exception ex)
            {
                OpsContext.ShowMessage(string.Format(ex.Message));
            }
        }

        private void TestEventMethode(object sender, EventArgs e)
        {
            //Micros.PosCore.Extensibility.EventProcessingInstruction instruction = EventProcessingInstruction.Continue;
            OpsContext.ShowMessage(string.Format("Methode de Test est appelee, un plat est selectionné"));
            //return instruction;
        }

        protected virtual void OnFinalTender(EventArgs e)
        {
            EventHandler handler = FinalTender;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        /* public void OnPropertyChanged(string name)
         {
             if (this.PropertyChanged == null)
                 return;
             this.PropertyChanged((object)this, new PropertyChangedEventArgs(name));
         }*/

        [ExtensibilityMethod]
        public void PaymentEspece()
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            string printInfo = null;

            opsCommand.Command = OpsCommandType.Payment;
            opsCommand.Arguments = "Cash:Cash  ";
            opsCommand.Number = 300;
            opsCommand.Index = 300;
            if (IdentPlayer)
            {
                /***** Credit/Debit Point ***********************/
                //if (OpsContext.Check.TotalDue >= 0) {
                if ((Niveau == 1 || Niveau == 2) && UseInvit)
                {
                    DebitInvitations();
                }
                //CreditPoints();
                // }
                /* else
                 {
                     DebitPoints();
                 }*/
                /****** Post Information resultat ***********/
                printInfo = PostPlayerInformations();
                /********** Impression Post Infos **********/
                PrintPostInfoPlayer(printInfo);
                IdentPlayer = false;
            }
            /****** Lancer Payment ********************/
            command = OpsContext.ProcessCommand(opsCommand);
            InitPlayerInformation();
            /*Micros.PosCore.Common.ResultType result = command.ResultType;
            OpsContext.ShowMessage(string.Format("result is {0}", result));*/
        }

        [ExtensibilityMethod]
        public void PaymentVisa()
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            string printInfo = null;

            opsCommand.Command = OpsCommandType.Payment;
            opsCommand.Arguments = "Cash:Cash  ";
            opsCommand.Number = 110;
            opsCommand.Index = 110;
            if (IdentPlayer)
            {
                /***** Credit/Debit Point ***********************/
                //if (OpsContext.Check.TotalDue >= 0){
                if ((Niveau == 1 || Niveau == 2) && UseInvit)
                {
                    DebitInvitations();
                }
                //CreditPoints();
                //}
                /*else
                {
                    DebitPoints();
                }*/
                /****** Post Information resultat ***********/
                printInfo = PostPlayerInformations();
                /********** Impression Post Infos **********/
                PrintPostInfoPlayer(printInfo);
                IdentPlayer = false;
            }
            /****** Lancer Payment ********************/
            command = OpsContext.ProcessCommand(opsCommand);
            InitPlayerInformation();
        }

        [ExtensibilityMethod]
        public void PaymentMasterCard()
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            string printInfo = null;

            opsCommand.Command = OpsCommandType.Payment;
            opsCommand.Arguments = "Cash:Cash  ";
            opsCommand.Number = 150;
            opsCommand.Index = 150;
            if (IdentPlayer)
            {
                /***** Credit/Debit Point ***********************/
                //if (OpsContext.Check.TotalDue >= 0)
                //{
                if ((Niveau == 1 || Niveau == 2) && UseInvit)
                {
                    DebitInvitations();
                }
                //CreditPoints();
                //}
                /*else
                {
                    DebitPoints();
                }*/
                /****** Post Information resultat ***********/
                printInfo = PostPlayerInformations();
                /********** Impression Post Infos **********/
                PrintPostInfoPlayer(printInfo);
                IdentPlayer = false;
            }
            /****** Lancer Payment ********************/
            command = OpsContext.ProcessCommand(opsCommand);
            InitPlayerInformation();
        }

        [ExtensibilityMethod]
        public void PaymentMaestro()
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            string printInfo = null;

            opsCommand.Command = OpsCommandType.Payment;
            opsCommand.Arguments = "Cash:Cash  ";
            opsCommand.Number = 580;
            opsCommand.Index = 580;
            if (IdentPlayer)
            {
                /***** Credit/Debit Point ***********************/
                //if (OpsContext.Check.TotalDue >= 0)
                //{
                if ((Niveau == 1 || Niveau == 2) && UseInvit)
                {
                    DebitInvitations();
                }
                //CreditPoints();
                //}
                /*else
                {
                    DebitPoints();
                }*/
                /****** Post Information resultat ***********/
                printInfo = PostPlayerInformations();
                /********** Impression Post Infos **********/
                PrintPostInfoPlayer(printInfo);
                IdentPlayer = false;
            }
            /****** Lancer Payment ********************/
            command = OpsContext.ProcessCommand(opsCommand);
            InitPlayerInformation();
        }

        [ExtensibilityMethod]
        public void PaymentVpay()
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            string printInfo = null;

            opsCommand.Command = OpsCommandType.Payment;
            opsCommand.Arguments = "Cash:Cash  ";
            opsCommand.Number = 570;
            opsCommand.Index = 570;
            if (IdentPlayer)
            {
                /***** Credit/Debit Point ***********************/
                // if (OpsContext.Check.TotalDue >= 0)
                //{
                if ((Niveau == 1 || Niveau == 2) && UseInvit)
                {
                    DebitInvitations();
                }
                //CreditPoints();
                //}
                /*else
                {
                    DebitPoints();
                }*/
                /****** Post Information resultat ***********/
                printInfo = PostPlayerInformations();
                /********** Impression Post Infos **********/
                PrintPostInfoPlayer(printInfo);
                IdentPlayer = false;
            }
            /****** Lancer Payment ********************/
            command = OpsContext.ProcessCommand(opsCommand);
            InitPlayerInformation();
        }

        [ExtensibilityMethod]
        public void PaymentCheque()
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            string printInfo = null;

            opsCommand.Command = OpsCommandType.Payment;
            opsCommand.Arguments = "Cash:Cash  ";
            opsCommand.Number = 120;
            opsCommand.Index = 120;
            if (IdentPlayer)
            {
                /***** Credit/Debit Point ***********************/
                //if (OpsContext.Check.TotalDue >= 0)
                //{
                if ((Niveau == 1 || Niveau == 2) && UseInvit)
                {
                    DebitInvitations();
                }
                //CreditPoints();
                //}
                /*else
                {
                    DebitPoints();
                }*/
                /****** Post Information resultat ***********/
                printInfo = PostPlayerInformations();
                /********** Impression Post Infos **********/
                PrintPostInfoPlayer(printInfo);
                IdentPlayer = false;
            }
            /****** Lancer Payment ********************/
            command = OpsContext.ProcessCommand(opsCommand);
            InitPlayerInformation();
        }

        [ExtensibilityMethod]
        public void ApplyConfigSimphony()
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            string sql = null;
            string ID = null;

            try
            {
                ID = OpsContext.RequestAlphaEntry("Entrez l'Id de la configuration à appliquer", "Config Simphony");
                if (ID != null)
                {
                    connetionString = @"Data Source=V-FRPRSR7-CAPS1\MSSQLSERVER01;Database=PLAYER_BI;User ID=sa;Password=Sup3r$3cur8";
                    connection = new SqlConnection(connetionString);
                    connection.Open();

                    sql = "SELECT TauxOr,TauxPlat,TauxPrem,InvtPlaSx,InvtPlaEnt,InvtPre,RbsOr,RbsPlat,RbsPrem,RbsMntPlaSx,RbsMntPlaEnt,RbsMntPre,RbsTotal from PLAYER_BI.dbo.ConfigSimphony WHERE IdConfig=@idconf;";
                    command = new SqlCommand(sql, connection);
                    command.Parameters.Add("@idconf", System.Data.SqlDbType.VarChar);
                    command.Parameters["@idconf"].Value = ID;
                    rdr = command.ExecuteReader();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            //if (!rdr.IsDBNull(0))
                            //{
                            Taux_Or = Convert.ToInt32(rdr["TauxOr"]);
                            Taux_Platine = Convert.ToInt32(rdr["TauxPlat"]);
                            Taux_Premium = Convert.ToInt32(rdr["TauxPrem"]);
                            Valeur_Invit_Platine = Convert.ToDouble(rdr["InvtPlaEnt"]);
                            Valeur_Invit_Premium = Convert.ToDouble(rdr["InvtPre"]);
                            Rabais_Or = Convert.ToInt32(rdr["RbsOr"]);
                            Rabais_Platine = Convert.ToInt32(rdr["RbsPlat"]);
                            Rabais_Premium = Convert.ToInt32(rdr["RbsPrem"]);
                            Rabais_Invit_Platine = Convert.ToInt32(rdr["RbsMntPlaEnt"]);
                            Rabais_all = Convert.ToInt32(rdr["RbsTotal"]);

                            OpsContext.ShowMessage(string.Format("Configuration effectué avec succès"));
                            //}

                        }
                    }
                    else
                    {
                        OpsContext.ShowError(string.Format("Aucune Configuration qui correspond à l'Id : {0}", ID));
                    }

                    command.Dispose();
                    connection.Close();
                }
                else
                {
                    OpsContext.ShowError("Vous n'avez pas entrer l'id de configuration");

                }

            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }


        }

        public void PrintPostInfoPlayer(string info)
        {
            object printdata = (object)info;
            PrinterStatus status = OpsContext.SubmitPrintJob(OpsContext.GuestCheckPrinterNum, OpsContext.BackupPrinterNum, "Print info Factures", printdata);
            //OpsContext.ShowMessage(string.Format("resultat print : {0}",status));
        }

        public void ApplyDiscountByNbr(int RabaisNumber)
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;

            opsCommand.Command = OpsCommandType.Discount;
            opsCommand.Number = RabaisNumber;
            opsCommand.Index = RabaisNumber;
            opsCommand.Arguments = OpsContext.CheckTotalDue;
            command = OpsContext.ProcessCommand(opsCommand);

        }

        public string GeneratePdfFacture(string content)
        {
            string filePath = null;
            Document doc = new Document(iTextSharp.text.PageSize.LETTER, 10, 10, 45, 35);
            //filePath = @"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\" + Nom + Prenom + OpsContext.CheckNumber + ".pdf";
            filePath = @"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\" + Prenom + OpsContext.CheckNumber + ".pdf";
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.Create));
            doc.Open(); // Ouvrir Fichier
            Paragraph body = new Paragraph(content);
            doc.Add(body);
            doc.Close();

            return filePath;
        }

        [ExtensibilityMethod]
        public string PrepareTicketContent()
        {
            System.Collections.ObjectModel.ReadOnlyCollection<string> collheader = OpsContext.CheckHeaderLines;
            System.Collections.ObjectModel.ReadOnlyCollection<string> colltrailer = OpsContext.CheckTrailerLines;
            //OpsContext.ShowMessage(string.Format(coll[1]+coll[2]+coll[3]+coll[4]+coll[5]));
            StringBuilder dataprint = new StringBuilder(1600);
            for (int i = 0; i < OpsContext.CheckHeaderLinesCount; i++)
            {
                dataprint.AppendLine(collheader[i]);
            }

            dataprint.AppendLine("-----------------------------------------");
            dataprint.AppendFormat("CHK {0}                                       TBL {1}\n", OpsContext.CheckNumber, OpsContext.CheckTableNumber);
            dataprint.AppendLine("-----------------------------------------");
            long? nbrRps = OpsContext.RequestNumericEntry("Entrez le nombre de repas", "Note sans details");
            dataprint.AppendFormat("Nombre de Repas : {0}\n", nbrRps);
            dataprint.AppendLine("-----------------------------------------");
            //OpsContext.tax
            dataprint.AppendFormat("Tax: {0}\n", OpsContext.Check.Tax);
            dataprint.AppendFormat("Payment: {0}\n", OpsContext.Check.Payment);
            dataprint.AppendFormat("Sous-Total: ${0}\n", OpsContext.Check.SubTotal);
            dataprint.AppendFormat("Change Due : ${0}\n", OpsContext.CheckChangeDue);
            dataprint.AppendFormat("TOTAL DU: ${0}\n", OpsContext.Check.TotalDue);

            dataprint.AppendLine(saveFile);

            for (int j = 0; j < OpsContext.CheckTrailerLinesCount; j++)
            {
                dataprint.AppendLine(colltrailer[j]);
            }

            return dataprint.ToString();
        }

        [ExtensibilityMethod]
        public void SendFactureToMail() // Avant tout il faut activer l'option lesssecureapps pour l'adresse de l'émmeteur
        {
            string pdfTicket = null;
            string ticketContent = null;
            string emailTo = null;

            // Ticket Content
            ticketContent = PrepareTicketContent();

            bool ok = OpsContext.AskQuestion("Voulez Vous envoyer le ticket par mail ?");

            if (ok)
            {
                // search e-mail in data base 
                emailTo = GetMailClient(PlayerID);

                if (emailTo != null)
                {
                    bool okMail = OpsContext.AskQuestion("l'adresse mail trouvé est :" + emailTo + ", Est-ce-que c'est la bonne ? ");
                    if (!okMail)
                    {
                        emailTo = OpsContext.RequestAlphaEntry("Entrez l'adresse mail svp", "Mail Address");

                    }
                }
                else
                {
                    OpsContext.ShowMessage(string.Format("Pas d'adresse mail dans la base pour le client numéro : {0}", PlayerID));
                    emailTo = OpsContext.RequestAlphaEntry("Entrez l'adresse mail svp", "Mail Address");
                }

                //filePath = @"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\"+ Nom + Prenom + OpsContext.CheckNumber + ".pdf";
                //System.IO.File.WriteAllText(filePath, dataprint.ToString());
                //System.IO.File.WriteAllText(@"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\ticket", dataprint.ToString());
                pdfTicket = GeneratePdfFacture(ticketContent);

                string smtpAddress = "smtp.gmail.com";
                int portNumber = 587;
                bool enableSSL = true;

                string emailFrom = "trabelsimohamed29@gmail.com";
                string password = "**********";
                //string emailTo = "m.trabelsi@intelliway.fr";
                string subject = "Note sans details";
                //string body = dataprint.ToString();
                string body = "vous trouvez ci-joint votre ticket";

                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(emailFrom);
                    mail.To.Add(emailTo);
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.IsBodyHtml = false;
                    // Can set to false, if you are sending pure text.

                    // mail.Attachments.Add(new Attachment(@"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\ticket"));
                    mail.Attachments.Add(new Attachment(pdfTicket));

                    using (SmtpClient smtp = new SmtpClient(smtpAddress, portNumber))
                    {
                        smtp.UseDefaultCredentials = true;
                        smtp.Credentials = new NetworkCredential(emailFrom, password);
                        smtp.EnableSsl = enableSSL;
                        smtp.Send(mail);
                    }
                }
            }
            else
            {
                OpsContext.ShowMessage(string.Format(ticketContent));
                object printdata = (object)ticketContent;
                PrinterStatus status = OpsContext.SubmitPrintJob(2, OpsContext.BackupPrinterNum, "Print info Factures", printdata);
            }
        }


        [ExtensibilityMethod]
        public void GestionPointsEtRabais()
        {

            // TODO: rename method and implement
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            double TotalWithRbs = 0; // montant avec rabais
            decimal TotalRbs = 0;  // l'ensemble des rabais pour un client 
            //bool call;
            if (OpsContext.Check.Guid == "") { OpsContext.ShowError(string.Format("Une Carte Club a Deja ete Utilisee sur Cette Facture !!!")); return; }
            if (OpsContext.CheckGuestCount > 9) { OpsContext.ShowError(string.Format("Trop de Couverts pour Utiliser la Carte. 9 max.")); return; }
            if (OpsContext.RvcNumber == 6 || OpsContext.RvcNumber == 7) { OpsContext.ShowError(string.Format("Carte Non Utilisable Dans Ce Point De vente...")); return; }
            else
            {
                RechecheInfosClient();

                /* ----------------------------------------- CLIENT PRIVILIEGE -----------------------------------------------*/
                if (Niveau == 4)
                {
                    double a = (double.Parse(OpsContext.CheckTotalDue, System.Globalization.CultureInfo.InvariantCulture) / 100);
                    PointsAcquis = Convert.ToInt32(a * Coefficient);
                    InfoFacture();
                }
                /* ----------------------------------------- CLIENT OR -----------------------------------------------*/
                else if (Niveau == 3)
                {
                    double tltdue = double.Parse(OpsContext.CheckTotalDue, System.Globalization.CultureInfo.InvariantCulture) / 100;
                    //double tltdue = double.Parse(OpsContext.CheckTotalDue, System.Globalization.CultureInfo.InvariantCulture);
                    if (OpsContext.RvcNumber != 106 && OpsContext.CheckGuestCount < 3 && OpsContext.CheckDiscountZero == true)
                    {
                        /************** rabais OR ***********/
                        OpsContext.ShowMessage(string.Format("Application Rabais Or"));
                        opsCommand.Command = OpsCommandType.Discount;
                        opsCommand.Number = Rabais_Or;
                        opsCommand.Index = Rabais_Or;
                        opsCommand.Arguments = OpsContext.CheckTotalDue;
                        command = OpsContext.ProcessCommand(opsCommand);

                        /********************************/
                        tltdue = double.Parse(OpsContext.CheckTotalDue, System.Globalization.CultureInfo.InvariantCulture) / 100;
                        TotalWithRbs = tltdue - ((tltdue * Taux_Or) / 100);
                        PointsAcquis = Convert.ToInt32(Math.Abs(TotalWithRbs) * Coefficient);
                        InfoFacture();

                    }
                    else
                    {
                        tltdue = Convert.ToDouble(OpsContext.Check.TotalDue);
                        PointsAcquis = Convert.ToInt32(Math.Abs(tltdue) * Coefficient);
                        //CreditPoints();
                        //service.CreditExternalPlayerPrimeAccount(PlayerID, true, PointsAcquis, true, 1, true, 2, true, out int acctresult, out acctresspecified);
                        InfoFacture();
                    }
                }
                /* ----------------------------------------- CLIENT PLATINE -----------------------------------------------*/
                else if (Niveau == 2)
                {
                    // rvcnumber = 106 pour fouquet's
                    if (OpsContext.RvcNumber != 106 && Invitations > 0 && OpsContext.CheckDiscountZero == true)
                    {

                        GestionPlatineOrPremium(Valeur_Invit_Platine, Taux_Platine, Invitations, UtilInvitation);
                    }
                    else
                    {
                        if (OpsContext.CheckGuestCount < 3 && OpsContext.CheckDiscountZero == true)
                        {
                            /***** Application Rabais Platine *****/
                            OpsContext.ShowMessage(string.Format("Application Rabais Platine"));
                            opsCommand.Command = OpsCommandType.Discount;
                            opsCommand.Number = Rabais_Platine;
                            opsCommand.Index = Rabais_Platine;
                            opsCommand.Arguments = OpsContext.CheckTotalDue;
                            command = OpsContext.ProcessCommand(opsCommand);
                            /*******************************/
                            TotalRbs = OpsContext.Check.TotalDue - ((OpsContext.Check.TotalDue * Taux_Platine) / 100);
                            PointsAcquis = Convert.ToInt32(Math.Abs(TotalRbs) * Convert.ToDecimal(Coefficient));
                            //CreditPoints();
                            /* credit pts par PCS */
                            //service.CreditExternalPlayerPrimeAccount(PlayerID, true, PointsAcquis, true, 1, true, 2, true, out int acctresult, out acctresspecified);
                            /**********************/
                            InfoFacture();
                        }
                        else // if invit=0 et guest >= 3
                        {
                            double tltdue = Convert.ToDouble(OpsContext.Check.TotalDue);
                            PointsAcquis = Convert.ToInt32(Math.Abs(tltdue) * Coefficient);
                            InfoFacture();
                        }
                    }

                }
                /* ----------------------------------------- CLIENT PREMIUM -----------------------------------------------*/
                else if (Niveau == 1)
                {
                    if (OpsContext.RvcNumber != 106 && Invitations > 0 && OpsContext.CheckDiscountZero == true)
                    {
                        GestionPlatineOrPremium(Valeur_Invit_Premium, Taux_Premium, Invitations, UtilInvitation);
                    }
                    else
                    {
                        if (OpsContext.CheckGuestCount < 3 && OpsContext.CheckDiscountZero == true)
                        {

                            OpsContext.ShowMessage(string.Format("Application Rabais Premium"));
                            opsCommand.Command = OpsCommandType.Discount;
                            opsCommand.Number = Rabais_Premium;
                            opsCommand.Index = Rabais_Premium;
                            opsCommand.Arguments = OpsContext.CheckTotalDue;
                            command = OpsContext.ProcessCommand(opsCommand);

                            TotalRbs = OpsContext.Check.TotalDue - ((OpsContext.Check.TotalDue * Taux_Premium) / 100);
                            PointsAcquis = Convert.ToInt32(Math.Abs(TotalRbs) * Convert.ToDecimal(Coefficient));
                            //CreditPoints();
                            /* credit pts par PCS */
                            //service.CreditExternalPlayerPrimeAccount(PlayerID, true, PointsAcquis, true, 1, true, 2, true, out int acctresult, out acctresspecified);
                            /**********************/
                            InfoFacture();

                        }
                        else
                        {
                            double tltdue = Convert.ToDouble(OpsContext.Check.TotalDue);
                            PointsAcquis = Convert.ToInt32(Math.Abs(tltdue) * Coefficient);
                            //CreditPoints();
                            /* credit pts par PCS */
                            //service.CreditExternalPlayerPrimeAccount(PlayerID, true, PointsAcquis, true, 1, true, 2, true, out int acctresult, out acctresspecified);
                            /**********************/
                            InfoFacture();
                        }
                    }

                }
                else
                {
                    OpsContext.ShowError(string.Format("Niveau Membre Inconnu"));
                    return;
                }
            }

        }

        [ExtensibilityMethod]
        public void RechecheInfosClient()
        {
            string msg = null;

            //bool inputmode = OpsContext.AskQuestion("Input Mode : ? \n" + "* Mag card Swipe : Yes\n" + "* Alpha Num pop-up : No");
            //if (inputmode)
            //{
            /************************* Test Lecteur Carte magnetique ********************/
            Micros.Ops.Input.RequestEntryParameters input = new Micros.Ops.Input.RequestEntryParameters
            {
                AllowMagCard = true,
                RequestEntryType = Micros.Ops.Input.RequestEntryType.String,
                Prompt = "Passez la Carte du Membre SVP",
                Title = "Magnetic card"

            };
            Micros.Ops.Input.InputEntryResponseBase magCard = new Micros.Ops.Input.InputResponseMagCard("", "", "")
            {

                EntryMethod = Micros.Ops.Input.ResponseEntryMethod.MagCard

            };

            magCard = OpsContext.RequestEntry(input);
            Micros.Ops.Input.InputResponseMagCard inputResponse = (Micros.Ops.Input.InputResponseMagCard)magCard;
            string number = inputResponse.Track1 + inputResponse.Track2 + inputResponse.Track3;
            cardID = number.Substring(1, number.Length - 2);
            /***************************************************************************/
            //}

                //cardID = OpsContext.RequestAlphaEntry("Magnetic card", "Passez la carte");


            OpsContext.ShowMessage(string.Format(cardID));
            try
            {

                GetPlayerInformations(cardID);
                //GetPlayerInfo(cardID);
                msg = "         Informations Client         \n" + "--------------------------------------\n";
                msg = msg + "Nom Prenom   : " + Nom + " " + Prenom + "\n";
                msg = msg + "Niveau       : " + NomNiveau + "\n";
                //msg = msg + "Solde courant de points : " + AncienSolde + "\n";
                if (Niveau == 1 || Niveau == 2) { msg = msg + "Invitations : " + Invitations + "\n"; }
                msg = msg + "Commentaire  : " + Commentaire;
                msg = msg + " \n" + " \n" + "Version : V.1.9";
                OpsContext.ShowMessage(string.Format(msg));
                //GetPhotoClient(Picture);
                /* bool voir = OpsContext.AskQuestion("Voulez vous voir le solde de points ?");
                 if (voir)
                 {
                     PlayerAccountServiceService service = new PlayerAccountServiceService();
                     PlayerPointsDTO playerPoints = new PlayerPointsDTO();
                     playerPoints = service.GetPlayerAccountBalanceByPlayerId(PlayerID, true);
                     OpsContext.ShowMessage(string.Format("Solde de points : {0}",playerPoints.Points));

                 }*/

                IdentPlayer = true;
            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }

        }


        public void GestionPlatineOrPremium(double valeur, int taux, int invt, int utlInvt)
        {
            OpsCommand opsCommand = new OpsCommand();
            CommandResult command;
            bool inferieur = false;
            double TotalRbs = 0;
            //PlayerAccountServiceService service = new PlayerAccountServiceService();
            //bool acctresspecified = false;
            double tltdue = double.Parse(OpsContext.CheckTotalDue, System.Globalization.CultureInfo.InvariantCulture) / 100;
            UtilInvitation = 0;
            UseInvit = OpsContext.AskQuestion("Client a des invitations, Utiliser?");
            if (UseInvit == true)
            {
                long? inv = OpsContext.RequestNumericEntry("Saisissez le Nombre d'Invitations", "Input Invitations");
                if (inv != null)
                {
                    utlInvt = Convert.ToInt32(inv);
                    if (utlInvt < 1) { OpsContext.ShowMessage(string.Format("Nombre d'invitation invalide !!")); return; }
                    else if (utlInvt > invt) { OpsContext.ShowMessage(string.Format("Pas assez d'invitations disponibles, reste: {0} Invitations", invt)); return; }
                    else if (utlInvt > OpsContext.CheckGuestCount) { OpsContext.ShowMessage(string.Format("Trop d'invitations pour {0} Couverts !!", OpsContext.CheckGuestCount)); return; }
                    else
                    {
                        montant_Invit = valeur * utlInvt;
                        if (tltdue + Convert.ToDouble(OpsContext.CheckDiscount) < (montant_Invit - valeur))
                        {
                            OpsContext.ShowError(string.Format("Trop d'invitation pour ce total facture !"));
                            return;
                        }
                        if (tltdue < montant_Invit)
                        {
                            bool ok = OpsContext.AskQuestion(string.Format("Total = {0} et Invit = {1} , Continuer ?", tltdue, valeur));
                            if (ok != true)
                            {
                                return;
                            }
                            else
                            {
                                /************** Calcul Rabais 100% ***********/
                                inferieur = true;
                                OpsContext.ShowMessage(string.Format("Application rabais"));
                                opsCommand.Command = OpsCommandType.Discount;
                                opsCommand.Number = 5061;
                                opsCommand.Index = 5061;
                                opsCommand.Arguments = OpsContext.CheckTotalDue;
                                command = OpsContext.ProcessCommand(opsCommand);
                            }
                        }
                        /************** Calcul Rabais Invitation ***********/
                        if (!inferieur)
                        {
                            if (Niveau == 1)
                            {
                                OpsContext.ShowMessage(string.Format("Application rabais montant invitation"));
                                opsCommand.Command = OpsCommandType.Discount;
                                opsCommand.Arguments = OpsContext.CheckTotalDue;
                                switch (utlInvt)
                                {
                                    case 1:
                                        opsCommand.Number = 5040;
                                        opsCommand.Index = 5040;
                                        break;
                                    case 2:
                                        opsCommand.Number = 5041;
                                        opsCommand.Index = 5041;
                                        break;
                                    case 3:
                                        opsCommand.Number = 5042;
                                        opsCommand.Index = 5042;
                                        break;
                                    case 4:
                                        opsCommand.Number = 5043;
                                        opsCommand.Index = 5043;
                                        break;
                                    default:
                                        OpsContext.ShowError(string.Format("Vous pouvez pas utiliser plus de 4 invitations"));
                                        break;
                                }

                                command = OpsContext.ProcessCommand(opsCommand);
                            }
                            else
                            {

                                OpsContext.ShowMessage(string.Format("Application rabais montant invitation"));
                                opsCommand.Command = OpsCommandType.Discount;
                                opsCommand.Arguments = OpsContext.CheckTotalDue;
                                switch (utlInvt)
                                {
                                    case 1:
                                        opsCommand.Number = 5062;
                                        opsCommand.Index = 5062;
                                        break;
                                    case 2:
                                        opsCommand.Number = 5063;
                                        opsCommand.Index = 5063;
                                        break;
                                    case 3:
                                        opsCommand.Number = 5064;
                                        opsCommand.Index = 5064;
                                        break;
                                    case 4:
                                        opsCommand.Number = 5065;
                                        opsCommand.Index = 5065;
                                        break;
                                    default:
                                        OpsContext.ShowError(string.Format("Vous pouvez pas utiliser plus de 4 invitations"));
                                        break;
                                }
                                command = OpsContext.ProcessCommand(opsCommand);

                            }
                        }
                        UtilInvitation = utlInvt;
                        if (montant_Invit == tltdue + Convert.ToDouble(OpsContext.CheckDiscount))
                        {
                            PointsAcquis = 0;
                        }
                        else
                        {
                            if (!inferieur)
                            {
                                tltdue = Convert.ToDouble(OpsContext.Check.TotalDue);
                                PointsAcquis = Convert.ToInt32(Math.Abs(tltdue) * Coefficient);
                            }
                            else { PointsAcquis = 0; }

                        }
                    }

                }

            }
            else
            {
                if (OpsContext.CheckGuestCount < 3 && OpsContext.CheckDiscountZero == true) // + DSCP[] rabais
                {
                    if (Niveau == 2)
                    {
                        OpsContext.ShowMessage(string.Format("Application rabais Platine"));
                        opsCommand.Command = OpsCommandType.Discount;
                        opsCommand.Number = Rabais_Platine;
                        opsCommand.Index = Rabais_Platine;
                        opsCommand.Arguments = OpsContext.CheckTotalDue;
                        //opsCommand.Index = 11;
                        command = OpsContext.ProcessCommand(opsCommand);
                    }
                    else if (Niveau == 1)
                    {
                        OpsContext.ShowMessage(string.Format("Application rabais Premium"));
                        opsCommand.Command = OpsCommandType.Discount;
                        opsCommand.Number = Rabais_Premium;
                        opsCommand.Index = Rabais_Premium;
                        opsCommand.Arguments = OpsContext.CheckTotalDue;
                        //opsCommand.Index = 11;
                        command = OpsContext.ProcessCommand(opsCommand);
                    }
                    TotalRbs = tltdue - ((tltdue * taux) / 100);
                    PointsAcquis = Convert.ToInt32(Math.Abs(TotalRbs) * Coefficient);
                    InfoFacture();

                }
                else
                {
                    tltdue = Convert.ToDouble(OpsContext.Check.TotalDue);
                    PointsAcquis = Convert.ToInt32(Math.Abs(tltdue) * Coefficient);
                    InfoFacture();
                }
            }
        }

        [ExtensibilityMethod]
        public void LectureCardEtPaiementPoints() // TMED Event
        {
            int i;
            double use_amount = 0;
            bool after_void = false;

            if (OpsContext.TrainingModeEnabled == true)    // Verification Mode Formation
            {
                OpsContext.ShowError(string.Format("Mode Formation Interdit"));
                return;
            }
            if (OpsContext.CheckIsOpen != true) // Facture ouverte ?
            {
                OpsContext.ShowError(string.Format("Pas de Facture Ouverte, Impossible de Continuer"));
                return;
            }
            try
            {

                for (i = 0; i < OpsContext.CheckDetail.Count - 2; i++) // OR OpsContext.CheckContext.CheckDetail.Count
                {
                    if (OpsContext.CheckDetail[i].DetailIndex == 1)
                    {
                        if (OpsContext.CheckDetail[i].DetailType.ToString() == "T")
                        {
                            if (OpsContext.CheckDetail[i].Status == "Paiement_Points")
                            {
                                if (OpsContext.CheckTrailerLinesCount >= i + 1)
                                {
                                    if (OpsContext.CheckDetail[i + 1].DetailType.ToString() == "R" && OpsContext.CheckDetail[i + 1].Name != "")
                                    {
                                        PlayerID = Convert.ToInt32(OpsContext.CheckDetail[i + 1].Name);
                                        string str = OpsContext.CheckDetail[i + 2].Name.Substring(14, 10);
                                        PointsAcquis = Convert.ToInt32(str);
                                        PointsAcquis = -PointsAcquis;
                                        DebitPoints();
                                        after_void = true;

                                    }
                                }
                            }
                        }
                    }
                }

                if (after_void == false)
                {
                    cardID = OpsContext.RequestAlphaEntry("Passez la Carte du Membre SVP", "Magnetic Code");
                    GetPlayerInformations(cardID);
                    //OpsContext.ShowMessage(string.Format(res));
                    if (Convert.ToDouble(OpsContext.CheckSubtotal) != 0)
                    {
                        use_amount = Convert.ToDouble(OpsContext.CheckSubtotal);
                    }
                    else
                    {
                        use_amount = double.Parse(OpsContext.CheckTotalDue, System.Globalization.CultureInfo.InvariantCulture) / 100;
                        //use_amount = double.Parse(OpsContext.CheckTotalDue, System.Globalization.CultureInfo.InvariantCulture);
                    }

                    PointsAcquis = Convert.ToInt32((use_amount * Valeur_Point) / 100);

                    if (AncienSolde < PointsAcquis)
                    {
                        OpsContext.ShowError(string.Format("Pas Assez de Points, Manque: {0} Paiement Maximum: CHF {1}", PointsAcquis - AncienSolde, AncienSolde / Valeur_Point));
                        return;
                    }
                    DebitPoints();
                }

                saveRef = PlayerID.ToString() + "\n";
                saveRef = saveRef + "Pts Utilises : " + PointsAcquis.ToString();
                saveRef = saveRef + "Pts Restant : " + NouveauSolde.ToString();

            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }


        }

        /*******************************************************************************/
        // Informations a imprimer sur la fature				//
        /*******************************************************************************/
        [ExtensibilityMethod]
        public void InfoFacture()
        {

            try
            {

                saveFile = "============================\n" + "Club Barriere Courrendlin\n" + "============================\n";
                //saveFile = saveFile + Nom + "\n" + Prenom + "\n" + "Solde Actuel de Points: " + AncienSolde.ToString() + "\n";
                saveFile = saveFile + Nom + "\n" + Prenom + "\n" + "Niveau : " + NomNiveau + "\n";
                saveFile = saveFile + "Pas de credit point pour cette version \n";
                //saveFile = saveFile + "Points Acquis: " + PointsAcquis.ToString() + "\n" + "Nouveau Solde Points: " + NouveauSolde.ToString() + "\n";

                if (Niveau == 1 || Niveau == 2)
                {
                    if ((Invitations - UtilInvitation) > 0)
                    {

                        saveFile = saveFile + "Solde Invitations : " + (Invitations - UtilInvitation).ToString() + "\n";
                        DateValidite = JourValiditeM1.ToString() + "." + MoisValiditeM1.ToString() + "." + AnneeValiditeM1.ToString();
                        saveFile = saveFile + "Validite : " + DateValidite + "\n";
                    }
                    else
                    {
                        saveFile = saveFile + "Plus d'invitations\n";
                        saveFile = saveFile + "Nouvelles disponibles\n";
                        DateValidite = JourValidite.ToString() + "." + MoisValidite.ToString() + "." + AnneeValidite.ToString();
                        saveFile = saveFile + "Des le : " + DateValidite + "\n";
                    }
                }
                saveFile = saveFile + "============================";

                OpsContext.ShowMessage(string.Format(saveFile));
                /*object printdata = (object)saveFile;
                PrinterStatus status = OpsContext.SubmitPrintJob(2, OpsContext.BackupPrinterNum, "Print info Factures", printdata);
                OpsContext.ShowMessage(string.Format("le status d'impression est : {0}", status));*/

            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }

        }

        [ExtensibilityMethod]
        public void GetAncienSolde()
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            string sql = null;
            //int soldeAct = 0;

            try
            {
                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=winoasis;User ID=microssvc;Password=microssvc; MultipleActiveResultSets=True";
                connection = new SqlConnection(connetionString);
                connection.Open();

                sql = "SELECT PtsMgrAdjusted FROM winoasis.dbo.CDS_STATBALANCE WHERE Meta_ID=@nocli;";
                command = new SqlCommand(sql, connection);
                command.Parameters.Add("@nocli", System.Data.SqlDbType.Int);
                command.Parameters["@nocli"].Value = PlayerID;
                rdr = command.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        AncienSolde = Convert.ToInt32(rdr["PtsMgrAdjusted"]);

                    }
                }

                command.Dispose();
                connection.Close();

            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }

        }

        public string GetMailClient(int noOCM)
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            string sql = null;
            string mail = null;

            try
            {
                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=CasinoCore;User ID=microssvc;Password=microssvc";
                connection = new SqlConnection(connetionString);
                connection.Open();

                sql = "SELECT email from CasinoCore.dbo.PlayerBannedContactData WHERE noOCM=@nocli;";
                command = new SqlCommand(sql, connection);
                command.Parameters.Add("@nocli", System.Data.SqlDbType.Int);
                command.Parameters["@nocli"].Value = PlayerID;
                rdr = command.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        if (!rdr.IsDBNull(0))
                        {
                            mail = (rdr["email"]).ToString();
                        }

                    }
                }
                else
                {
                    mail = null;
                }

                command.Dispose();
                connection.Close();

            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }

            return mail;

        }

        [ExtensibilityMethod]
        public bool IdentificationClient()
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command, command1;
            SqlDataReader rdr, rdr1;
            string sql = null, sql1 = null;
            string name = null, firstname = null;
            bool rplc, exist = false;
            int noCli = 0, casinoID = 0, niv = 0;

            OpsContext.ShowMessage(string.Format("Identification du client le propriétaire de la carte magnetique numéro : {0}", cardID));

            try
            {

                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=pos;User ID=microssvc;Password=microssvc; MultipleActiveResultSets=True";
                connection = new SqlConnection(connetionString);
                connection.Open();

                sql = "SELECT CASINO_ID,NOCLI FROM pos.dbo.HISTCARD WHERE MAGNETIC=@mgn;";
                command = new SqlCommand(sql, connection);
                command.Parameters.Add("@mgn", System.Data.SqlDbType.VarChar);
                command.Parameters["@mgn"].Value = cardID;
                rdr = command.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        casinoID = Convert.ToInt32(rdr["CASINO_ID"]);
                        noCli = Convert.ToInt32(rdr["NOCLI"]);

                    }
                }
                else
                {
                    OpsContext.ShowMessage(string.Format("Aucun client correspond à la carte magnetique num : {0} ", cardID));
                    return false;
                }
                command.Dispose();
                connection.Close();

                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=CasinoCore;User ID=microssvc;Password=microssvc; MultipleActiveResultSets=True";
                connection = new SqlConnection(connetionString);
                connection.Open();

                sql = "SELECT cardLevelID, lastName, firstName FROM CasinoCore.dbo.V_PlayerAccount WHERE noOCM = @noCLi;";
                command = new SqlCommand(sql, connection);
                command.Parameters.Add("@noCLi", System.Data.SqlDbType.Int);
                command.Parameters["@noCLi"].Value = noCli;
                rdr = command.ExecuteReader();
                while (rdr.Read())
                {
                    //guid = rdr[0].ToString();
                    niv = Convert.ToInt32(rdr[0]);
                    name = rdr[1].ToString();
                    firstname = rdr[2].ToString();
                }

                command.Dispose();
                connection.Close();

                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=PLAYER_BI;User ID=microssvc;Password=microssvc; MultipleActiveResultSets=True";
                connection = new SqlConnection(connetionString);
                connection.Open();
                string guid = null;
                int checknum = 0, revctrId = 0;


                sql = "SELECT CheckNumber,GUID ,REVCTRID FROM PLAYER_BI.dbo.Player_DTL WHERE NOCLI = @ncli AND MAGCARD=@mgn; ";

                command = new SqlCommand(sql, connection);
                command.Parameters.Add("@ncli", System.Data.SqlDbType.Int);
                command.Parameters["@ncli"].Value = noCli;
                command.Parameters.Add("@mgn", System.Data.SqlDbType.VarChar);
                command.Parameters["@mgn"].Value = cardID;
                rdr = command.ExecuteReader();
                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        //guid = rdr[0].ToString();
                        checknum = Convert.ToInt32(rdr["CheckNumber"]);
                        guid = rdr["GUID"].ToString();
                        revctrId = Convert.ToInt32(rdr["REVCTRID"]);
                    }
                }
                else
                {
                    OpsContext.ShowMessage(string.Format("Aucune facture est assignée au client {0}", noCli));
                    return false;
                }
                command.Dispose();

                if (checknum == OpsContext.CheckNumber && guid == OpsContext.Check.Guid && revctrId == OpsContext.Check.RevCtrID.Value)
                {
                    rplc = OpsContext.AskQuestion(string.Format("Une Carte Club a Deja ete Utilisee sur Cette Facture ! Voulez Vous la remplacer ?"));
                    if (rplc == true)
                    {
                        sql1 = "UPDATE PLAYER_BI.dbo.Player_DTL set Firstname=@prenom, Lastname=@name, CASINOID=@casid, LevelCardId=@lvl, MAGCARD=@mgn WHERE NOCLI=@ncli AND CheckNumber=@checknum;";
                        command1 = new SqlCommand(sql1, connection);
                        command1.Parameters.Add("@prenom", System.Data.SqlDbType.VarChar);
                        command1.Parameters["@prenom"].Value = firstname;
                        command1.Parameters.Add("@name", System.Data.SqlDbType.VarChar);
                        command1.Parameters["@name"].Value = name;
                        command1.Parameters.Add("@casid", System.Data.SqlDbType.Int);
                        command1.Parameters["@casid"].Value = casinoID;
                        command1.Parameters.Add("@lvl", System.Data.SqlDbType.Int);
                        command1.Parameters["@lvl"].Value = niv;
                        command1.Parameters.Add("@mgn", System.Data.SqlDbType.VarChar);
                        command1.Parameters["@mgn"].Value = cardID;
                        command1.Parameters.Add("@ncli", System.Data.SqlDbType.Int);
                        command1.Parameters["@ncli"].Value = noCli;
                        command1.Parameters.Add("@checknum", System.Data.SqlDbType.Int);
                        command1.Parameters["@checknum"].Value = OpsContext.CheckNumber;
                        rdr1 = command1.ExecuteReader();

                        command1.Dispose();

                    }
                    exist = true;
                }
                else
                {
                    sql1 = "INSERT INTO PLAYER_BI.dbo.Player_DTL (CASINOID,NOCLI,Firstname,Lastname,LevelCardId,MAGCARD,CheckNumber,GUID,REVCTRID) VALUES (@casid, @ncli, @prenom, @name, @lvl, @mgn, @checknum, @guid, @revid);";
                    command1 = new SqlCommand(sql1, connection);
                    command1.Parameters.Add("@casid", System.Data.SqlDbType.Int);
                    command1.Parameters["@casid"].Value = casinoID;
                    command1.Parameters.Add("@ncli", System.Data.SqlDbType.Int);
                    command1.Parameters["@ncli"].Value = noCli;
                    command1.Parameters.Add("@prenom", System.Data.SqlDbType.VarChar);
                    command1.Parameters["@prenom"].Value = name;
                    command1.Parameters.Add("@name", System.Data.SqlDbType.VarChar);
                    command1.Parameters["@name"].Value = name;
                    command1.Parameters.Add("@lvl", System.Data.SqlDbType.Int);
                    command1.Parameters["@lvl"].Value = niv;
                    command1.Parameters.Add("@mgn", System.Data.SqlDbType.VarChar);
                    command1.Parameters["@mgn"].Value = cardID;
                    command1.Parameters.Add("@checknum", System.Data.SqlDbType.Int);
                    command1.Parameters["@checknum"].Value = OpsContext.CheckNumber;
                    command1.Parameters.Add("@guid", System.Data.SqlDbType.VarChar);
                    command1.Parameters["@guid"].Value = OpsContext.Check.Guid;
                    command1.Parameters.Add("@revid", System.Data.SqlDbType.Int);
                    command1.Parameters["@revid"].Value = OpsContext.Check.RevCtrID.Value;
                    rdr1 = command1.ExecuteReader();

                    command1.Dispose();
                    OpsContext.ShowMessage(string.Format("Le client {0} est associé à la facture nuémro : {1}", noCli, OpsContext.CheckNumber));
                    exist = false;
                }

                connection.Close();

            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }

            return exist;
        }

        [ExtensibilityMethod]
        public string PostPlayerInformations()
        {
            string sql = null;
            string connetionString = null;
            SqlDataReader rdr;
            SqlCommand command1;
            SqlConnection connection;
            int invitRes = 0;
            StringBuilder builder = new StringBuilder(600);
            builder.AppendLine("==================================");
            builder.AppendLine("      Club Barriere Courrendlin      ");
            builder.AppendLine("==================================");
            builder.AppendLine(Nom + " " + Prenom);
            builder.AppendLine(NomNiveau);
            /*if (OpsContext.Check.TotalDue >= 0)
            {
                builder.AppendFormat("Ancien solde points : {0}\n", AncienSolde);
                builder.AppendFormat("Point acquis : {0}\n", PointsAcquis);
                builder.AppendFormat("Nouveau solde points : {0}\n", NouveauSolde);
            }
            else
            {
                builder.AppendFormat("Ancien solde points : {0}\n", AncienSolde);
                builder.AppendLine("Point acquis : 0");
            }*/
            if (Niveau == 1 || Niveau == 2)
            {
                if ((Invitations - UtilInvitation) > 0)
                {
                    invitRes = Invitations - UtilInvitation;
                    builder.AppendFormat("Solde Invitations : {0}\n", invitRes);
                    DateValidite = JourValiditeM1.ToString() + "." + MoisValiditeM1.ToString() + "." + AnneeValiditeM1.ToString();
                    builder.AppendLine("Validite : " + DateValidite);
                }
                else
                {
                    builder.AppendLine("Plus d'invitations");
                    builder.AppendLine("Nouvelles disponibles");
                    DateValidite = JourValidite.ToString() + "." + MoisValidite.ToString() + "." + AnneeValidite.ToString();
                    builder.AppendLine("Des le : " + DateValidite);
                }
            }
            builder.AppendLine("-----------------------------------------");
            //ExtensibilityDataInfo dataInfo = new ExtensibilityDataInfo("", "", "");
            //dataInfo.Name = builder.ToString();
            OpsContext.ShowMessage(string.Format(builder.ToString()));
            //OpsContext.Check.AddExtensibilityData(dataInfo);

            try
            {

                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=PLAYER_BI;User ID=microssvc;Password=microssvc";
                connection = new SqlConnection(connetionString);
                connection.Open();


                sql = "INSERT INTO PLAYER_BI.dbo.Player_DTL (NOCLI,CASINOID,Firstname,Lastname,LevelCardId,MAGCARD,CheckNumber,GUID,REVCTRID) VALUES (@nocli,@casid,@prenom,@name,@lvl,@mgn,@checknum,@guid,@rectrid);";
                command1 = new SqlCommand(sql, connection);
                command1.Parameters.Add("@nocli", System.Data.SqlDbType.Int);
                command1.Parameters["@nocli"].Value = PlayerID;
                command1.Parameters.Add("@casid", System.Data.SqlDbType.Int);
                command1.Parameters["@casid"].Value = CasId;
                command1.Parameters.Add("@prenom", System.Data.SqlDbType.VarChar);
                command1.Parameters["@prenom"].Value = Prenom;
                command1.Parameters.Add("@name", System.Data.SqlDbType.VarChar);
                command1.Parameters["@name"].Value = Nom;
                command1.Parameters.Add("@lvl", System.Data.SqlDbType.Int);
                command1.Parameters["@lvl"].Value = Niveau;
                command1.Parameters.Add("@mgn", System.Data.SqlDbType.VarChar);
                command1.Parameters["@mgn"].Value = cardID;
                command1.Parameters.Add("@checknum", System.Data.SqlDbType.Int);
                command1.Parameters["@checknum"].Value = OpsContext.CheckNumber;
                command1.Parameters.Add("@guid", System.Data.SqlDbType.VarChar);
                command1.Parameters["@guid"].Value = OpsContext.Check.Guid;
                command1.Parameters.Add("@rectrid", System.Data.SqlDbType.Int);
                command1.Parameters["@rectrid"].Value = OpsContext.Check.RevCtrID.Value;
                rdr = command1.ExecuteReader();

                command1.Dispose();
                connection.Close();
            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }

            return builder.ToString();

        }

        /**********************************************************************/
        // Recherche des informations sur le membre			//
        /*********************************************************************/
        public void GetPlayerInformations(string card)
        {
            string connetionString = null;
            // string chaine = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            // string[] chainSplt = new string[16];
            string sql = null;
            //connetionString = @"Data Source=V-FRPRSR7-CAPS1\SQLEXPRESS;Database=CasinoCore;User ID=sa;Password=Sup3r$3cur8";
            connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=CasinoCore;User ID=microssvc;Password=microssvc; MultipleActiveResultSets=True";
            connection = new SqlConnection(connetionString);
            connection.Open();
            //sql = "SELECT NOCLI,CASINO_ID FROM pos.dbo.HISTCARD WHERE ACTIVE = 1 AND MAGNETIC = @magneticCode;";
            sql = "Execute CasinoCore.dbo.Micros_GetPlayerInformations @magneticCode;";
            command = new SqlCommand(sql, connection);
            command.Parameters.Add("@magneticCode", System.Data.SqlDbType.Char);
            command.Parameters["@magneticCode"].Value = card;
            rdr = command.ExecuteReader();
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    PlayerID = Convert.ToInt32(rdr[0]);
                    Nom = rdr[1].ToString();
                    Prenom = rdr[2].ToString();
                    Niveau = Convert.ToInt16(rdr[3]);
                    if (!rdr.IsDBNull(3) && !rdr.IsDBNull(4) && !rdr.IsDBNull(5) && !rdr.IsDBNull(6))
                    {
                        Invitations = Convert.ToInt32(rdr[4]);
                        JourValidite = Convert.ToInt16(rdr[5]);
                        MoisValidite = Convert.ToInt16(rdr[6]);
                        AnneeValidite = Convert.ToInt32(rdr[7]);
                    }
                    Coefficient = Convert.ToDouble(rdr[8]);
                    if (!rdr.IsDBNull(9)) { Commentaire = rdr[9].ToString(); }
                    if (!rdr.IsDBNull(10) && !rdr.IsDBNull(11) && !rdr.IsDBNull(12))
                    {
                        JourValiditeM1 = Convert.ToInt16(rdr[10]);
                        MoisValiditeM1 = Convert.ToInt16(rdr[11]);
                        AnneeValiditeM1 = Convert.ToInt16(rdr[12]);
                    }
                    AncienSolde = Convert.ToInt64(rdr[13]);
                    Valeur_Point = Convert.ToInt32(rdr[14]);


                }
            }
            else
            {

                OpsContext.ShowMessage(string.Format("Aucun client ne correspond à cette carte magnetique"));
                InitPlayerInformation();
                return;
            }
            command.Dispose();

            //GetAncienSolde();

            switch (Niveau)
            {
                case 1:
                    NomNiveau = "Premium";
                    break;
                case 2:
                    NomNiveau = "Platine";
                    break;
                case 3:
                    NomNiveau = "Or";
                    break;
                case 4:
                    NomNiveau = "Privilege";
                    break;
                default:
                    NomNiveau = "Pas de Niveau";
                    break;
            }

            connection.Close();


        }

        public void InitPlayerInformation()
        {

            Nom = null;
            Prenom = null;
            Niveau = 0;
            Invitations = 0;
            JourValidite = 0;
            MoisValidite = 0;
            AnneeValidite = 0;
            Commentaire = null;
            PointsAcquis = 0;
            AncienSolde = 0;
            JourValiditeM1 = 0;
            MoisValiditeM1 = 0;
            AnneeValiditeM1 = 0;
            UseInvit = false;

        }

        public void GetPlayerInfo(string card)
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            string sql = null;
            //connetionString = @"Data Source=V-FRPRSR7-CAPS1\SQLEXPRESS;Database=CasinoCore;User ID=sa;Password=Sup3r$3cur8";
            connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=CasinoCore;User ID=microssvc;Password=microssvc; MultipleActiveResultSets=True";
            connection = new SqlConnection(connetionString);
            connection.Open();
            //sql = "SELECT NOCLI,CASINO_ID FROM pos.dbo.HISTCARD WHERE ACTIVE = 1 AND MAGNETIC = @magneticCode;";
            sql = "Execute CasinoCore.dbo.Micros_GetPlayerInfo @magneticCode;";
            command = new SqlCommand(sql, connection);
            command.Parameters.Add("@magneticCode", System.Data.SqlDbType.Char);
            command.Parameters["@magneticCode"].Value = card;
            rdr = command.ExecuteReader();
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    PlayerID = Convert.ToInt32(rdr[0]);
                    Nom = rdr[1].ToString();
                    Prenom = rdr[2].ToString();
                    Niveau = Convert.ToInt16(rdr[3]);
                    if (!rdr.IsDBNull(4) && !rdr.IsDBNull(5) && !rdr.IsDBNull(6))
                    {
                        Deadline = Convert.ToDateTime(rdr[4]);
                        Invitations = Convert.ToInt32(rdr[5]);
                        Picture = (byte[])rdr[6];
                    }
                }
            }
            else
            {

                OpsContext.ShowMessage(string.Format("Aucun client ne correspond à cette carte magnetique"));
                Nom = null;
                Prenom = null;
                Niveau = 0;
                Invitations = 0;
                Deadline = default(DateTime);
                Picture = default(byte[]);
                PointsAcquis = 0;
                return;
            }
            command.Dispose();

            //GetAncienSolde();

            switch (Niveau)
            {
                case 1:
                    NomNiveau = "Premium";
                    break;
                case 2:
                    NomNiveau = "Platine";
                    break;
                case 3:
                    NomNiveau = "Or";
                    break;
                case 4:
                    NomNiveau = "Privilege";
                    break;
                default:
                    NomNiveau = "Pas de Niveau";
                    break;
            }

            connection.Close();
        }

        /***********************************************************************/
        // Utilisation de points sur le membre				//
        /***********************************************************************/
        [ExtensibilityMethod]
        public void DebitPoints()
        {
            string connetionString = null;
            string result = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            string sql = null;
            string fct = null;
            fct = OpsContext.RvcName + OpsContext.CheckNumber; //Formatage du texte pour la reference facture.
            connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=CasinoCore;User ID=microssvc;Password=microssvc";
            sql = "Execute CasinoCore.dbo.Micros_DebitPoints @noOCM , @nbrpts , @fct;"; // la procedure stockée à executer
            connection = new SqlConnection(connetionString); //Connexion à la bese de données 

            connection.Open();
            command = new SqlCommand(sql, connection);
            command.Parameters.Add("@noOCM", System.Data.SqlDbType.Int);
            command.Parameters.Add("@nbrpts", System.Data.SqlDbType.Int);
            command.Parameters.Add("@fct", System.Data.SqlDbType.VarChar);
            command.Parameters["@noOCM"].Value = PlayerID;
            command.Parameters["@nbrpts"].Value = PointsAcquis;
            command.Parameters["@fct"].Value = fct;
            rdr = command.ExecuteReader();
            while (rdr.Read())
            {
                NouveauSolde = Convert.ToInt64(rdr[0]);
                /*if (!rdr.IsDBNull(1))
                {
                    Invitations = Convert.ToInt32(rdr[1]);
                }
                if (!rdr.IsDBNull(2))
                {
                    Invitations2 = Convert.ToInt32(rdr[2]);
                }*/

            }
            command.Dispose();
            connection.Close();
            result = "Solde de points: " + NouveauSolde + "\nNombre d'invitation: " + (Invitations - UtilInvitation);

            OpsContext.ShowMessage(string.Format(result));


        }

        public void GetPhotoClient()
        {

            /*System.IO.File.WriteAllText(@"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\PhotoExe.txt", cardID);
            Process.Start(@"C:\Micros\Simphony\WebServer\wwwroot\EGateway\Handlers\PhotoExe");*/
            string connetionString = null;
            byte[] image = new byte[0]; ;
            SqlConnection connection;
            SqlCommand command;
            //SqlDataReader rdr;
            string sql = null;

            connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=jeux;User ID=microssvc;Password=microssvc";
            sql = "select PHOTO from jeux.dbo.PHOTO where NOCLI=@plyrId;";
            connection = new SqlConnection(connetionString);


            connection.Open();
            command = new SqlCommand(sql, connection);
            command.Parameters.Add("@plyrId", System.Data.SqlDbType.Int);
            command.Parameters["@plyrId"].Value = PlayerID;
            /*rdr = command.ExecuteReader();
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    image = (byte[])rdr[0];
                }
            }*/
            PictureBox pictureBox = new PictureBox();
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            da.Fill(ds);
            if (ds.Tables[0].Rows.Count > 0)
            {
                MemoryStream mst = new MemoryStream((byte[])ds.Tables[0].Rows[0]["PHOTO"]);
                pictureBox.Image = new Bitmap(mst);

            }

            command.Dispose();
            connection.Close();

            /*MemoryStream ms = new MemoryStream(image);

            PictureBox pictureBox = new PictureBox();
           //pictureBox.Image = System.Drawing.Bitmap.FromStream(ms);
           pictureBox.Image = System.Drawing.Image.FromStream(ms);
           pictureBox.Location = new Point(23, 19);
           pictureBox.Margin = new Padding(4);
           pictureBox.Name = "pictureBox";
           pictureBox.Size = new Size(194, 202);
           pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
           pictureBox.TabIndex = 0;
           pictureBox.TabStop = false;
           pictureBox.Show();*/

        }
        /****************************************************************************/
        /********************** Utilisation des invitation *************************/
        /***************************************************************************/
        public void DebitInvitations()
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            string sql = null;

            try
            {
                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=CasinoCore;User ID=microssvc;Password=microssvc";
                sql = "Execute CasinoCore.dbo.Micros_DebitInvitation @plyrId, @invutl1, @rvc, @tltpcntg, @niv, @checknum;";
                connection = new SqlConnection(connetionString);


                connection.Open();
                command = new SqlCommand(sql, connection);
                command.Parameters.Add("@plyrId", System.Data.SqlDbType.Int);
                command.Parameters.Add("@invutl1", System.Data.SqlDbType.Int);
                command.Parameters.Add("@rvc", System.Data.SqlDbType.Int);
                command.Parameters.Add("@tltpcntg", System.Data.SqlDbType.Float);
                command.Parameters.Add("@niv", System.Data.SqlDbType.SmallInt);
                command.Parameters.Add("@checknum", System.Data.SqlDbType.VarChar);
                command.Parameters["@plyrId"].Value = PlayerID;
                command.Parameters["@invutl1"].Value = UtilInvitation;
                command.Parameters["@rvc"].Value = OpsContext.RvcNumber;
                command.Parameters["@tltpcntg"].Value = montant_Invit;
                command.Parameters["@niv"].Value = Niveau;
                command.Parameters["@checknum"].Value = OpsContext.CheckNumber.ToString(); ;
                rdr = command.ExecuteReader();
                while (rdr.Read())
                {
                    //result = "Le solde des points: " + rdr[0].ToString();
                    NouveauSolde = Convert.ToInt64(rdr[0]);
                }
                command.Dispose();
                connection.Close();

                OpsContext.ShowMessage(string.Format("Le Client a consommé {0} invitations", UtilInvitation));
            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }

        }

        /*******************************************************************/
        // Ajout des points sur le membre				//
        /******************************************************************/
        [ExtensibilityMethod]
        public void CreditPoints()
        {
            string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            SqlDataReader rdr;
            string sql = null;
            string facture = null;
            //result =null;
            facture = OpsContext.RvcName + "Fact :" + OpsContext.CheckNumber.ToString(); //Formatage du texte pour la reference facture.

            try
            {
                connetionString = @"Data Source=172.26.60.94\OCM,1814;Database=CasinoCore;User ID=microssvc;Password=microssvc";
                sql = "Execute CasinoCore.dbo.Micros_CreditPoints @noOCM , @nbrpts , @fct;"; // la procedure stockée à executer
                connection = new SqlConnection(connetionString); // connexion à la base de données

                connection.Open();
                command = new SqlCommand(sql, connection);
                command.Parameters.Add("@noOCM", System.Data.SqlDbType.Int);
                command.Parameters.Add("@nbrpts", System.Data.SqlDbType.Int);
                command.Parameters.Add("@fct", System.Data.SqlDbType.VarChar);
                command.Parameters["@noOCM"].Value = PlayerID;
                command.Parameters["@nbrpts"].Value = PointsAcquis;
                command.Parameters["@fct"].Value = facture;
                rdr = command.ExecuteReader();
                while (rdr.Read())
                {

                    NouveauSolde = Convert.ToInt64(rdr[0]);

                }
                command.Dispose();
                connection.Close();

                //result = "Solde de points: " + NouveauSolde + "\nNombre d'invitation1: " + (Invitations - UtilInvitation);
                //OpsContext.ShowMessage(string.Format(result));
            }
            catch (Exception ex)
            {
                OpsContext.ShowError(string.Format(ex.Message));
            }


        }

        /**************************** Test Communication with Tpe pour O'tacos *******************/

        [ExtensibilityMethod]
        public void CommunicateWithTpe()
        {
            string name;
            string message;
            StringComparer stringComparer = StringComparer.OrdinalIgnoreCase;
            //Thread readThread = new Thread(Read);

            // Create a new SerialPort object with default settings.
            _serialPort = new SerialPort();

            // Allow the user to set the appropriate properties.
            _serialPort.PortName = "COM2";
            // _serialPort.PortName = SetPortName();
            _serialPort.BaudRate = 9600;
            _serialPort.Parity = System.IO.Ports.Parity.None;
            _serialPort.DataBits = 8;
            _serialPort.StopBits = System.IO.Ports.StopBits.One;
            _serialPort.Handshake = System.IO.Ports.Handshake.XOnXOff;
            _serialPort.ReadTimeout = 1000000;

            /*_serialPort.PortName = SetPortName(_serialPort.PortName);
            _serialPort.BaudRate = SetPortBaudRate(_serialPort.BaudRate);
            _serialPort.Parity = SetPortParity(_serialPort.Parity);
            _serialPort.DataBits = SetPortDataBits(_serialPort.DataBits);
            _serialPort.StopBits = SetPortStopBits(_serialPort.StopBits);
            _serialPort.Handshake = SetPortHandshake(_serialPort.Handshake);*/

            // Set the read/write timeouts
            _serialPort.ReadTimeout = 500;
            _serialPort.WriteTimeout = 500;

            _serialPort.Open();
            _continue = true;
            //readThread.Start();

            //Console.Write("Name: ");
            name = "Intelliway";

            OpsContext.ShowMessage("Type QUIT to exit");

            while (_continue)
            {
                message = OpsContext.RequestAlphaEntry("entrez le message", "Serial Port");

                if (stringComparer.Equals("quit", message))
                {
                    _continue = false;
                }
                else
                {
                    _serialPort.WriteLine(
                        String.Format("<{0}>: {1}", name, message));
                }
            }

            //readThread.Join();
            _serialPort.Close();
        }

        public void Read()
        {
            while (_continue)
            {
                try
                {
                    string message = _serialPort.ReadLine();
                    Console.WriteLine(message);
                }
                catch (TimeoutException) { }
            }
        }

        public string SetPortName()
        {
            string portName;

            OpsContext.ShowMessage("Available Ports:");
            foreach (string s in SerialPort.GetPortNames())
            {
                OpsContext.ShowMessage(string.Format(s));
            }
            portName = OpsContext.RequestAlphaEntry("Enter COM port value", "Serial Port");
            return portName;
        }


        /*****************************************************************************************/

    }


    /// <summary>
    ///  Implements interface used by Simphony POS Client to create an instance of the extension application
    /// </summary>
    public class ApplicationFactory : IExtensibilityAssemblyFactory
    {
        public ExtensibilityAssemblyBase Create(IExecutionContext context)
        {
            return new Application(context);
        }

        public void Destroy(ExtensibilityAssemblyBase app)
        {
            app.Destroy();
        }
    }
}
