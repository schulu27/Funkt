using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.Runtime.Remoting;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClassTesterFinal
{


    public class MailSender
    {

        private string _textinhalt;
        private string _attachmentPath;
        private string _recepient;
        private string _subject;
        private string _ccRecepient;


        public string Subject
        {
            get { return _subject; }
            set { _subject = value; }
        }


        public string AttachmentPath
        {
            get { return _attachmentPath; }
            set { _attachmentPath = value; }
        }


        public string Textinhalt
        {
            get { return _textinhalt; }
            set { _textinhalt = value; }
        }

        public string Recepient
        {
            get { return _recepient; }
            set { _recepient = value; }
        }

        public string CCRecepient
        {
            get { return _ccRecepient; }
            set { _ccRecepient = value; }
        }


        public void SendMail()
        {

            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook._MailItem oMailItem = (Outlook._MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                if (outlookApp.ActiveInspector() == null)
                {


                    //Add Recipient
                    Outlook.Recipient oRecip = oMailItem.Recipients.Add(Recepient);
                    oRecip.Resolve();

                    //Add CC
                    oMailItem.CC = CCRecepient;


                    // Add Subject
                    oMailItem.Subject = Subject;

                    //Add Attachment
                    //Outlook.Attachment attachment = oMailItem.Attachments.Add(AttachmentPath);

                    // Body
                    string textinhalt = Textinhalt;
                    string absender = string.Empty;
                    string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                    username = username.ToLower();

                    if (username.Contains("schlui"))
                    {
                        absender = "\n\n\r\nFreundliche Grüße aus Garmisch-Partenkirchen" +
                                   "\r\n\r\nSchubert Luis" +
                                   "\r\n\r\nWerkstudent Entwicklung Mechatronik | Prototypenbau " +
                                   "\n\rTelefon +49 8821 920-272" +
                                   "\nFax         +49 8821 920-2167" +
                                   "\neMail " + "l.schubert@langmatz.de" +
                                   "\nWerk I Am Gschwend 10, 82467 Garmisch-Partenkirchen";
                    }
                    else if (username.Contains("andmen"))
                    {

                        absender = "\n\n\r\nFreundliche Grüße aus Garmisch-Partenkirchen" +
                                   "\r\n\r\nAndreas Menhart" +
                                   "\r\n\r\nEntwicklung Mechatronik | Prototypenbau" +
                                   "\n\rTelefon +49 8821 920-267" +
                                   "\nFax         +49 8821 920-2167" +
                                   "\neMail " + "a.menhart@langmatz.de" +
                                   "\nWerk I Am Gschwend 10, 82467 Garmisch-Partenkirchen";
                    }
                    else if (username.Contains("hebkar"))
                    {

                        absender = "\n\n\r\nFreundliche Grüße aus Garmisch-Partenkirchen" +
                                   "\r\n\r\nKarl-Heinz Hebermehl" +
                                   "\r\n\r\nVertrieb Technischer Innendienst" +
                                   "\n\rTelefon +49 8821 920-126" +
                                   "\nFax         +49 8821 920-2417" +
                                   "\neMail " + "k.hebermehl@langmatz.de" +
                                   "\nWerk I Am Gschwend 10, 82467 Garmisch-Partenkirchen";
                    }
                    else if (username.Contains("friste"))
                    {

                        absender = "\n\n\r\nFreundliche Grüße aus Garmisch-Partenkirchen" +
                                   "\r\n\r\nStefan Fritsch" +
                                   "\r\n\r\nVertrieb Technischer Innendienst" +
                                   "\n\rTelefon +49 8821 920-175" +
                                   "\nFax         +49 8821 920-2417" +
                                   "\neMail " + "s.fritsch@langmatz.de" +
                                   "\nWerk I Am Gschwend 10, 82467 Garmisch-Partenkirchen";
                    }
                    else if (username.Contains("siekai"))
                    {

                        absender = "\n\n\r\nFreundliche Grüße aus Garmisch-Partenkirchen" +
                                  "\r\n\r\nKai Siemer" +
                                  "\r\n\r\nEntwicklung Mechatronik | Prototypenbau Funktionsstellenleitung" +
                                  "\n\rTelefon +49 8821 920-167" +
                                  "\nFax         +49 8821 920-2167" +
                                  "\neMail " + "k.siemer@langmatz.de" +
                                  "\nWerk I Am Gschwend 10, 82467 Garmisch-Partenkirchen";


                    }

                    else
                    {
                        absender = "\n\n\r\nFreundliche Grüße aus Garmisch-Partenkirchen" +
                                  "\r\n\r\nIhr Name" +
                                  "\r\n\r\nAbteilung | Funktion" +
                                  "\n\rTelefon +49 8821 920-xxx" +
                                  "\nFax         +49 8821 920-xxxx" +
                                  "\neMail " + "Ihre.Email@langmatz.de" +
                                  "\nWerk I Am Gschwend 10, 82467 Garmisch-Partenkirchen";



                        MessageBox.Show("Bitte ergänzen Sie Angaben.", "Email Eingeben");

                    }
                    textinhalt = textinhalt + absender;

                    oMailItem.Body = textinhalt;

                    // Get the Inspector object
                    //Outlook.Inspector oInspector = oMailItem.Inspector;

                    // Display the mailbox
                    //oInspector.Display(true);

                    // Display the mailbox
                    oMailItem.Display(true);

                }
                else
                {
                    MessageBox.Show("Entwurf bereits geöffnet, bitte schließen um fortzufahren.");
                }
            }
            catch (Exception objEx)
            {
                // Handle Outlook error
                return;
            }
        }
    }
}