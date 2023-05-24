using ClassTester;
using System.ComponentModel;
using static ClassTester.UbootClass;

namespace ClassTesterFinal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CalcCodeClass calcCode = new CalcCodeClass();
            calcCode.FreischaltCodeBerechnen(txtBoxEin.Text);
            label1.Text = Convert.ToString(calcCode.CodeTelicm);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IsOpen isOpen = new IsOpen();
            UbootClass uboot = new UbootClass();
            BackgroundWorker backgroundWorker = new BackgroundWorker();

            uboot.StartAnimation();

            Action task = () =>
            {
                SendEmail();
            };

            backgroundWorker.DoWork += (sender, e) =>
            {
                task();
            };

            backgroundWorker.RunWorkerAsync();

            isOpen.StartCheckingDraft();

            // Hintergrundthread für das Stoppen der Animation starten
            Thread stopAnimationThread = new Thread(() =>
            {
                Thread.Sleep(2700); // Optional: Wartezeit vor dem Stoppen der Animation

                // Zugriff auf das Steuerelement im Hauptthread ermöglichen
                uboot.Invoke((MethodInvoker)(() =>
                {
                    
                    uboot.StopAnimation();
                    
                }));
            });
            stopAnimationThread.Start();
        }


        private void SendEmail()
            {
                MailSender mailSender = new MailSender();
                mailSender.Recepient = "herr@muster.de";
                mailSender.CCRecepient = "frau@muster.de";
                mailSender.Subject = "Testzeile";
                mailSender.AttachmentPath = "Filepath";
                mailSender.Textinhalt = "Sehr geehrte Damen und Herren, \r\n\r\nim Anhang erhalten Sie zu der Seriennummer:[" + txtBoxEin.Text + "] die ausgewähltem Freischaltcodes:\n";
                mailSender.SendMail();
            }

        }
    }


//MailSender mailSender = new MailSender();
//mailSender.Recepient = "herr@muster.de";
//mailSender.CCRecepient = "frau@muster.de";
//mailSender.Subject = "Testzeile";
//mailSender.AttachmentPath = "Filepath";
//mailSender.Textinhalt = "Sehr geehrte Damen und Herren, \r\n\r\nim Anhang erhalten Sie zu der Seriennummer:[" + txtBoxEin.Text + "] die ausgewähltem Freischaltcodes:\n";
//mailSender.SendMail();




//while (!draftChecker.IsDraftOpen())
//{
//    // Verzögerung vor der nächsten Prüfung
//    System.Threading.Thread.Sleep(10);
//}

//uboot.StopAnimation();