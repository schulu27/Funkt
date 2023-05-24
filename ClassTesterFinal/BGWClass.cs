//__________________________________________________________________Beschreibung:________________________________________________________________________________________________
//
// Mithilfe dieser Klasse koennen Sie einen Backgroundworker aufrufen der Ihnen eine Progressbar in einer neuen Form oeffnet um wartezeiten zu ungefaehr zu visualisieren
//
//      //Erstellen der Instanzen der Klasse
//      BGWorking bgWorker = new BGWorking();
//      ProgressForm progressForm = new ProgressForm();
//      bgWorker.StartProgressBar(progressForm.ProgressBarControl);
//---------------------------------------------------------------------Start Testprogramm:----------------------------------------------------------------------------------------
//      Thread sendMailThread = new Thread(() =>
//---------------------------------------------------------------------Start Testprogramm:----------------------------------------------------------------------------------------
//      {
//      MailSender mailSender = new MailSender();
//      mailSender.Recepient = "herr@muster.de";
//      mailSender.CCRecepient = "frau@muster.de";
//      mailSender.Subject = "Testzeile";
//      mailSender.AttachmentPath = "Filepath";
//      mailSender.Textinhalt = "Sehr geehrte Damen und Herren, \r\n\r\nim Anhang erhalten Sie zu der Seriennummer:[" + txtBoxEin.Text + "] die ausgewähltem Freischaltcodes:\n";
//      mailSender.SendMail();
//---------------------------------------------------------------------Ende Testprogramm:-----------------------------------------------------------------------------------------
//      bgWorker.StopProgressBar();
//---------------------------------------------------------------------Ende Testprogramm:-----------------------------------------------------------------------------------------
//      });
//      sendMailThread.Start();
//
//__________________________________________________________________Beschreibung:__________________________________________________________________________________________________

using System;
using System.ComponentModel;

public class BGWorker
{
    private BackgroundWorker worker;

    public void MyBackgroundWorker()
    {
        worker = new BackgroundWorker();
        worker.DoWork += Worker_DoWork;
        worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
    }

    public void StartWork()
    {
        if (!worker.IsBusy)
            worker.RunWorkerAsync();
    }

    private void Worker_DoWork(object sender, DoWorkEventArgs e)
    {
        // Hier kannst du deine Hintergrundarbeit ausführen
        // Beachte, dass dieser Code im Hintergrund ausgeführt wird
        // und nicht auf dem Haupt-Thread läuft
        Console.WriteLine("Hintergrundarbeit läuft...");
        System.Threading.Thread.Sleep(2000);
    }

    private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        // Dieser Code wird ausgeführt, wenn die Hintergrundarbeit abgeschlossen ist
        Console.WriteLine("Hintergrundarbeit abgeschlossen.");
    }
}
