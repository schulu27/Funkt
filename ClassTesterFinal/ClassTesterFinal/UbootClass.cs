using ClassTesterFinal;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClassTester
{
       // Klasse 1
   

    public class UbootClass : Form
    {
        private PictureBox iconPictureBox;
        private BackgroundWorker animationWorker;
        private int movementSpeed = 3;
        private int direction = 1; // 1 für links nach rechts, -1 für rechts nach links
        private int startX = 0;
        private int originalY; 
        private BackgroundWorker _customWorker;
        private System.Timers.Timer animationTimer;


        //-----------------------------------uBoot-------------------------------------------------------------
        public UbootClass()
        {
            InitializeComponent();
            InitializeIcon();
            InitializeAnimationWorker();


        }

        //......................................Spielerein.....................................................
            private void InitializeComponent()
            {
                BackColor = Color.LightBlue;
                FormBorderStyle = FormBorderStyle.None;
                Width = 427;
                Height = 50;
                Top = (Screen.PrimaryScreen.Bounds.Height - Height) / 2;
                Left = (Screen.PrimaryScreen.Bounds.Width - Width) / 2;
            }

            private void InitializeIcon()
            {
                iconPictureBox = new PictureBox();
                iconPictureBox.Image = ClassTesterFinal.Properties.Resources.YelSub.ToBitmap();
                iconPictureBox.Size = new Size(50, 50);
                iconPictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
                iconPictureBox.Location = new Point(startX, (ClientSize.Height - iconPictureBox.Height) / 2);
                originalY = iconPictureBox.Location.Y;
                Controls.Add(iconPictureBox);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                GraphicsPath path = new GraphicsPath();
                int cornerRadius = 5;
                path.AddArc(0, 0, cornerRadius * 2, cornerRadius * 2, 180, 90);
                path.AddArc(Width - cornerRadius * 2, 0, cornerRadius * 2, cornerRadius * 2, 270, 90);
                path.AddArc(Width - cornerRadius * 2, Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90);
                path.AddArc(0, Height - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90);
                path.CloseFigure();
                Region = new Region(path);
            }

        //......................................Spielerein.....................................................

        private void MoveIcon()
        {
            if (IsHandleCreated)
            {
                if (InvokeRequired)
                {
                    Invoke(new MethodInvoker(MoveIcon));
                    return;
                }

                int newX = iconPictureBox.Location.X + (movementSpeed * direction);
                int newY = iconPictureBox.Location.Y;

                if (newX > ClientSize.Width - iconPictureBox.Width || newX < 0)
                {
                    direction *= -1;
                    Image originalImage = iconPictureBox.Image;
                    originalImage.RotateFlip(RotateFlipType.RotateNoneFlipY);
                    originalImage.RotateFlip(RotateFlipType.Rotate180FlipNone);
                    iconPictureBox.Image = originalImage;
                    newX = iconPictureBox.Location.X;
                    newY = originalY;
                }

                iconPictureBox.Location = new Point(newX, newY);
            }
        }


        private void AnimationWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            while (!animationWorker.CancellationPending)
            {
                MoveIcon();
                Thread.Sleep(10);
            }
        }

        public void StartAnimation()
        {
            Show();
            InitializeAnimationWorker();
            animationWorker.RunWorkerAsync();
        }

        public void StopAnimation()
        {
            if (animationWorker.IsBusy)
            {
                animationWorker.CancelAsync();
            }
            Hide();
        }

        private void InitializeAnimationWorker()
        {
            animationWorker = new BackgroundWorker();
            animationWorker.WorkerSupportsCancellation = true;
            animationWorker.DoWork += AnimationWorker_DoWork;
        }



        //-----------------------------------uBoot-------------------------------------------------------------


        //-----------------------------------customBgWorker----------------------------------------------------

        public class BackgroundTaskRunner
        {
            private BackgroundWorker _worker;

            public BackgroundTaskRunner()
            {
                _worker = new BackgroundWorker();
                _worker.WorkerSupportsCancellation = true;
                _worker.DoWork += Worker_DoWork;
            }

            public void RunTask(Action task)
            {
                _worker.RunWorkerAsync(task);
            }

            private void Worker_DoWork(object sender, DoWorkEventArgs e)
            {
                Action task = e.Argument as Action;
                if (task != null)
                {
                    task.Invoke();
                }
            }
        }

        //-----------------------------------customBgWorker----------------------------------------------------
    }


    //-------------------------------------------------isOpenClass---------------------------------------------

    public class IsOpen
    {
        private System.Timers.Timer timer;
        private BackgroundWorker backgroundWorker;

        public IsOpen()
        {
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
        }

        public void StartCheckingDraft()
        {
            StartTimer();
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            if (!backgroundWorker.IsBusy)
            {
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (IsDraftOpen())
            {
                StopTimer();
                // Weitere Aktionen durchführen, wenn ein Entwurf geöffnet ist
            }
        }

        public bool IsDraftOpen()
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.Inspector inspector = outlookApp.ActiveInspector();

                if (inspector == null)
                {
                    // Es ist kein Entwurf geöffnet
                    return false;
                }

                Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
                if (mailItem != null && mailItem.EntryID == null)
                {
                    // Es ist ein Entwurf geöffnet
                    return true;
                }

                // Es ist kein Entwurf geöffnet
                return false;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Outlook ist nicht geöffnet oder nicht installiert
                return false;
            }
        }

        private void StartTimer()
        {
            if (timer == null)
            {
                timer = new System.Timers.Timer(10);
                timer.Elapsed += TimerElapsed;
            }
            timer.Start();
        }

        private void StopTimer()
        {
            if (timer != null)
            {
                timer.Stop();
            }
        }
    }


    //-------------------------------------------------isOpenClass---------------------------------------------

}

































//Klasse 2
//public class BGWorking
//{
//    private BackgroundWorker _customWorker;

//    public BGWorking()
//    {
//        _customWorker = new BackgroundWorker();
//        _customWorker.WorkerSupportsCancellation = true;
//    }

//    public void StartCustomWorker(DoWorkEventHandler customWorkerDoWork, RunWorkerCompletedEventHandler customWorkerRunWorkerCompleted)
//    {
//        _customWorker.DoWork += customWorkerDoWork;
//        _customWorker.RunWorkerCompleted += customWorkerRunWorkerCompleted;
//        _customWorker.RunWorkerAsync();
//    }

//    public void StopCustomWorker()
//    {
//        if (_customWorker.IsBusy)
//        {
//            _customWorker.CancelAsync();
//        }
//    }
//}