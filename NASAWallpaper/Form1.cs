using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Reactive.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace NASAWallpaper
{
    public partial class Form1 : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;

        private readonly WallpaperSetter wallpaper = new WallpaperSetter();
        private readonly IScheduler scheduler = StdSchedulerFactory.GetDefaultScheduler();

        public Form1()
        {
            // Create a simple tray menu with only one item.
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);
            trayMenu.MenuItems.Add("Update", OnUpdate);

            // Create a tray icon. In this example we use a
            // standard system icon for simplicity, but you
            // can of course use your own custom icon too.
            trayIcon = new NotifyIcon();
            trayIcon.Text = "NASA WP";
            trayIcon.Icon = new Icon("avatar-nasa40x40.ico", 40, 40);
            

            // Add menu to tray icon and show it.
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
        }

        private void OnUpdate(object sender, EventArgs e)
        {
            wallpaper.Execute(null);
        }

        protected async override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.

            
            wallpaper.Execute(null);

            scheduler.Start();

            IJobDetail job = JobBuilder.Create<WallpaperSetter>().Build();

            ITrigger trigger = TriggerBuilder.Create()
                .WithDailyTimeIntervalSchedule
                  (s =>
                     s.WithIntervalInHours(5)
                    .OnEveryDay()
                  )
                .Build();

            scheduler.ScheduleJob(job, trigger);

            base.OnLoad(e);
        }

        private void OnExit(object sender, EventArgs e)
        {
            scheduler.Shutdown();
            Application.Exit();
        }        
    }

    class WallpaperSetter : IJob
    {
        private const string NasaDeilyImage = "http://apod.nasa.gov/apod/";

        public void Execute(IJobExecutionContext context)
        {
            var imageUrlManager = new ImagerUrlManager();
            var imageUrl = imageUrlManager.Image(NasaDeilyImage);

            var fileExtension = imageUrl.Substring(imageUrl.LastIndexOf('.'));
            var wallpaperImageName = "wallpaper" + fileExtension;

            var imageSaver = new ImageSaver();
            imageSaver.SaveImage(NasaDeilyImage + "/" + imageUrl, wallpaperImageName);

            WINWallpaperSetter.SetWallpaper(wallpaperImageName);
        }
    }

    class WINWallpaperSetter
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern Int32 SystemParametersInfo(UInt32 uiAction, UInt32 uiParam, String pvParam, UInt32 fWinIni);
        private static UInt32 SPI_SETDESKWALLPAPER = 20;
        private static UInt32 SPIF_UPDATEINIFILE = 0x01;
        public static void SetWallpaper(string file)
        {
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, Environment.CurrentDirectory + "\\" + file, SPIF_UPDATEINIFILE);
        }
    }

    class ImageSaver
    {
        public void SaveImage(string imageUrl, string imageName)
        {
            using (var webClient = new WebClient())
            {
                webClient.DownloadFile(imageUrl, imageName);
            }
        }
    }

    class ImagerUrlManager
    {
        public string Image(string url)
        {
            string htmlPage;
            using (var webClient = new WebClient())
            {
                try
                {
                    htmlPage = webClient.DownloadString(url);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            return ImageUrl(htmlPage);
        }

        private string ImageUrl(string html)
        {
            const string tag = "<IMG SRC=\"";
            int firstIndex;
            if ((firstIndex = html.IndexOf(tag)) == -1)
            {
                throw new Exception("Can't find tag! Datetime = " + DateTime.Now);
            }
            firstIndex += tag.Length;
            int endIndex = html.IndexOf("\"", firstIndex);
            return html.Substring(firstIndex, endIndex - firstIndex);
        }
    }
}
