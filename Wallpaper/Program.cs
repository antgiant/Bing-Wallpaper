using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Wallpaper
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string file = "";
            uint monitor = uint.MaxValue;

            //Only evaluate 1 or two arguments
            if (args.Length >= 1 && args.Length <= 2)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    //If argument is a number set wallpaper on only that monitor (Assumes less than 10,000 monitors.)
                    if (args[i].Length < 5)
                    {
                        try
                        {
                            //Subtract one so that passed in monitor number equals what windows displays
                            monitor = Convert.ToUInt32(args[i]) - 1;

                            //If specified monitor number is greater than the current number of monitors do nothing and exit.
                            var wallpaper = (IDesktopWallpaper)(new DesktopWallpaperClass());
                            if (wallpaper.GetMonitorDevicePathCount() < monitor)
                            {
                                Environment.Exit(0);
                            }
                        }
                        catch
                        {
                            //If monitor number is invalid do nothing and exit.
                            Environment.Exit(0);
                        }
                    }
                    //If argument is a URL download and set wallpaper to that. (Only use trusted URLs as this will download anything.)
                    else if (args[i].Substring(0, 4).ToLower() == "http")
                    {
                        string path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\Wallpapers";

                        if (!Directory.Exists(path))
                            Directory.CreateDirectory(path);

                        var bd =
                            new Downloader(path)
                            {
                                UseHttps = true,
                                Size =
                                    new Rect
                                    {
                                        Right = Screen.PrimaryScreen.Bounds.Width,
                                        Bottom = Screen.PrimaryScreen.Bounds.Height
                                    }
                            };

                        file = bd.DownloadSync(args[i]);

                        //If specified URL cannot be downloaded do nothing and exit.
                        if (file == string.Empty)
                        {
                            Environment.Exit(0);
                        }
                    }
                    //Assume argument was a wallpaper image
                    else
                    {
                        //If specified file does not exist and path is not a folder do nothing and exit.
                        if (File.Exists(args[i]))
                        {
                            file = args[i];
                        }
                        else if (Directory.Exists(args[i]))
                        {
                            var files = Directory.EnumerateFiles(args[i], "*.*")
                                                .Where(s => s.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".gif", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".tif", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)
                                                            || s.EndsWith(".tiff", StringComparison.OrdinalIgnoreCase)
                                                       );
                            if (files.Count() > 0)
                            {
                                Random rand = new Random();
                                file = files.ElementAt(rand.Next(0, files.Count()));
                            }
                            else
                            {
                                Environment.Exit(0);
                            }
                        }
                        else
                        {
                            Environment.Exit(0);
                        }
                    }
                }
            }

            //Default to Bing Wallpaper
            if (file == "")
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\Wallpapers";

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                var bd =
                    new Downloader(path)
                    {
                        UseHttps = true,
                        Size =
                            new Rect
                            {
                                Right = Screen.PrimaryScreen.Bounds.Width,
                                Bottom = Screen.PrimaryScreen.Bounds.Height
                            }
                    };

                file = bd.DownloadSync();
            }

            if (!string.IsNullOrEmpty(file))
            {
                if (monitor == uint.MaxValue)
                {
                    SetAsWallPaper(file);
                }
                else
                {
                    SetAsWallPaper(file, monitor);
                }
            }
        }

        private static void SetAsWallPaper(string file)
        {
            try
            {
                var wallpaper = (IDesktopWallpaper)(new DesktopWallpaperClass());

                for (uint i = 0; i < wallpaper.GetMonitorDevicePathCount(); i++)
                {
                    wallpaper.SetWallpaper(wallpaper.GetMonitorDevicePathAt(i), file);
                    wallpaper.SetPosition(DesktopWallpaperPosition.Fill);
                }
            }
            catch
            {
                Environment.Exit(0);
            }
        }
        private static void SetAsWallPaper(string file, uint monitor)
        {
            try
            {
                var wallpaper = (IDesktopWallpaper)(new DesktopWallpaperClass());

                wallpaper.SetWallpaper(wallpaper.GetMonitorDevicePathAt(monitor), file);
                wallpaper.SetPosition(DesktopWallpaperPosition.Fill);
            }
            catch
            {
                Environment.Exit(0);
            }
        }
    }
}