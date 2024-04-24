using DSZahirDesktop;
using Serilog;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Management;
using System.Runtime.InteropServices;

using WindowsInput;

namespace BeeAcct29_Automation
{
    internal class Program
    {
        static string dtID;
        static string dtName;
        static string dbpath;
        static string loginId;
        static string password;
        static string dbname;
        static string runasadmin;
        static string waitappload;
        static string erpappnamepath;
        static string issandbox;
        static string enableconsolelog;
        static string appfolder, uploadfolder, sharingfolder;
        static string datapicturefolder;
        static string screenshotlogfolder;
        static clsSearch MySearch = null;
        static InputSimulator iSim = new InputSimulator();
        enum leftClick
        {
            sngl,
            dbl
        }


        static string logfilename = "";
        static int pid = 0;

        [DllImport("user32.dll")]
        private static extern bool BlockInput(bool fBlockIt);

        private static void ThisClassInit()
        {
            dtID = ConfigurationManager.AppSettings["dtID"];
            dtName = ConfigurationManager.AppSettings["dtName"];
            dbpath = ConfigurationManager.AppSettings["dbpath"];
            loginId = ConfigurationManager.AppSettings["loginId"];
            password = ConfigurationManager.AppSettings["password"];
            dbname = ConfigurationManager.AppSettings["dbname"];
            runasadmin = ConfigurationManager.AppSettings["runasadmin"].ToUpper();
            waitappload = ConfigurationManager.AppSettings["waitappload"];
            erpappnamepath = ConfigurationManager.AppSettings["erpappnamepath"];
            issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
            enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
            appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
            uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
            sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
            datapicturefolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\" + ConfigurationManager.AppSettings["datapicturefolder"];
            screenshotlogfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotlogfolder"];
        }


        private static bool findimage(string imagename, out Point pnt)
        {
            Point p, absp;
            p = new Point(0, 0);
            absp = new Point(0, 0);

            try
            {
                Bitmap ImgToFind = new Bitmap(datapicturefolder + $@"\{imagename}.png");

                if (MySearch.CompareImages(ImgToFind, datapicturefolder, out p, out absp))
                {
                    Log.Information($"Search and found image name => {imagename}.png");
                    pnt = absp;
                    return true;
                }
                else
                {
                    Log.Information($"Cannot find image named => {imagename}.png !!!");
                    pnt = absp;
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.Information($"{ex.Message}=> Quitting.. End of FindImage automation function !!!");
                pnt = absp;
                return false;
            }

        }

        private static void SimulateMouseClick(Point point, leftClick click)
        {
            iSim.Mouse.MoveMouseTo(point.X, point.Y);
            Int32 x = point.X / 100;
            Int32 y = point.Y / 100;
            if (click == leftClick.sngl)
            {
                iSim.Mouse.LeftButtonClick();
                Log.Information($"Single click interaction with above named image at X={x} & Y={y} point.");
            }
            else
            {
                iSim.Mouse.LeftButtonDoubleClick();
                Log.Information($"Double click interaction with above named image at X={x} & Y={y} point.");
            }

        }

        static void Main(string[] args)
        {
            Int32 errStep = 0;
            MyDirectoryManipulator myFileUtil = new MyDirectoryManipulator();
            try
            {
                ThisClassInit();
                int resX = 0;
                int resY = 0;
                ManagementObjectSearcher mydisplayResolution = new ManagementObjectSearcher("SELECT CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController");
                foreach (ManagementObject record in mydisplayResolution.Get())
                {
                    resX = Convert.ToInt32(record["CurrentHorizontalResolution"]);
                    resY = Convert.ToInt32(record["CurrentVerticalResolution"]);
                }
                MySearch = new clsSearch(resX, resY);

                //* Call this method to disable keyboard input
                int maxWidth = Console.LargestWindowWidth;
                Console.Title = "Bee Accounting Desktop Version 2.9 - Automasi - By PT FAIRBANC TECHNOLOGIEST INDONESIA";
                Console.WindowLeft = 0;
                Console.WindowTop = 0;
                Console.SetWindowPosition(0, 0);
                //Console.WindowWidth = maxWidth;
                Console.BackgroundColor = ConsoleColor.DarkGray;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"");
                Console.WriteLine($"******************************************************************");
                Console.WriteLine($"                    Automasi akan dimulai !                       ");
                Console.WriteLine($"             Keyboard dan Mouse akan di matikan...                ");
                Console.WriteLine($"     Komputer akan menjalankan oleh applikasi robot automasi...   ");
                Console.WriteLine($" Aktifitas penggunakan komputer akan ter-BLOKIR sekitar 10 menit. ");
                Console.WriteLine($"******************************************************************");
                Console.WriteLine($"      Resolusi layar adalah lebar: {resX.ToString("0000")}, dan tinggi: {resY.ToString("0000")}         ");

#if DEBUG
                BlockInput(false);
#else
                BlockInput(true);
#endif

                if (!Directory.Exists(appfolder))
                {
                    myFileUtil.CreateDirectory(appfolder);
                    myFileUtil.CreateDirectory(uploadfolder);
                    myFileUtil.CreateDirectory(sharingfolder);
                    myFileUtil.CreateDirectory(datapicturefolder);
                }
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.BackgroundColor = ConsoleColor.Black;
                var temp0 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Csv);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp0}"));
                do
                {
                } while (!Task.CompletedTask.IsCompleted);
                var temp1 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Excel);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp1}"));
                do
                {
                } while (!Task.CompletedTask.IsCompleted);
                var temp2 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Log);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp2}"));
                var temp3 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Zip);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp3}"));
                var config = new LoggerConfiguration();
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                Log.Logger = config.CreateLogger();

                Log.Information("Bee Accounting Desktop Automation - *** Started *** ");


                if (!OpenApp())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("Application automation failed when running app (OpenApp) !!!");
                    return;
                }

                if (!OpenDB(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"Application automation failed when running app (OpenDB) on step: {errStep.ToString()} !!!");
                    return;
                }

                //if (!OpenReport(out errStep, "sales"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> Sales) on step: {errStep} !!!");
                    return;
                }
                //if (!ClosingReport(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (ClosingReport) on step: {errStep} !!!");
                    return;
                }

                //if (!OpenReport(out errStep, "ar"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> AR) on step: {errStep.ToString()}  !!!");
                    return;
                }

                //if (!ClosingReport(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (ClosingReport) on step: {errStep.ToString()}  !!!");
                    return;
                }
                //if (!OpenReport(out errStep, "outlet"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (OpenReport -> Outlet) on step: {errStep.ToString()} !!!");
                    return;
                }

                //if (!ClosingReport(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (ClosingWorkspace) on step: {errStep} !!!");
                    return;
                }
                //if (!CloseApp(out errStep))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information($"application automation failed when running app (CloseApp) on step: {errStep} !!!");
                    return;
                }
                //if (ZipandSend())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ZipandSend) !!!");
                    return;
                }
            }
            catch (Exception ex)
            {
                Log.Information($"Zahir v5 automation error => {ex.ToString()}");
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")} INF] Zahir Desktop automation error => {ex.ToString()}"));
            }
            finally
            {
                Console.Beep();
                Task.Delay(500);
                Console.Beep();
                Task.Delay(500);
                //* Call this method to enable keyboard input
                BlockInput(false);

                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Zahir Desktop Automation - ***   END   ***"));

                Log.CloseAndFlush();
            }
        }

        static bool OpenApp()
        {
            try
            {
                Log.Information($"Wait app loading for {Convert.ToInt32(waitappload) / 1000} sec.");
                if (runasadmin == "Y")
                {
                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = Environment.GetEnvironmentVariable("WINDIR") + "\\explorer.exe";
                    //psi.Verb = "runasuser";
                    //psi.LoadUserProfile = true;
                    psi.Arguments = @$"{erpappnamepath}";
                    Process p = Process.Start(psi);
                }
                else
                {
                    Process.Start(erpappnamepath);
                }
                Thread.Sleep(Convert.ToInt32(waitappload));
                Log.Information("Done waiting app to opened.");

                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Quitting.. End of OpenApp function -> {ex.Message} !!!");
                return false;
            }
        }

        static bool OpenDB(out int errStep)
        {
            Point pnt = new Point(0, 0);
            bool isFound = false;
            errStep = 1;
            try
            {
                isFound = findimage("1-1.opendata", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;
                Thread.Sleep(2000);

                isFound = findimage("1-2.localdatabase", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                errStep += 1;
                Thread.Sleep(2000);


                isFound = findimage("1-3.databasename", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                    errStep += 1;
                    Thread.Sleep(1000);
                    iSim.Mouse.LeftButtonClick();
                    errStep += 1;
                    Thread.Sleep(1000);
                    iSim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    errStep += 1;
                    Thread.Sleep(500);
                    iSim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.VK_A);
                    errStep += 1;
                    Thread.Sleep(500);
                    iSim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    errStep += 1;
                    Thread.Sleep(1000);
                    iSim.Keyboard.TextEntry($"{dbpath}{dbname}");
                }
                else
                {
                    return false;
                }

                isFound = findimage("1-4.selectokdatabase", out pnt);
                if (isFound)
                {
                    SimulateMouseClick(pnt, leftClick.sngl);
                }
                else
                {
                    return false;
                }
                Thread.Sleep(12500);

                return true;
            }
            catch
            {
                Log.Information("Quitting.. End of open DB automation function !!");
                return false;
            }
        }
    }
