using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using CommandLine;
using CSCore.CoreAudioAPI;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using PlexAutoIntroSkip.Models;

namespace PlexAutoIntroSkip
{
    public class Program
    {
        public class ProgramOptions
        {
            [Option('d', "debug", Required = false,
                HelpText = "Show console window.")]
            public bool ShowConsoleWindow { get; set; }

            [Option('w', "wait-time", Required = false, Default = 2500,
                HelpText = "Time to wait after Skip Button becomes visible before clicking.")]
            public int SkipButtonWaitTime { get; set; }

            [Option('v', "auto-volume", Required = false, Default = false)]
            public bool AutoVolume { get; set; }

            [Value(0, MetaName = "plex-url", HelpText = "Plex URL to use.")]
            public string PlexUrl { get; set; }
        }
        

        /// <summary>
        /// Sets the specified window's show state.
        /// </summary>
        /// <remarks>
        /// See <see href="https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow">ShowWindow</see> MS Docs for more information.
        /// </remarks>
        /// <param name="hWnd">A handle to the window.</param>
        /// <param name="nCmdShow">Controls how the window is to be shown.</param>
        /// <returns>
        /// If the window was previously visible, the return value is nonzero.
        /// If the window was previously hidden, the return value is zero.
        /// </returns>
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        /// <summary>
        /// Program entry point.
        /// </summary>
        /// <param name="args">Command-line arguments.</param>
        public static void Main(string[] args)
        {

#if DEBUG
            args = new[] {"", "--v"};
#endif


            var options = GetProgramOptions(args);
            var hWnd = Process.GetCurrentProcess().MainWindowHandle;
            var edgeProcessName = "msedge";

            // Called using own process window?
            if (options.ShowConsoleWindow == false && hWnd.ToInt32() != 0)
            {
                // Hide console application's console window.
                ShowWindow(hWnd, 0);
            }

            var edgeOptions = new EdgeOptions();
            edgeOptions.UseChromium = true;

            // Disable "Chrome is being controlled by automated test software" infobar.
            edgeOptions.AddExcludedArgument("enable-automation");
            edgeOptions.AddAdditionalOption("useAutomationExtension", false);

            edgeOptions.AddArguments(
                $"user-data-dir={Directory.GetCurrentDirectory()}\\User Data",
                "profile-directory=Profile 1",
                $"app={options.PlexUrl}");

            var edgeProcessIds = Process.GetProcessesByName(edgeProcessName).Select(p => p.Id);
            var service = EdgeDriverService.CreateDefaultService();
            var driver = new EdgeDriver(service, edgeOptions);
            var browserProcessId = Process.GetProcessesByName(edgeProcessName).Select(p => p.Id)
                .Except(edgeProcessIds)
                .First();

            if (options.AutoVolume)
            {
                AudioControlSettings audioControlSettings = new AudioControlSettings { AudioLogs = new List<double>(), VolumeTolerance = 0.05, VolumeGatherTimeInSeconds = 5};

                Task.Factory.StartNew(() =>
                {
                    Debug.WriteLine("Searching for audio process...");
                    while (audioControlSettings.AudioProcess == null)
                    {

                        // Gets all processes that have audio
                        using var audioSessionManager = AudioSessionManager2.FromMMDevice(new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia));

                        // Loop through each audio session to find the process associated with the webdriver
                        foreach (var session in audioSessionManager.GetSessionEnumerator().Where(s => s.QueryInterface<AudioSessionControl2>().Process.ProcessName.Contains("edge")))
                        {

                            var tempAudioProcessId = session.QueryInterface<AudioSessionControl2>().ProcessID;

                            // Get the parent process ID of the audio session
                            using ManagementObject managementObject = new ManagementObject("win32_process.handle='" + tempAudioProcessId + "'");
                            managementObject.Get();

                            if(Convert.ToInt32(managementObject["ParentProcessId"]) != browserProcessId) continue;

                            Debug.WriteLine("Audio process found!  PID: {0}", tempAudioProcessId);

                            audioControlSettings.AudioProcessID = tempAudioProcessId;
                            audioControlSettings.AudioProcess = audioSessionManager.GetSessionEnumerator().FirstOrDefault(s => s.QueryInterface<AudioSessionControl2>().ProcessID == tempAudioProcessId)?.QueryInterface<AudioSessionControl2>();

                            managementObject.Dispose();
                        }
                        audioSessionManager.Dispose();
                    }

                    GatherAudioSamples(audioControlSettings, browserProcessId);
                    audioControlSettings.VolumeTarget = audioControlSettings.AverageVolume;


                    var volumeSliderThumb = driver.FindElementByClassName("VerticalSlider-thumb-1y9RTp");
                    int volumeSliderPos = int.TryParse(volumeSliderThumb.GetAttribute("aria-valuenow"), out int result) ? result : 50;
                    audioControlSettings.VolumeSliderPosition = volumeSliderPos;


                    while (ProcessExistsById(browserProcessId))
                    {
                        try
                        {
                            ReadCurrentAudio(audioControlSettings, driver, browserProcessId);
                            Thread.Sleep(500);
                        }
                        catch (Exception ex)
                        {
                            LogException(ex);
                        }
                    }
                });
            }

            while (ProcessExistsById(browserProcessId))
            {
                try
                {
                    MainProgramLoop(driver, options, browserProcessId);
                }
                catch (Exception exception)
                {
                    LogException(exception);
                }
            }

            Process.GetProcessById(service.ProcessId).Kill();
        }


        private static void ReadCurrentAudio(AudioControlSettings audioControlSettings, EdgeDriver driver, int browserPID)
        {

            try
            {
                float? audioLog = audioControlSettings.AudioProcess.QueryInterface<AudioMeterInformation>()?.PeakValue;
                // ReSharper disable once CompareOfFloatsByEqualityOperator
                if (audioLog == null || audioLog == 0 || audioLog.ToString().Contains("E")) return;

                //var volumeSlider = driver.FindElementByClassName("PlayerControls-volumeSlider-qbq5Ws");
                var volumeSliderThumb = driver.FindElementByClassName("VerticalSlider-thumb-1y9RTp");
                int? currentVolumeSliderPos = int.TryParse(volumeSliderThumb.GetAttribute("aria-valuenow"), out int result) ? result : (int?)null;
                if (!currentVolumeSliderPos.HasValue) return;

                // If the slider moved by user get the new target volume
                if (audioControlSettings.VolumeSliderPosition != currentVolumeSliderPos)
                {
                    GatherAudioSamples(audioControlSettings, browserPID);
                    audioControlSettings.VolumeTarget = audioControlSettings.AverageVolume;
                    audioControlSettings.VolumeSliderPosition = (int) currentVolumeSliderPos;
                    return;
                }

                audioControlSettings.AudioLogs.Add((double) audioLog);
                audioControlSettings.AudioLogs.RemoveAt(0);

                // Adjust audio to meet user settings
                if (audioControlSettings.AverageVolume - audioControlSettings.VolumeTolerance > audioControlSettings.VolumeTarget || audioControlSettings.AverageVolume + audioControlSettings.VolumeTolerance < audioControlSettings.VolumeTarget)
                {

                    //int pixelsPerPercent = 100 / volumeSlider.Size.Height;

                    Actions move = new Actions(driver);
                    //var moveSlider = move.DragAndDropToOffset(volumeSliderThumb,0, _audioLogs.Average() - .05 > _targetVolume ? pixelsPerPercent : -pixelsPerPercent).Build();
                    var moveSlider = move.SendKeys(audioControlSettings.AverageVolume - audioControlSettings.VolumeTolerance > audioControlSettings.VolumeTarget ? Keys.Down : Keys.Up).Build();
                    moveSlider.Perform();

                    // Set the new slider position
                    volumeSliderThumb = driver.FindElementByClassName("VerticalSlider-thumb-1y9RTp");
                    currentVolumeSliderPos = int.TryParse(volumeSliderThumb.GetAttribute("aria-valuenow"), out result) ? result : 0;
                    audioControlSettings.VolumeSliderPosition = (int) currentVolumeSliderPos;

                    GatherAudioSamples(audioControlSettings, browserPID);
                    Debug.WriteLine("Adjusted volume " + (audioControlSettings.AverageVolume - audioControlSettings.VolumeTolerance > audioControlSettings.VolumeTarget ? "down" : "up"));
                }
            }
            catch (Exception exception)
            {
                LogException(exception);
            }

            Debug.WriteLine("Average Volume: {0}", audioControlSettings.AverageVolume);
        }


        /// <summary>
        /// Run main program loop.
        /// </summary>
        /// <param name="driver"><see name="RemoteWebDriver"/> to be used to drive web browser.</param>
        /// <param name="options"><see name="ProgramOptions"/>.</param>
        /// <param name="browserProcessId">Web browser process ID to check if exists.</param>
        private static void MainProgramLoop(RemoteWebDriver driver, ProgramOptions options, int browserProcessId)
        {
            var waitDriver = new WebDriverWait(driver, TimeSpan.FromDays(365));
            var nullWebElement = new RemoteWebElement(driver, string.Empty);

            var skipIntroButtonXPath = "//button[text()='Skip Intro']";
            while (ProcessExistsById(browserProcessId))
            {
                // Waiting for Skip Intro button to be visible, or for browser to be closed manually.
                var skipIntroButton = (RemoteWebElement)waitDriver.Until(webDriver =>
                    ProcessExistsById(browserProcessId)
                        ? webDriver.FindElement(By.XPath(skipIntroButtonXPath))
                        : nullWebElement);

                // Skip Intro button will only be equal to `nullWebElement` if the browser process
                // no longer exists.
                if (Equals(skipIntroButton, nullWebElement))
                {
                    break;
                }

                // Skip Intro button is visible before intro starts, so wait for the actual intro to start.
                Thread.Sleep(options.SkipButtonWaitTime);

                skipIntroButton.Click();

                // Waiting for Skip Intro button to no longer be visible.
                waitDriver.Until(webDriver =>
                {
                    try
                    {
                        webDriver.FindElement(By.XPath(skipIntroButtonXPath));

                        return false;
                    }
                    catch
                    {
                        return true;
                    }
                });
            }
        }

        /// <summary>
        /// Parse <paramref name="args"/> into <see cref="ProgramOptions"/>.
        /// </summary>
        /// <param name="args">Command-line arguments.</param>
        /// <returns><see cref="ProgramOptions"/></returns>
        private static ProgramOptions GetProgramOptions(string[] args)
        {
            ProgramOptions options = null;
            CommandLine.Parser.Default.ParseArguments<ProgramOptions>(args).WithParsed(parsedOptions => options = parsedOptions);

            return options;
        }

        /// <summary>
        /// Gathers audio samples for a specified amount of time.
        /// </summary>
        /// <param name="audioControlSettings"></param>
        /// <param name="browserPID"></param>
        private static void GatherAudioSamples(AudioControlSettings audioControlSettings, int browserPID)
        {
            audioControlSettings.AudioLogs.Clear();

            // Gather enough data for an average volume
            while (audioControlSettings.AudioLogs.Count < audioControlSettings.VolumeGatherTimeInSeconds)
            {
                try
                {
                    // Make sure process is still running
                    if (!ProcessExistsById(browserPID)) return;

                    using var audioMeterInformation = audioControlSettings.AudioProcess?.QueryInterface<AudioMeterInformation>();

                    if (audioMeterInformation?.PeakValue == null || Math.Abs(audioMeterInformation.PeakValue) < 0 || audioMeterInformation.PeakValue.ToString(CultureInfo.InvariantCulture).Contains("E")) continue;

                    audioControlSettings.AudioLogs.Add(audioMeterInformation.PeakValue);

                    if (audioControlSettings.AudioLogs.Count < audioControlSettings.VolumeGatherTimeInSeconds) Thread.Sleep(1000);
                }
                catch (Exception exception)
                {
                    LogException(exception);
                }
            }

        }

        /// <summary>
        /// Recursively write <see name="Exception"/> to error log file.
        /// </summary>
        /// <param name="exception">Root <see name="Exception"/>.</param>
        private static void LogException(Exception exception)
        {
            using (var writer = new StreamWriter("error.log", append: true))
            {
                writer.WriteLine("--------------------------------------------------");
                writer.WriteLine($"[{DateTime.Now}]");
                writer.WriteLine();

                while (exception != null)
                {
                    writer.WriteLine(exception.GetType().FullName);
                    writer.WriteLine(exception.Message);
                    writer.WriteLine(exception.StackTrace);
                    writer.WriteLine();

                    exception = exception.InnerException;
                }
            }
        }

        /// <summary>
        /// Check if the process exists by given <paramref name="processId"/>.
        /// </summary>
        /// <param name="processId">Process ID.</param>
        /// <returns>True if process exists, otherwise false.</returns>
        private static bool ProcessExistsById(int processId) => Process.GetProcesses().Any(p => p.Id == processId);
    }
}
