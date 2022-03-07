﻿using FuzzySharp;
using IniParser;
using IniParser.Model;
using Playnite.SDK;
using Playnite.SDK.Data;
using Playnite.SDK.Events;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using SpecialKHelper.Common;
using SpecialKHelper.Models;
using SpecialKHelper.ViewModels;
using SpecialKHelper.Views;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Process = System.Diagnostics.Process;

namespace SpecialKHelper
{
    public class SpecialKHelper : GenericPlugin
    {
        private static readonly ILogger logger = LogManager.GetLogger();
        private readonly string emptyReshadePreset;
        private readonly string reshadeBase;
        private readonly string reshadeIniPath;
        private readonly FileIniDataParser iniParser;
        private readonly string pluginInstallPath;
        private readonly string skifPath;
        private readonly string defaultConfigPath;
        private static readonly Regex reshadeTechniqueRegex = new Regex(@"technique ([^\s]+)", RegexOptions.Compiled);

        private SpecialKHelperSettingsViewModel settings { get; set; }

        public override Guid Id { get; } = Guid.Parse("71349310-9ed8-4bf5-8bf2-e92cdb222748");

        public SpecialKHelper(IPlayniteAPI api) : base(api)
        {
            pluginInstallPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            skifPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "My Mods", "SpecialK");
            defaultConfigPath = Path.Combine(skifPath, "Global", "default_SpecialK.ini");
            emptyReshadePreset = Path.Combine(pluginInstallPath, "Resources", "ReshadeDefaultPreset.ini");
            reshadeBase = Path.Combine(skifPath, @"PlugIns\ThirdParty\ReShade");
            reshadeIniPath = Path.Combine(reshadeBase, "ReShade.ini");

            iniParser = new FileIniDataParser();
            iniParser.Parser.Configuration.AssigmentSpacer = string.Empty;

            settings = new SpecialKHelperSettingsViewModel(this);
            Properties = new GenericPluginProperties
            {
                HasSettings = true
            };
        }

        public string GetReshadeTechniqueSorting()
        {
            var shadersDirectory = Path.Combine(reshadeBase, "reshade-shaders", "Shaders");
            if (!Directory.Exists(shadersDirectory))
            {
                logger.Warn($"Reshade Shaders directory not found in {shadersDirectory}");
                return string.Empty;
            }

            var fxFiles = Directory.GetFiles(shadersDirectory, "*.fx", SearchOption.AllDirectories);
            var techniqueList = new List<string>();
            foreach (var fxFile in fxFiles)
            {
                var fxContent = File.ReadAllText(fxFile);
                var result = reshadeTechniqueRegex.Match(fxContent);
                if (result.Success)
                {
                    techniqueList.Add($"{result.Groups[1]}@{Path.GetFileName(fxFile)}");
                }
            }

            if (techniqueList.Count == 0)
            {
                return string.Empty;
            }

            techniqueList.Sort((x, y) => x.CompareTo(y));
            return string.Join(",", techniqueList);
        }

        public override void OnGameInstalled(OnGameInstalledEventArgs args)
        {
            // Add code to be executed when game is finished installing.
        }

        public override void OnGameStarted(OnGameStartedEventArgs args)
        {
            // Add code to be executed when game is started running.
        }

        public override void OnGameStarting(OnGameStartingEventArgs args)
        {
            var game = args.Game;
            var cpuArchitectures = new string[] { "32", "64" };
            var startServices = GetShouldStartService(game);

            foreach (var cpuArchitecture in cpuArchitectures)
            {
                if (startServices)
                {
                    StartSpecialkService(cpuArchitecture, skifPath);
                }
                else
                {
                    //Check if leftover service is running and close it
                    StopSpecialkService(cpuArchitecture, skifPath);
                }
            }

            if (!startServices)
            {
                return;
            }

            ValidateDefaultProfile();
            ConfigureSteamApiInject(game);
            ValidateReshadeConfiguration(game);
        }

        private void ValidateReshadeConfiguration(Game game)
        {
            var gameReshadePresetSubPath = Path.Combine("reshade-presets", game.Id.ToString() + ".ini");
            var gameReshadePreset = Path.Combine(reshadeBase, gameReshadePresetSubPath);

            if (!File.Exists(gameReshadePreset))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(gameReshadePreset));
                var techniqueSortingLine = GetReshadeTechniqueSorting();
                if (!techniqueSortingLine.IsNullOrEmpty())
                {
                    File.WriteAllText(gameReshadePreset, $"PreprocessorDefinitions=\nTechniques=\nTechniqueSorting={techniqueSortingLine}", Encoding.UTF8);
                }
                else
                {
                    File.Copy(emptyReshadePreset, gameReshadePreset, true);
                }
            }

            if (File.Exists(reshadeIniPath))
            {
                IniData ini = iniParser.ReadFile(reshadeIniPath);
                var updatedValues = 0;
                updatedValues += ValidateIniValue(ini, "GENERAL", "PresetPath", ".\\" + gameReshadePresetSubPath);
                updatedValues += ValidateIniValue(ini, "APP", "ForceVSync", "0");
                updatedValues += ValidateIniValue(ini, "APP", "ForceWindowed", "0");
                updatedValues += ValidateIniValue(ini, "APP", "ForceFullscreen", "0");
                updatedValues += ValidateIniValue(ini, "APP", "ForceResolution", "0,0");
                updatedValues += ValidateIniValue(ini, "APP", "Force10BitFormat", "0");

                if (updatedValues > 0)
                {
                    iniParser.WriteFile(reshadeIniPath, ini, Encoding.UTF8);
                }
            }
        }

        private void ValidateDefaultProfile()
        {
            if (!File.Exists(defaultConfigPath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(defaultConfigPath));
                File.Create(defaultConfigPath).Dispose();
                logger.Info($"Created default profile file blank file in {defaultConfigPath} since it was missing");
            }

            IniData ini = iniParser.ReadFile(defaultConfigPath);
            var updatedValues = 0;
            if (settings.Settings.EnableStOverlayOnNewProfiles)
            {
                updatedValues += ValidateIniValue(ini, "Steam.System", "PreLoadSteamOverlay", "true");
            }
            else
            {
                updatedValues += ValidateIniValue(ini, "Steam.System", "PreLoadSteamOverlay", "false");
            }

            if (settings.Settings.EnableReshadeOnNewProfiles)
            {
                updatedValues += ValidateIniValue(ini, "Render.FrameRate", "SleeplessRenderThread", "false");
                updatedValues += ValidateIniValue(ini, "Render.OSD", "ShowInVideoCapture", "false");

                updatedValues += ValidateIniValue(ini, "Import.ReShade64", "Architecture", "x64");
                updatedValues += ValidateIniValue(ini, "Import.ReShade64", "Role", "ThirdParty");
                updatedValues += ValidateIniValue(ini, "Import.ReShade64", "When", "PlugIn");
                updatedValues += ValidateIniValue(ini, "Import.ReShade64", "Filename", @"..\..\PlugIns\ThirdParty\ReShade\ReShade64.dll");

                updatedValues += ValidateIniValue(ini, "Import.ReShade32", "Architecture", "Win32");
                updatedValues += ValidateIniValue(ini, "Import.ReShade32", "Role", "ThirdParty");
                updatedValues += ValidateIniValue(ini, "Import.ReShade32", "When", "PlugIn");
                updatedValues += ValidateIniValue(ini, "Import.ReShade32", "Filename", @"..\..\PlugIns\ThirdParty\ReShade\ReShade32.dll");
            }
            else
            {
                updatedValues += ValidateIniValue(ini, "Import.ReShade64", "Filename", string.Empty);
                updatedValues += ValidateIniValue(ini, "Import.ReShade32", "Filename", string.Empty);
            }

            if (settings.Settings.SetDefaultFpsOnNewProfiles && settings.Settings.DefaultFpsLimit != 0)
            {
                updatedValues += ValidateIniValue(ini, "Render.FrameRate", "TargetFPS", settings.Settings.DefaultFpsLimit.ToString());
            }
            else
            {
                updatedValues += ValidateIniValue(ini, "Render.FrameRate", "TargetFPS", "0.0");
            }

            if (settings.Settings.DisableNvidiaBlOnNewProfiles)
            {
                updatedValues += ValidateIniValue(ini, "Compatibility.General", "DisableBloatWare_NVIDIA", "true");
            }
            else
            {
                updatedValues += ValidateIniValue(ini, "Compatibility.General", "DisableBloatWare_NVIDIA", "false");
            }

            if (settings.Settings.UseFlipModelOnNewProfiles)
            {
                updatedValues += ValidateIniValue(ini, "Render.DXGI", "UseFlipDiscard", "true");
            }
            else
            {
                updatedValues += ValidateIniValue(ini, "Render.DXGI", "UseFlipDiscard", "false");
            }

            if (updatedValues > 0)
            {
                iniParser.WriteFile(defaultConfigPath, ini, Encoding.UTF8);
                logger.Info($"Default ini validated and updated {updatedValues} new values");
            }
        }

        private int ValidateIniValue(IniData ini, string key, string subKey, string newValue)
        {
            var currentValue = ini[key][subKey];
            if (currentValue == null || currentValue != newValue)
            {
                ini[key][subKey] = newValue;
                logger.Info($"Default ini validated and updated value in [{key}]{subKey} to \"{newValue}\"");
                return 1;
            }

            return 0;
        }

        private bool GetShouldStartService(Game game)
        {
            if (game.Features != null)
            {
                if (settings.Settings.StopExecutionIfVac && game.Features.Any(x => x.Name == "Valve Anti-Cheat Enabled"))
                {
                    return false;
                }

                if (settings.Settings.SpecialKExecutionMode == SpecialKExecutionMode.Global)
                {
                    if (game.Features.Any(x => x.Name == "[SK] Global Mode Disable"))
                    {
                        return false;
                    }
                }
                else if (settings.Settings.SpecialKExecutionMode == SpecialKExecutionMode.Selective)
                {
                    if (!game.Features.Any(x => x.Name == "[SW] Selective Mode Enable"))
                    {
                        return false;
                    }
                }
            }

            if (settings.Settings.OnlyExecutePcGames && !IsGamePcGame(game))
            {
                return false;
            }

            return true;
        }

        private bool IsGamePcGame(Game game)
        {
            if (game.Platforms == null)
            {
                logger.Info($"Game {game.Name} doesn't have platforms set");
                return false;
            }
            else if (!game.Platforms.Any(x => x.Name == "PC (Windows)" ||
                      !x.SpecificationId.IsNullOrEmpty() && x.SpecificationId == "pc_windows"))
            {
                logger.Info($"Game {game.Name} is not PC platform");
                return false;
            }

            return true;
        }

        private bool GetIsGameInstallDirValid(Game game)
        {
            if (game.InstallDirectory.IsNullOrEmpty() || !Directory.Exists(game.InstallDirectory))
            {
                return false;
            }

            return true;
        }

        private bool ConfigureSteamApiInject(Game game)
        {
            if (SteamCommon.IsGameSteamGame(game))
            {
                return true;
            }

            var appIdTextPath = string.Empty;
            var isInstallDirValid = GetIsGameInstallDirValid(game);
            if (isInstallDirValid)
            {
                appIdTextPath = Path.Combine(game.InstallDirectory, "steam_appid.txt");
                if (File.Exists(appIdTextPath))
                {
                    return true;
                }
            }

            var previousId = string.Empty;
            var historyFlagFile = Path.Combine(GetPluginUserDataPath(), "attempted" + game.Id.ToString());
            if (File.Exists(historyFlagFile))
            {
                previousId = File.ReadAllText(historyFlagFile).Trim();
                logger.Warn($"Detected attempt flag file for game {game.Name} in {historyFlagFile}. Previous Id: {previousId}");
            }

            var steamId = "0";
            if (!previousId.IsNullOrEmpty())
            {
                // We use the previously found Id to not have to search again
                steamId = previousId;
            }
            else if (IsGamePcGame(game))
            {
                var isBackgroundDownload = false;
                if (PlayniteApi.ApplicationInfo.Mode == ApplicationMode.Fullscreen)
                {
                    isBackgroundDownload = true;
                }

                var steamIdSearch = GetSteamIdFromSearch(game, isBackgroundDownload, true);
                if (!steamId.IsNullOrEmpty())
                {
                    steamId = steamIdSearch;
                }
            }

            if (isInstallDirValid && !appIdTextPath.IsNullOrEmpty())
            {
                try
                {
                    File.WriteAllText(appIdTextPath, steamId);
                }
                catch (Exception e)
                {
                    logger.Error(e, $"Error while creating steam id file in {appIdTextPath}");
                }
            }

            // As an alternative in case file creation fails or created steam id file
            // is not used, we attempt to modify the default profile
            // so the id gets copied if a new profile is created
            if (File.Exists(defaultConfigPath))
            {
                IniData ini = iniParser.ReadFile(defaultConfigPath);
                var currentAppId = ini["Steam.System"]["AppID"];
                if (currentAppId == null || currentAppId != steamId)
                {
                    ini["Steam.System"]["AppID"] = steamId;
                    iniParser.WriteFile(defaultConfigPath, ini, Encoding.UTF8);
                }
            }

            // Flag file so we don't attempt to search again in future startups.
            if (!File.Exists(historyFlagFile))
            {
                File.WriteAllText(historyFlagFile, steamId, Encoding.UTF8);
            }
            
            return true;
        }

        public override void OnGameStopped(OnGameStoppedEventArgs args)
        {
            var cpuArchitectures = new string[] { "32", "64" };
            foreach (var cpuArchitecture in cpuArchitectures)
            {
                StopSpecialkService(cpuArchitecture, skifPath);
            }
        }

        private bool StartSpecialkService(string cpuArchitecture, string skifPath)
        {
            var dllPath = Path.Combine(skifPath, "SpecialK" + cpuArchitecture + ".dll");
            if (!File.Exists(dllPath))
            {
                logger.Info($"Special K dll not found in {dllPath}");
                PlayniteApi.Notifications.Add(new NotificationMessage(
                    "sk_dll_notfound" + cpuArchitecture,
                    $"Special K dll not found in {dllPath}",
                    NotificationType.Error
                ));

                return false;
            }

            var exePath = "rundll32.exe";
            var exe64Path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "rundll32.exe");
            if (cpuArchitecture == "64" && File.Exists(exe64Path))
            {
                exePath = exe64Path;
            }

            var sinfo = new ProcessStartInfo();
            sinfo.UseShellExecute = true; // Do not wait - make the process stand alone
            sinfo.FileName = exePath;
            sinfo.WorkingDirectory = skifPath;
            sinfo.Arguments = $"\"{dllPath}\",RunDLL_InjectionManager Install";
            Process.Start(sinfo);

            var eventName = "Local\\SK_GlobalHookTeardown" + cpuArchitecture;
            var i = 0;

            while (i < 12)
            {
                Thread.Sleep(40);
                if (EventWaitHandle.TryOpenExisting(eventName, out var eventWaitHandle))
                {
                    eventWaitHandle.Close();
                    eventWaitHandle.Dispose();
                    logger.Info($"Special K global service for {dllPath} started");
                    return true;
                }
                else
                {
                    i++;
                }
            }

            logger.Info($"Special K global service for \"{eventName}\" could not be opened");
            return false;
        }

        private bool StopSpecialkService(string cpuArchitecture, string skifPath)
        {
            var dllPath = Path.Combine(skifPath, "SpecialK" + cpuArchitecture + ".dll");
            if (!File.Exists(dllPath))
            {
                logger.Info($"Special K dll not found in {dllPath}");
                return false;
            }

            var exePath = "rundll32.exe";
            var exe64Path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "rundll32.exe");
            if (cpuArchitecture == "64" && File.Exists(exe64Path))
            {
                exePath = exe64Path;
            }

            try
            {
                var sinfo = new ProcessStartInfo();
                sinfo.UseShellExecute = true; // Do not wait - make the process stand alone
                sinfo.FileName = exePath;
                sinfo.WorkingDirectory = skifPath;
                sinfo.Arguments = $"\"{dllPath}\",RunDLL_InjectionManager Remove";
                Process.Start(sinfo);

                logger.Info($"Special K {cpuArchitecture} service has been removed");
                return true;
            }
            catch
            {
                logger.Error($"Special K {cpuArchitecture} service could not be removed");
                return false;
            }
        }

        public string GetSteamIdFromSearch(Game game, bool isBackgroundDownload, bool matchFuzzyMethods = false)
        {
            var normalizedName = game.Name.NormalizeGameName();
            var results = SteamCommon.GetSteamSearchResults(normalizedName);

            results.ForEach(a => a.Name = a.Name.NormalizeGameName());

            // Try to see if there's an exact match, to not prompt the user unless needed
            var matchingGameName = normalizedName.GetMatchModifiedName();
            var exactMatch = results.FirstOrDefault(x => x.Name.GetMatchModifiedName() == matchingGameName);

            // Automatic match method 1: Remove all symbols
            if (exactMatch != null)
            {
                logger.Info($"Found steam id for game {game.Name} via method 1, Id: {exactMatch.GameId}, Match: {exactMatch.Name}");
                return exactMatch.GameId;
            }

            var currentLevenshteinId = string.Empty;
            var currentDistance = 99;
            if (matchFuzzyMethods)
            {
                // Automatic match method 2: Fuzzy search
                var currentFuzzyValue = 0;
                var currentFuzzyId = string.Empty;
                foreach (var result in results)
                {
                    var proximity = Fuzz.Ratio(normalizedName.ToLower(), result.Name.ToLower());
                    if (proximity > currentFuzzyValue)
                    {
                        currentFuzzyValue = proximity;
                        currentFuzzyId = result.GameId;
                    }
                }

                if (!currentFuzzyId.IsNullOrEmpty() && currentFuzzyValue > 88)
                {
                    logger.Info($"Found steam id for game {game.Name} via method 2, Id: {currentFuzzyId}, Proximity: {currentFuzzyValue}");
                    return currentFuzzyId;
                }

                // Automatic match method 3: LevenshteinDistance
                foreach (var result in results)
                {
                    var distance = LevenshteinDistance.Distance(normalizedName.ToLower(), result.Name.ToLower());
                    if (distance < currentDistance)
                    {
                        currentDistance = distance;
                        currentLevenshteinId = result.GameId;
                    }
                }

                if (!currentLevenshteinId.IsNullOrEmpty() && currentDistance < 3)
                {
                    logger.Info($"Found steam id for game {game.Name} via method 3, Id: {currentLevenshteinId}, Distance: {currentDistance}");
                    return currentLevenshteinId;
                }
            }

            if (!isBackgroundDownload)
            {
                var selectedGame = PlayniteApi.Dialogs.ChooseItemWithSearch(
                    results.Select(x => new GenericItemOption(x.Name, x.GameId)).ToList(),
                    (a) => SteamCommon.GetSteamSearchGenericItemOptions(a),
                    normalizedName,
                    "Select game to use for Steam configuration or cancel to use a generic option");
                if (selectedGame != null)
                {
                    return selectedGame.Description;
                }
            }

            // As a last resort, if the search was background download and fuzzy methods were used,
            // return the best Levenshtein result. Better than nothing I guess
            // and will probably be the appropiate result
            if (isBackgroundDownload && !currentLevenshteinId.IsNullOrEmpty())
            {
                logger.Info($"Found steam id for game {game.Name} via method 4, Id: {currentLevenshteinId}, Distance: {currentDistance}");
                return currentLevenshteinId;
            }

            logger.Info($"Steam id for game {game.Name} not found");
            return null;
        }

        public override void OnGameUninstalled(OnGameUninstalledEventArgs args)
        {
            // Add code to be executed when game is uninstalled.
        }

        public override void OnApplicationStarted(OnApplicationStartedEventArgs args)
        {
            // Add code to be executed when Playnite is initialized.
        }

        public override void OnApplicationStopped(OnApplicationStoppedEventArgs args)
        {
            // Add code to be executed when Playnite is shutting down.
        }

        public override void OnLibraryUpdated(OnLibraryUpdatedEventArgs args)
        {
            // Add code to be executed when library is updated.
        }

        public override ISettings GetSettings(bool firstRunSettings)
        {
            return settings;
        }

        private void OpenEditorWindow(string searchTerm = null)
        {
            var window = PlayniteApi.Dialogs.CreateWindow(new WindowCreationOptions
            {
                ShowMinimizeButton = false,
                ShowMaximizeButton = true
            });

            window.Height = 700;
            window.Width = 900;
            window.Title = "Special K Profile Editor";

            window.Content = new SpecialKProfileEditorView();
            window.DataContext = new SpecialKProfileEditorViewModel(PlayniteApi, iniParser, skifPath, searchTerm);
            window.Owner = PlayniteApi.Dialogs.GetCurrentAppWindow();
            window.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            window.ShowDialog();
        }

        public override IEnumerable<MainMenuItem> GetMainMenuItems(GetMainMenuItemsArgs args)
        {
            return new List<MainMenuItem>
            {
                new MainMenuItem
                {
                    Description = "Open Special K Profiles Editor",
                    MenuSection = "@Special K Helper",
                    Action = o => {
                        OpenEditorWindow();
                    }
                },
            };
        }

        public override UserControl GetSettingsView(bool firstRunSettings)
        {
            return new SpecialKHelperSettingsView();
        }
    }
}