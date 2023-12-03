using Playnite.SDK;
using Playnite.SDK.Events;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace RainmeterExporter
{
    public class RainmeterExporter : GenericPlugin
    {

        private RainmeterExporterSettingsViewModel settings { get; set; }

        public override Guid Id { get; } = Guid.Parse("54bf64c6-c453-4cbc-92f8-4960b56f930e");

        public RainmeterExporter(IPlayniteAPI api) : base(api)
        {
            settings = new RainmeterExporterSettingsViewModel(this);
            Properties = new GenericPluginProperties
            {
                HasSettings = false
            };
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool CreateSymbolicLink(
        string lpSymlinkFileName, string lpTargetFileName, SymbolicLinkType dwFlags);

        enum SymbolicLinkType
        {
            File = 0,
            Directory = 1
        }

        public override IEnumerable<MainMenuItem> GetMainMenuItems(GetMainMenuItemsArgs args)
        {
            return new List<MainMenuItem>
    {
        new MainMenuItem
        {
            Description = ResourceProvider.GetString("RainmeterExporterAdvanced_MenuItemDescriptionOpenExportWindow"),
            MenuSection = "@Rainmeter Exporter",
            Action = _ =>
            {
                ExportGameList();
                ExportCoverStyles();
            }
        }
    };
        }

        private void ExportGameList()
        {
            IPlayniteAPI p = PlayniteApi;

            FilterPresetSettings filterSettings = new FilterPresetSettings { IsInstalled = true };
            IEnumerable<Game> gameFiltered = p.Database.GetFilteredGames(filterSettings);
            List<Game> gamesList = gameFiltered
                .Where(game => !game.Hidden)
                .OrderBy(game => game.Name)
                .ThenBy(game => game.Genres?.FirstOrDefault())
                .ToList();

            string playniteInstallDir = p.Paths.ApplicationPath;
            string selectedPathMy = GetRainmeterFilePath("Gamelist.inc");

            using (StreamWriter writer = new StreamWriter(selectedPathMy))
            {
                writer.WriteLine("[Variables]");
                writer.WriteLine("TotalGame=" + gamesList.Count);
                writer.WriteLine("");

                for (int i = 0; i < gamesList.Count; i++)
                {
                    WriteGameInfo(writer, gamesList[i], i + 1, playniteInstallDir);
                }

                writer.WriteLine("DynamicVariables = 1");
            }
        }

        private void ExportCoverStyles()
        {
            List<Game> gamesList = GetFilteredGames();
            string selectedPathCover = GetRainmeterFilePath("Cover.inc");

            using (StreamWriter writer = new StreamWriter(selectedPathCover))
            {
                for (int i = 0; i < gamesList.Count; i++)
                {
                    WriteCoverStyle(writer, i + 1);
                }
            }
        }

        private void WriteGameInfo(StreamWriter writer, Game game, int index, string playniteInstallDir)
        {
            writer.WriteLine($"Gamename{index}={game.Name}");
            writer.WriteLine($"Gamecover{index}=default.png");

            string backgroundPath = game.BackgroundImage;

            if (backgroundPath != null)
            {
                string result = Path.GetFileName(backgroundPath);
                string imagePath = Path.Combine(playniteInstallDir, "library", "files", backgroundPath);
                string destinationImagePath = GetRainmeterFilePath("Spectrum", "Cover", result);

                try
                {
                    if (File.Exists(imagePath))
                    {
                        writer.WriteLine($"Gamewall{index}={result}");
                        CreateSymbolicLink(destinationImagePath, imagePath, SymbolicLinkType.File);
                        Console.WriteLine($"Created symlink {destinationImagePath} to {imagePath}");
                    }
                    else
                    {
                        Console.WriteLine($"Source file {imagePath} does not exist.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating symlink: {ex.Message}");
                }
            }
            else
            {
                writer.WriteLine($"Gamewall{index}=default.jpg");
            }

            string installPath = game.InstallDirectory;
            WriteGameDir(writer, game, index, installPath);
            writer.WriteLine("");
        }

        private void WriteGameDir(StreamWriter writer, Game game, int index, string installPath)
        {
            if (game.GameActions != null)
            {
                string originalString = $"{game.GameActions[0].Path} {game.GameActions[0].Arguments}";
                string modifiedString = originalString.Replace("{InstallDir}", installPath);
                writer.WriteLine($"Gamedir{index}={modifiedString}");
            }
            else
            {
                switch (game.Source.Name)
                {
                    case "Steam":
                        writer.WriteLine($"Gamedir{index}=steam://rungameid/{game.GameId}");
                        break;
                    case "Epic":
                        writer.WriteLine($"Gamedir{index}=com.epicgames.launcher://apps/{game.GameId}?action=launch&silent=true");
                        break;
                    case "Uplay":
                        writer.WriteLine($"Gamedir{index}=uplay://launch/{game.GameId}");
                        break;
                    default:
                        writer.WriteLine($"Gamedir{index}=");
                        break;
                }
            }
        }

        private void WriteCoverStyle(StreamWriter writer, int index)
        {
            writer.WriteLine($"[Game{index}]");
            writer.WriteLine("Meter=Image");
            writer.WriteLine("MeterStyle=GameIcon");
            writer.WriteLine("X=0");
            writer.WriteLine("Y=0");
            writer.WriteLine($"MouseOverAction=[Play \"#@#Sounds\\Hover.wav\"][!SetVariable Select \"{index}\"][!CommandMeasure \"Script\" \"animation_update()\"]");
            writer.WriteLine($"LeftMouseUpAction=!Execute [!CommandMeasure InputExecute \"Execute {index}\"][!CommandMeasure Exit \"Execute {index}\" \"ElegantLauncher\"]");
            writer.WriteLine("");

            writer.WriteLine($"[Title{index}]");
            writer.WriteLine("Meter=String");
            writer.WriteLine("MeterStyle=GameTitle");
            writer.WriteLine($"Text=#Gamename{index}#");
            writer.WriteLine("");
        }

        private List<Game> GetFilteredGames()
        {
            IPlayniteAPI p = PlayniteApi;
            FilterPresetSettings filterSettings = new FilterPresetSettings { IsInstalled = true };
            return p.Database.GetFilteredGames(filterSettings)
                .Where(game => !game.Hidden)
                .OrderBy(game => game.Name)
                .ThenBy(game => game.Genres?.FirstOrDefault())
                .ToList();
        }

        private string GetRainmeterFilePath(params string[] pathSegments)
        {
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string rainmeterPath = Path.Combine(documentsPath, "Rainmeter", "Skins", "ElegantLauncher", "@Resources");
            return Path.Combine(rainmeterPath, Path.Combine(pathSegments));
        }


        public override ISettings GetSettings(bool firstRunSettings)
        {
            return settings;
        }

        public override UserControl GetSettingsView(bool firstRunSettings)
        {
            return new RainmeterExporterSettingsView();
        }
    }
}