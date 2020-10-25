using DiscordRPC;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace NetEaseMusic_DiscordRPC
{
    class Program
    {
        static void Main()
        {
            // check run once
            _ = new Mutex(true, "NetEase Cloud Music RPC", out var allow);
            if (!allow)
            {
                MessageBox.Show("NetEase Cloud Music RPC is already running.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(-1);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var ApplicationId = "750224620476694658";

            // Menu setup
            var notifyMenu = new ContextMenu();

            var actiButton = new MenuItem("启用");
            var themeMenu = new MenuItem("个性化");
            var settingMenu = new MenuItem("设置");
            var linkMenu = new MenuItem("相关链接");
            var exitButton = new MenuItem("退出");

            var playingStats = new MenuItem("暂停时显示");
            var defDark = new MenuItem("Dark");
            var defDiscord = new MenuItem("Discord");
            var defNetease = new MenuItem("Netease Red");
            var defWhite = new MenuItem("Netease White");
            var showSmallImg = new MenuItem("显示状态");

            var autoButton = new MenuItem("开机自启");
            var settingFullscreen = new MenuItem("在全屏状态下显示");
            var settingWhitelists = new MenuItem("忽略白名单");

            var mainLink = new MenuItem("主程序 项目地址");
            var thisLink = new MenuItem("此程序 项目地址");


            // Main menu
            notifyMenu.MenuItems.Add(actiButton);
            notifyMenu.MenuItems.Add("-");
            notifyMenu.MenuItems.Add(themeMenu);
            notifyMenu.MenuItems.Add(settingMenu);
            notifyMenu.MenuItems.Add("-");
            notifyMenu.MenuItems.Add(exitButton);

            // Theme submenu
            themeMenu.MenuItems.Add(playingStats);
            themeMenu.MenuItems.Add("-");
            themeMenu.MenuItems.Add(defDark);
            themeMenu.MenuItems.Add(defDiscord);
            themeMenu.MenuItems.Add(defNetease);
            themeMenu.MenuItems.Add(defWhite);
            themeMenu.MenuItems.Add("-");
            themeMenu.MenuItems.Add(showSmallImg);

            // Settings submenu
            settingMenu.MenuItems.Add(autoButton);
            settingMenu.MenuItems.Add(settingFullscreen);
            settingMenu.MenuItems.Add(settingWhitelists);
            settingMenu.MenuItems.Add("-");
            settingMenu.MenuItems.Add(linkMenu);

            // Links submenu
            linkMenu.MenuItems.Add(mainLink);
            linkMenu.MenuItems.Add(thisLink);

            // Check the saved option
            actiButton.Checked = Properties.Settings.Default.DefActive;
            if (Properties.Settings.Default.DefSkin == "default_dark")
                defDark.Checked = true;
            if (Properties.Settings.Default.DefSkin == "default_discord")
                defDiscord.Checked = true;
            if (Properties.Settings.Default.DefSkin == "default_netease")
                defNetease.Checked = true;
            if (Properties.Settings.Default.DefSkin == "default_white")
                defWhite.Checked = true;
            autoButton.Checked = AutoStart.Check();
            settingFullscreen.Checked = Properties.Settings.Default.FullscreenRun;
            settingWhitelists.Checked = Properties.Settings.Default.WhitelistsRun;
            playingStats.Checked = Properties.Settings.Default.showPaused;

            // Icon
            var notifyIcon = new NotifyIcon()
            {
                BalloonTipIcon = ToolTipIcon.Info,
                ContextMenu = notifyMenu,
                Text = "NetEase Cloud Music RPC",
                Icon = Properties.Resources.icon,
                Visible = true,
            };

            // First time popup
            if (Properties.Settings.Default.IsFirstTime)
            {
                notifyIcon.BalloonTipTitle = "NetEase Music RPC";
                notifyIcon.BalloonTipText = "插件已启动!";
                notifyIcon.ShowBalloonTip(5000);
                Thread.Sleep(5500);
                AutoStart.Set(); // Set Autostart
                autoButton.Checked = AutoStart.Check();
                notifyIcon.BalloonTipTitle = "NetEase Music RPC";
                notifyIcon.BalloonTipText = "应用首次运行, 已勾选开机自启";
                notifyIcon.ShowBalloonTip(5000);
                Properties.Settings.Default.IsFirstTime = false;
                Properties.Settings.Default.Save();
            }

            // Button clicked action/
            exitButton.Click += (sender, args) =>
            {
                Properties.Settings.Default.Save();
                notifyIcon.Visible = false;
                Thread.Sleep(100);
                Environment.Exit(0);
            };

            autoButton.Click += (sender, args) =>
            {
                if (AutoStart.Check())
                    AutoStart.Remove();
                else
                    AutoStart.Set();
                autoButton.Checked = AutoStart.Check();
            };

            actiButton.Click += (sender, args) =>
            {
                actiButton.Checked = !actiButton.Checked;
                Properties.Settings.Default.Save();
            };

            settingFullscreen.Click += (sender, args) =>
            {
                settingFullscreen.Checked = !settingFullscreen.Checked;
                Properties.Settings.Default.FullscreenRun = settingFullscreen.Checked;
                Properties.Settings.Default.Save();
            };

            settingWhitelists.Click += (sender, args) =>
            {
                settingWhitelists.Checked = !settingWhitelists.Checked;
                Properties.Settings.Default.WhitelistsRun = settingWhitelists.Checked;
                Properties.Settings.Default.Save();
            };

            defDark.Click += (sender, args) =>
            {
                defDark.Checked = true;
                Properties.Settings.Default.DefSkin = "default_dark";
                Properties.Settings.Default.Save();
                defDiscord.Checked = false;
                defNetease.Checked = false;
                defWhite.Checked = false;
            };

            defDiscord.Click += (sender, args) =>
            {
                defDiscord.Checked = true;
                Properties.Settings.Default.DefSkin = "default_discord";
                Properties.Settings.Default.Save();
                defDark.Checked = false;
                defNetease.Checked = false;
                defWhite.Checked = false;
            };

            defNetease.Click += (sender, args) =>
            {
                defNetease.Checked = true;
                Properties.Settings.Default.DefSkin = "default_netease";
                Properties.Settings.Default.Save();
                defDark.Checked = false;
                defDiscord.Checked = false;
                defWhite.Checked = false;
            };

            defWhite.Click += (sender, args) =>
            {
                defWhite.Checked = true;
                Properties.Settings.Default.DefSkin = "default_white";
                Properties.Settings.Default.Save();
                defDark.Checked = false;
                defDiscord.Checked = false;
                defNetease.Checked = false;
            };

            playingStats.Click += (sender, args) =>
            {
                playingStats.Checked = !playingStats.Checked;
                Properties.Settings.Default.showPaused = playingStats.Checked;
                Properties.Settings.Default.Save();
            };

            showSmallImg.Click += (sender, args) =>
            {
                showSmallImg.Checked = !showSmallImg.Checked;
                Properties.Settings.Default.showSmallImg = showSmallImg.Checked;
                Properties.Settings.Default.Save();
            };

            mainLink.Click += (sender, args) =>
            {
                Process.Start("https://github.com/Kxnrl/NetEase-Cloud-Music-DiscordRPC");
            };

            thisLink.Click += (sender, args) =>
            {
                Process.Start("https://github.com/GaXyRay/NetEase-Cloud-Music-DiscordRPC/");
            };

            // Run
            Task.Run(() =>
            {
                using var discord = new DiscordRpcClient(ApplicationId);
                discord.Initialize();
                var musicId = string.Empty;
                var playerState = false;
                var currentSong = string.Empty;
                var oldplaySong = string.Empty;
                var currentSing = string.Empty;
                var currentRate = 0.0;
                var maxSongLens = 0.0;
                var lastRate = -0.01;
                var diffRate = 0.01;
                var largeImg = Properties.Settings.Default.DefSkin;
                var largeImgText = "Netease Cloud Music";
                var albumArtUrl = string.Empty;
                

                while (true)
                {
                    // 用户就喜欢超低内存占用 但是实际上来说并没有什么卵用
                    GC.Collect();
                    GC.WaitForFullGCComplete();
                    Thread.Sleep(5000);

                    lastRate = currentRate;
                    var lastLens = maxSongLens;

                    Debug.Print($"{currentRate} | {lastRate} | {(diffRate)}");

                    if (!Win32Api.User32.GetWindowTitle("OrpheusBrowserHost", out var title, out var pid) || pid == 0)
                    {
                        Debug.Print($"player is not running");
                        playerState = false;
                        goto done;
                    }

                    // load memory
                    MemoryUtil.LoadMemory(pid, ref currentRate, ref maxSongLens);
                    diffRate = currentRate - lastRate;

                    if (diffRate == 0) //currentRate.Equals(lastRate)
                    {
                        Debug.Print($"Music pause? {currentRate} | {lastRate} | {(diffRate)}");
                        playerState = false;
                    }
                    else if (!playerState || !maxSongLens.Equals(lastLens))
                    {
                        var match = title.Replace(" - ", "\t").Split('\t');
                        if (match.Length > 1)
                        {
                            currentSong = match[0];
                            currentSing = match[1]; // like spotify
                        }
                        else
                        {
                            currentSong = title;
                            currentSing = string.Empty;
                        }

                        playerState = true;
                    }
                    // check
                    else if (Math.Abs(diffRate) < 0.0416)
                    {
                        // skip playing
                        Debug.Print($"Skip Rpc {Math.Abs(diffRate)}");
                        continue;
                    }

                done:
                    Debug.Print($"playerState -> {playerState} | Equals {maxSongLens} | {lastLens}");

                    // update
                    var isfullScreen = false;
                    var iswhiteLists = false;
                    var largeImgPlaying = Properties.Settings.Default.DefSkin;

                    // User setting
                    if (!settingFullscreen.Checked && Win32Api.User32.IsFullscreenAppRunning())
                        isfullScreen = true;
                    if (!settingWhitelists.Checked && Win32Api.User32.IsWhitelistAppRunning())
                        iswhiteLists = true;

                    var smallImgPlaying = "";
                    var smallImgText = "";
                    var timeNow = new Timestamps(DateTime.UtcNow.Subtract(TimeSpan.FromSeconds(currentRate)), DateTime.UtcNow.Subtract(TimeSpan.FromSeconds(currentRate)).Add(TimeSpan.FromSeconds(maxSongLens)));

                    if (playingStats.Checked && !playerState && actiButton.Checked && !isfullScreen && !iswhiteLists) // Show paused?
                    {
                        smallImgPlaying = "";
                        largeImgText = "Paused";
                        if (showSmallImg.Checked)
                        {
                            if (defDark.Checked)
                                smallImgPlaying = "dark_inactive";
                            if (defDiscord.Checked)
                                smallImgPlaying = "discord_inactive";
                            if (defNetease.Checked)
                                smallImgPlaying = "netease_inactive";
                            if (defWhite.Checked)
                                smallImgPlaying = "white_inactive";
                        }

                        smallImgText = "Paused";
                        timeNow = new Timestamps(); // Clear timestamps only
                        goto Update;
                    }
                    else if (!playerState || !actiButton.Checked || isfullScreen || iswhiteLists)
                    {
                        discord.ClearPresence();
                        Debug.Print("Clear Rpc");
                        continue;
                    }
                    else
                    {
                        timeNow = new Timestamps(DateTime.UtcNow.Subtract(TimeSpan.FromSeconds(currentRate)), DateTime.UtcNow.Subtract(TimeSpan.FromSeconds(currentRate)).Add(TimeSpan.FromSeconds(maxSongLens)));
                    }

                    if (playingStats.Checked)
                    {
                        smallImgPlaying = "";
                        largeImgText = "Netease Cloud Music";
                        if (showSmallImg.Checked)
                        {
                            if (defDark.Checked)
                                smallImgPlaying = "dark_active";
                            if (defDiscord.Checked)
                                smallImgPlaying = "discord_active";
                            if (defNetease.Checked)
                                smallImgPlaying = "netease_active";
                            if (defWhite.Checked)
                                smallImgPlaying = "white_active";
                            smallImgText = "Playing";
                        }
                    }

                Update:
                    discord.SetPresence(new RichPresence
                    {
                        Assets = new Assets
                        {
                            LargeImageKey = largeImgPlaying,
                            LargeImageText = largeImgText,
                            SmallImageKey = smallImgPlaying,
                            SmallImageText = smallImgText,
                        },
                        Timestamps = timeNow,
                        Details = currentSong,
                        State = $"By: {currentSing}",
                    });
                    Debug.Print("Update Rpc");
                }
                
            });

            Application.Run();
        }
    }
}
