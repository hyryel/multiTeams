using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;

using MMT.Core;

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;
using System.Windows.Forms;

namespace MMT.UI
{
    public partial class MainWindow : MetroWindow
    {
        private readonly ProfileManager _profileManager;
        private readonly TeamsLauncher _teamsLauncher;
        private readonly RegistryManager _registryManager;

        public NotifyIcon TrayIcon { get; set; }
        #region Tray Icon initialization
        private void InitializaTrayIcon()
        {
            TrayIcon = new NotifyIcon()
            {
                Icon = Resource.Taskbar,
                Visible = true,
                BalloonTipTitle = MMT.Core.StaticResources.AppName,
                Text = StaticResources.AppName,
            };

            //context tray icon initialization
            TrayIcon.ContextMenuStrip = new ContextMenuStrip();
            //exit button
            new ToolStripButton("Exit");
            var trayExitButton = TrayIcon.ContextMenuStrip.Items.Add("Exit");
            trayExitButton.Click += TrayExitButton_Click;
            var trayRestaureButton = TrayIcon.ContextMenuStrip.Items.Add("Restore");
            trayRestaureButton.Click += TrayIcon_DoubleClick;
            TrayIcon.DoubleClick += TrayIcon_DoubleClick;
        }

        private void TrayExitButton_Click(object sender, EventArgs e)
        {
            App.Current.Shutdown();
        }

        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            this.WindowState = WindowState.Normal;
            this.ShowInTaskbar = true;
        }
        #endregion

        public MainWindow()
        {
            InitializeComponent();
            InitializaTrayIcon();

            _profileManager = new ProfileManager();
            _teamsLauncher = new TeamsLauncher();
            _registryManager = new RegistryManager();
            DataContext = _profileManager;
            App.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            ChangeTabVisibility();
            AutoStartCheck();
        }


        private void ChangeTabVisibility()
        {
            if (tbiNewProfile.Visibility == Visibility.Visible)
            {
                tbiProfiles.Visibility = Visibility.Visible;
                tbiNewProfile.Visibility = Visibility.Collapsed;
                tbcMain.SelectedItem = tbiProfiles;
            }
            else
            {
                tbiProfiles.Visibility = Visibility.Collapsed;
                tbiNewProfile.Visibility = Visibility.Visible;
                tbcMain.SelectedItem = tbiNewProfile;
            }
        }
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.DataContext is Profile profile)
            {
                _teamsLauncher.Start(profile);
            }
        }

        private void TrayMouseDoubleClick(object sender, RoutedEventArgs e)
        {
            Show();
            WindowState = WindowState.Normal;
        }

        private void AutoStartCheck()
        {
            chkAutoStart.IsChecked = _registryManager.IsApplicationInStartup(StaticResources.AppName);

            if (chkAutoStart.IsChecked.HasValue && chkAutoStart.IsChecked.Value)
            {
                Show();
                WindowState = WindowState.Minimized;

                Thread thread = new(() =>
                {
                    foreach (Profile item in lstProfiles.Items.OfType<Profile>())
                    {
                        if (!item.IsDisabled)
                        {
                            _teamsLauncher.Start(item);
                        }
                    }
                });
                thread.Start();
            }
        }

        private void ChkAutoStart_Click(object sender, RoutedEventArgs e)
        {
            if (chkAutoStart.IsChecked.HasValue && chkAutoStart.IsChecked.Value)
            {
                _registryManager.AddApplicationInStartup(StaticResources.AppName);
            }
            else if (_registryManager.IsApplicationInStartup(StaticResources.AppName))
            {
                _registryManager.RemoveApplicationFromStartup(StaticResources.AppName);
            }
        }

        private void BtnNewProfile_Click(object sender, RoutedEventArgs e)
        {
            txtProfileName.Clear();
            ChangeTabVisibility();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _profileManager.Save(txtProfileName.Text);
                ChangeTabVisibility();
            }
            catch (Exception ex)
            {
                MessageHelper.Info(ex.Message);
                txtProfileName.Focus();
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            ChangeTabVisibility();
        }

        private void BtnLaunchTeams_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (lstProfiles.SelectedItems?.Count > 0)
                {
                    lstProfiles.SelectedItems.OfType<Profile>()
                        .Where((item) => !item.IsDisabled)
                        .ToList()
                        .ForEach((item) =>
                    {
                        _teamsLauncher.Start(item);
                    });
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Info(ex.Message);
            }
        }

        private async void LstProfiles_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                for (int i = lstProfiles.SelectedItems.Count - 1; i >= 0; i--)
                {
                    if (lstProfiles.SelectedItems[i] is Profile selectedProfile && !selectedProfile.IsDefault &&
                        await MessageHelper.Confirm($"Delete profile?\nProfile name: {selectedProfile.Name}") == MessageDialogResult.Affirmative)
                    {
                        try
                        {
                            _profileManager.Delete(selectedProfile);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            MessageHelper.Info($"Profile {selectedProfile.Name} has not been deleted. Close Microsoft Teams and try again.");
                        }
                        catch (IOException ex)
                        {
                            if (await MessageHelper.Confirm($"{ex.Message} Do you want continue?") == MessageDialogResult.Affirmative)
                            {
                                var controller = await MessageHelper.Wait("Processing, please wait.");
                                _ = Task.Run(() =>
                                  {
                                      _teamsLauncher.CloseAllInstances();
                                      _profileManager.Delete(selectedProfile);
                                  }).ContinueWith(a => controller.CloseAsync());
                            }
                        }
                    }
                }
            }
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        }

        private async void LstProfiles_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is ListBoxItem item && item.DataContext is Profile selectedProfile)
            {
                if (selectedProfile.IsDisabled)
                {
                    _profileManager.Enable(selectedProfile);
                }
                else if (await MessageHelper.Confirm($"Disable profile?\nProfile name: {selectedProfile.Name}") == MessageDialogResult.Affirmative)
                {
                    _profileManager.Disable(selectedProfile);
                }
            }
        }

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.WindowState = WindowState.Minimized;
            this.ShowInTaskbar = false;
            TrayIcon.BalloonTipIcon = ToolTipIcon.Info;
            TrayIcon.BalloonTipTitle = "See... ?";
            TrayIcon.BalloonTipText = "Multi-Teams is in system tray right now !";
            TrayIcon.ShowBalloonTip(2000);
        }
    }
}

