/*
 * Applica, Inc.
 * By José O. Lara
 * 2023.02.28
 */

namespace wpfApp_Demo {
    
    
    
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;
    using System;
    using System.Configuration;
    using System.Diagnostics;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Input;
    using wpfApp_Demo.src.core.models;
    using wpfApp_Demo.src.core.viewModels;
    using wpfApp_Demo.src.ui.views;

    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application {
        private readonly IHost _host;
        public static IServiceProvider? ServiceProvider { get; private set; }

        private MainWpf? _MainWpf { get; set; }

        public App() {
            _host = Host.CreateDefaultBuilder()  // Use default settings
                                                 //new HostBuilder()          // Initialize an empty HostBuilder
              .ConfigureAppConfiguration((context, builder) => {
                  // Add other configuration files...
                  builder.AddJsonFile("appsettings.json", optional: true);

              }).ConfigureServices((context, services) => {
                  ConfigureServices(context.Configuration, services);
              })
              .ConfigureLogging(logging => {
                  // Add other loggers...
              })
            .Build();

            ServiceProvider = _host.Services;
        }

        private void ConfigureServices(IConfiguration configuration, IServiceCollection services) {
            services.Configure<AppSettings>(configuration.GetSection(nameof(AppSettings)));

            // Register all ViewModels.
            services.AddSingleton<MainVm>();

            //Read appsettings.json
            //MsSqlCnnStr Server
            string? myCnnString = configuration["AppSettings:MsSqlCnnStr"];

            // Register all the Windows of the applications.
            services.AddTransient<MainWpf>();

        }

        protected override async void OnStartup(StartupEventArgs e) {
            await _host.StartAsync();

            string miProcessName = Process.GetCurrentProcess().ProcessName;

            Process[] miProcess = Process.GetProcessesByName(miProcessName);

            if (miProcess.Length > 1) {
                MessageBox.Show($"{miProcessName} already running", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                Process.GetCurrentProcess().Kill();
            } else {
                _MainWpf = ServiceProvider!.GetRequiredService<MainWpf>();

                _MainWpf.Closing += OnClosing!;

                _MainWpf.MouseDown += Window_MouseDown;

                _MainWpf.PreviewKeyDown += Window_PreviewKeyDown;

                _MainWpf.Show();
            }

            base.OnStartup(e);
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e) {
            try {
                if (e.Key == Key.Escape) {
                    _MainWpf!.Close();
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.ToString(), ex.Message, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) {
            try {
                if (e.ChangedButton == MouseButton.Left) {
                    _MainWpf!.DragMove();
                }
            } catch (Exception) {
                // ignored
            }
        }

        protected override async void OnExit(ExitEventArgs e) {
            using (_host) {
                await _host.StopAsync(TimeSpan.FromSeconds(5));
            }

            base.OnExit(e);
        }

        private static void OnClosing(object sender, System.ComponentModel.CancelEventArgs e) {
            e.Cancel = OnSessionEnding();
        }

        private static bool OnSessionEnding() {

            MessageBoxResult response = MessageBox.Show("!!!Do you really want to exit?", "Exiting...", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);

            return response != MessageBoxResult.Yes;
        }

    }
}
