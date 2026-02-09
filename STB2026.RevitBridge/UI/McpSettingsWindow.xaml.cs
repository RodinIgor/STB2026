using System.Windows;
using STB2026.RevitBridge.Infrastructure;

namespace STB2026.RevitBridge.UI
{
    /// <summary>
    /// Окно настроек MCP-подключения.
    /// Аналог Nonica AI Connector Settings (Auto/Manual, тоглы, статусы).
    /// </summary>
    public partial class McpSettingsWindow : Window
    {
        private readonly McpSettingsManager _settings;
        private readonly bool _pipeActive;

        public McpSettingsWindow(McpSettingsManager settings, bool pipeActive)
        {
            InitializeComponent();
            _settings = settings;
            _pipeActive = pipeActive;

            // Загрузить текущие настройки в UI
            tglConnection.IsChecked = _settings.EnableConnection;
            tglEditTools.IsChecked = _settings.EnableEditTools;

            // Заполнить Manual JSON
            txtManualJson.Text = _settings.GetManualConfigJson();
            txtConfigPath.Text = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                "Claude", "claude_desktop_config.json");

            RefreshStatuses();
        }

        /// <summary>Обновить все статусы AI-приложений.</summary>
        private void RefreshStatuses()
        {
            // ═══ Claude Desktop ═══
            var claude = _settings.DetectClaudeDesktop();
            if (claude.Configured)
            {
                txtClaudeIcon.Text = "✓ ";
                txtClaudeIcon.Style = (Style)FindResource("StatusOk");
                txtClaudeStatus.Text = claude.NeedsRestart
                    ? "Claude App has been set up and needs to be restarted."
                    : "Claude Desktop App подключён.";
                btnSetupClaude.Visibility = Visibility.Collapsed;
            }
            else if (claude.Installed)
            {
                txtClaudeIcon.Text = "○ ";
                txtClaudeIcon.Style = (Style)FindResource("StatusWarn");
                txtClaudeStatus.Text = "Claude Desktop App найден, требуется настройка.";
                btnSetupClaude.Visibility = Visibility.Visible;
            }
            else
            {
                txtClaudeIcon.Text = "✗ ";
                txtClaudeIcon.Style = (Style)FindResource("StatusFail");
                txtClaudeStatus.Text = "Claude Desktop App не обнаружен.";
                btnSetupClaude.Visibility = Visibility.Collapsed;
            }

            // ═══ Cursor ═══
            var cursor = _settings.DetectCursor();
            if (cursor.Configured)
            {
                txtCursorIcon.Text = "✓ ";
                txtCursorIcon.Style = (Style)FindResource("StatusOk");
                txtCursorStatus.Text = "Cursor подключён.";
                btnSetupCursor.Visibility = Visibility.Collapsed;
            }
            else if (cursor.Installed)
            {
                txtCursorIcon.Text = "○ ";
                txtCursorIcon.Style = (Style)FindResource("StatusWarn");
                txtCursorStatus.Text = "Cursor найден, требуется настройка.";
                btnSetupCursor.Visibility = Visibility.Visible;
            }
            else
            {
                txtCursorIcon.Text = "✗ ";
                txtCursorIcon.Style = (Style)FindResource("StatusFail");
                txtCursorStatus.Text = "Cursor не обнаружен.";
                btnSetupCursor.Visibility = Visibility.Collapsed;
            }

            // ═══ Claude Code ═══
            var code = _settings.DetectClaudeCode();
            if (code.Installed)
            {
                txtCodeIcon.Text = "✓ ";
                txtCodeIcon.Style = (Style)FindResource("StatusOk");
                txtCodeStatus.Text = "Multi Agent Claude Code is ready to be used.";
            }
            else
            {
                txtCodeIcon.Text = "✗ ";
                txtCodeIcon.Style = (Style)FindResource("StatusFail");
                txtCodeStatus.Text = "Claude Code не обнаружен.";
            }

            // ═══ Pipe статус ═══
            if (_pipeActive)
            {
                txtPipeIcon.Text = "⚡";
                txtPipeStatus.Text = $"Named Pipe: активен (PID {System.Diagnostics.Process.GetCurrentProcess().Id})";
                txtPipeStatus.Foreground = System.Windows.Media.Brushes.Green;
            }
            else
            {
                txtPipeIcon.Text = "⏳";
                txtPipeStatus.Text = "Named Pipe: ожидание запуска...";
            }
        }

        // ═══ Кнопки Auto-setup ═══

        private void BtnSetupClaude_Click(object sender, RoutedEventArgs e)
        {
            var result = _settings.AutoSetupClaude();
            if (result.Success && result.NeedsRestart)
                pnlRestart.Visibility = Visibility.Visible;

            MessageBox.Show(result.Message, "STB2026",
                MessageBoxButton.OK,
                result.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);

            RefreshStatuses();
        }

        private void BtnSetupCursor_Click(object sender, RoutedEventArgs e)
        {
            var result = _settings.AutoSetupCursor();
            if (result.Success && result.NeedsRestart)
                pnlRestart.Visibility = Visibility.Visible;

            MessageBox.Show(result.Message, "STB2026",
                MessageBoxButton.OK,
                result.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);

            RefreshStatuses();
        }

        // ═══ Кнопка Copy JSON ═══

        private void BtnCopyJson_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(txtManualJson.Text);
            MessageBox.Show("Скопировано в буфер обмена!", "STB2026",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // ═══ Тоглы ═══

        private void TglConnection_Changed(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;
            _settings.EnableConnection = tglConnection.IsChecked == true;
            _settings.SaveSettings();
        }

        private void TglEditTools_Changed(object sender, RoutedEventArgs e)
        {
            if (_settings == null) return;
            _settings.EnableEditTools = tglEditTools.IsChecked == true;
            _settings.SaveSettings();
        }

        // ═══ Close ═══

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
