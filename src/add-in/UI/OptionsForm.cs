using System;
using System.Drawing;
using System.Windows.Forms;
using MyJournal.Notebook.Diagnostics;

namespace MyJournal.Notebook.UI
{
    public partial class OptionsForm : Form
    {
        #region Form setter properties

        internal string AboutIconImage_ResourceName
        {
            set
            {
                var assembly = GetType().Assembly;
                using (var stream = assembly.GetManifestResourceStream(value))
                {
                    using (var icon = new Icon(stream, pictureBox1.Size))
                    {
                        pictureBox1.Image = icon.ToBitmap();
                    }
                }
            }
        }

        internal string AboutText { set { textBox1.Text = value; } }

        internal string Title { set { Text = value; } }

        #endregion

        public OptionsForm()
        {
            InitializeComponent();
        }

        private void Form_Closing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (_settingsChangedCount > 0)
                {
                    var info = $"Saving user.config: {Component.UserConfigPath}";
                    Tracer.WriteInfoLine(info);
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                Utils.ExceptionHandler.HandleException(ex);
            }
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) Close();
        }

        private void Form_Load(object sender, EventArgs e)
        {
            Tracer.WriteTraceMethodLine();
            if (Properties.Settings.Default.UpgradeSettings)
            {
                var info = $"Upgrading user.config: {Component.UserConfigPath}";
                Tracer.WriteInfoLine(info);
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpgradeSettings = false;
                _settingsChangedCount++;
            }

            PageTemplates.Items.AddRange(Templates.TemplateFactory.Items);
            PageTemplates.SelectedIndex = PageTemplates.Items.IndexOf(
              Properties.Settings.Default.PageTemplate);
            PageTemplates.SelectedIndexChanged +=
              PageTemplates_SelectedIndexChanged;

            PageSize.Items.AddRange(Templates.PageSize.Items);
            PageSize.SelectedIndex = PageSize.Items.IndexOf(
              Properties.Settings.Default.PaperSize);
            PageSize.SelectedIndexChanged += PageSize_SelectedIndexChanged;

            StorageAccount.Items.AddRange(new Config.StorageAccount().Items);
            StorageAccount.SelectedIndex = StorageAccount.Items.IndexOf(
              Properties.Settings.Default.StorageAccount);
            StorageAccount.SelectedIndexChanged += StorageAccount_SelectedIndexChanged;
        }

        private void PageSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PaperSize =
              PageSize.SelectedItem as string;

            _settingsChangedCount++;
            Tracer.WriteTraceMethodLine("Paper Size = {0}",
              Properties.Settings.Default.PaperSize);
        }

        private void PageTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.PageTemplate =
              PageTemplates.SelectedItem as string;

            _settingsChangedCount++;
            Tracer.WriteTraceMethodLine("PageTemplate = {0}",
              Properties.Settings.Default.PageTemplate);
        }

        private void StorageAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.StorageAccount =
              StorageAccount.SelectedItem as string;

            _settingsChangedCount++;
            Tracer.WriteTraceMethodLine("StorageAccount = {0}",
              Properties.Settings.Default.StorageAccount);
        }

        private int _settingsChangedCount;
    }
}
