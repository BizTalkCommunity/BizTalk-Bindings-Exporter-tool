using System;
using System.Configuration;
using System.Windows.Forms;

namespace BizTalk_Exporter
{
    public partial class connectionForm : Form
    {
        public connectionForm()
        {
            InitializeComponent();
        }

        private void confirmBtn_Click(object sender, EventArgs e)
        {
            var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var settings = configFile.AppSettings.Settings;
            settings["connString"].Value = connTxt.Text;

            configFile.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            Close();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void connectionForm_Load(object sender, EventArgs e)
        {
            var configs = ConfigurationManager.AppSettings;
            connTxt.Text = configs["connString"];
        }
    }
}

