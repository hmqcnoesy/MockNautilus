using Microsoft.Win32;
using MockNautilus.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MockNautilus
{
    public partial class MockNautilusForm : Form
    {
        private List<string> _dllsToSkip = new List<string> { "adodb.dll" };
        private string _registryKey = "MockNautilus";


        public MockNautilusForm()
        {
            InitializeComponent();
            LoadLibraries();
            LoadTypes();
            LoadSavedParameters();
        }


        private void LoadLibraries()
        {
            var directory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var libraries = Directory.EnumerateFiles(directory, "*.dll");
            foreach(var lib in libraries)
            {
                var fileName = new FileInfo(lib).Name;
                if (_dllsToSkip.Contains(fileName)) continue;

                cmbLibrary.Items.Add(fileName);
            }
        }


        private void LoadTypes()
        {
            cmbType.Items.Clear();
            if (string.IsNullOrEmpty(cmbLibrary.Text)) return;

            var directory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var library = Path.Combine(directory, cmbLibrary.Text);

            var assembly = Assembly.LoadFrom(library);
            var allTypes = assembly.GetTypes();
            foreach(var type in allTypes)
            {
                if (!type.IsInterface && type.GetInterfaces().Select(i => i.Name).Contains("IEntityExtension"))
                    cmbType.Items.Add(type);
            }
        }


        private void Loaded(object sender, EventArgs e)
        {
            txtPassword.Select();
        }


        private void ClickedExecute(object sender, EventArgs e)
        {
            var type = cmbType.SelectedItem as Type;
            if (type == null) return;

            var instance = (LSEXT.IEntityExtension)Activator.CreateInstance(type);
            if (instance == null)
            {
                MessageBox.Show("Couldn't create instance of selected type: " + cmbType.Text);
                return;
            }

            var parameters = SetupMockParameters();
            if (parameters != null) ((LSEXT.IEntityExtension)instance).Execute(parameters);
        }


        private LSEXT.LSExtensionParameters SetupMockParameters()
        {
            var parameters = new MockExtensionParameters();
            try
            {
                parameters.Product = txtProduct.Text;
                parameters.Version = txtVersion.Text;
                parameters.Server = txtServer.Text;
                parameters.ComputerName = txtComputerName.Text;
                parameters.HKeyCurrentUser = txtHKeyCurrentUser.Text;
                parameters.HKeyLocalMachine = txtHKeyLocalMachine.Text;
                parameters.OperatorName = txtOperatorName.Text;
                parameters.RoleName = txtRoleName.Text;
                parameters.OperatorId = decimal.Parse(txtOperatorId.Text);
                parameters.RoleId = decimal.Parse(txtRoleId.Text);
                parameters.WorkstationId = decimal.Parse(txtWorkstationId.Text);
                parameters.SessionId = decimal.Parse(txtSessionId.Text);
                parameters.Username = txtUsername.Text;
                parameters.Password = txtPassword.Text;
                parameters.PasswordRole = txtPasswordRole.Text;
                parameters.ServerInfo = txtServerInfo.Text;

                parameters.EntityId = decimal.Parse(txtEntityId.Text);
                parameters.EntityRecords = SplitAndParseEntityRecordsText();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Couldn't load given parameters: " + ex.Message);
                return null;
            }

            return parameters;
        }


        private ADODB.Recordset SplitAndParseEntityRecordsText()
        {
            var parts = txtEntityRecords.Text.Split(";, \t\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            var entityRecordIds = new ADODB.Recordset();
            entityRecordIds.Fields.Append("ID", ADODB.DataTypeEnum.adVariant);
            entityRecordIds.Open();

            foreach (var id in parts)
            {
                entityRecordIds.AddNew("ID", id);
            }

            if (parts.Length > 0) entityRecordIds.MoveFirst();

            return entityRecordIds;
        }


        private void SaveParameters()
        {
            WriteSetting("Library", cmbLibrary.Text);
            WriteSetting("Type", cmbType.Text);
            WriteSetting("Product", txtProduct.Text);
            WriteSetting("Version", txtVersion.Text);
            WriteSetting("Server", txtServer.Text);
            WriteSetting("ComputerName", txtComputerName.Text);
            WriteSetting("HKeyCurrentUser", txtHKeyCurrentUser.Text);
            WriteSetting("HKeyLocalMachine", txtHKeyLocalMachine.Text);
            WriteSetting("OperatorName", txtOperatorName.Text);
            WriteSetting("RoleName", txtRoleName.Text);
            WriteSetting("OperatorId", txtOperatorId.Text);
            WriteSetting("RoleId", txtRoleId.Text);
            WriteSetting("WorkstationId", txtWorkstationId.Text);
            WriteSetting("SessionId", txtSessionId.Text);
            WriteSetting("Username", txtUsername.Text);
            WriteSetting("PasswordRole", System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(txtPasswordRole.Text)));
            WriteSetting("ServerInfo", txtServerInfo.Text);
            WriteSetting("EntityId", txtEntityId.Text);
            WriteSetting("EntityRecords", txtEntityRecords.Text);
        }


        private void WriteSetting(string settingName, string settingValue)
        {
            var key = Registry.CurrentUser.OpenSubKey("Software", true);
            key.CreateSubKey(_registryKey);
            key = key.OpenSubKey(_registryKey, true);
            key.SetValue(settingName, settingValue);
        }


        private void LoadSavedParameters()
        {
            var isFirstRun = !Registry.CurrentUser.OpenSubKey("Software", false).GetSubKeyNames().Contains(_registryKey);
            if (isFirstRun)
            {
                LoadDefaultParameters();
                return;
            }

            cmbLibrary.Text = ReadSetting("Library");
            cmbType.Text = ReadSetting("Type");
            txtProduct.Text = ReadSetting("Product");
            txtVersion.Text = ReadSetting("Version");
            txtServer.Text = ReadSetting("Server");
            txtComputerName.Text = ReadSetting("ComputerName");
            txtHKeyCurrentUser.Text = ReadSetting("HKeyCurrentUser");
            txtHKeyLocalMachine.Text = ReadSetting("HKeyLocalMachine");
            txtOperatorName.Text = ReadSetting("OperatorName");
            txtRoleName.Text = ReadSetting("RoleName");
            txtOperatorId.Text = ReadSetting("OperatorId");
            txtRoleId.Text = ReadSetting("RoleId");
            txtWorkstationId.Text = ReadSetting("WorkstationId");
            txtSessionId.Text = ReadSetting("SessionId");
            txtUsername.Text = ReadSetting("Username");
            txtPasswordRole.Text = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(ReadSetting("PasswordRole")));
            txtServerInfo.Text = ReadSetting("ServerInfo");
            txtEntityId.Text = ReadSetting("EntityId");
            txtEntityRecords.Text = ReadSetting("EntityRecords");
        }


        private void LoadDefaultParameters()
        {
            txtProduct.Text = "Nautilus";
            txtVersion.Text = "9.3";
            txtComputerName.Text = System.Environment.MachineName;
            txtHKeyCurrentUser.Text = @"HKEY_CURRENT_USER\Software\Thermo\Nautilus\9.3";
            txtHKeyLocalMachine.Text = @"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Thermo\Nautilus\9.3";
            txtRoleName.Text = "System";
            txtRoleId.Text = "1";
            txtUsername.Text = System.Environment.UserName;
        }


        private string ReadSetting(string settingName)
        {
            var key = Registry.CurrentUser.OpenSubKey("Software", true);
            key.CreateSubKey(_registryKey);
            key = key.OpenSubKey(_registryKey, false);
            return (string)key.GetValue(settingName);
        }


        private void ClosingForm(object sender, FormClosingEventArgs e)
        {
            try
            {
                SaveParameters();
            }
            catch
            {
                MessageBox.Show("Check that all numeric parameters have numeric values.");
                e.Cancel = true;
            }
        }

        private void cmbLibrary_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadTypes();
        }
    }
}
