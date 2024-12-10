using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace _1C8_Base_Manager
{
    public partial class Form1 : Form
    {
        private string currentUserPath1;
        private string currentUserPath2;

        public Form1()
        {
            InitializeComponent();
            InitializeForm();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void InitializeForm()
        {
            // Setup DataGridView
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = true;
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView2.MultiSelect = true;

            // Setup ComboBoxes with users
            PopulateUserList();

            // Event Handlers
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;
            comboBox2.SelectedIndexChanged += ComboBox2_SelectedIndexChanged;
            button1.Click += BtnCopyRight_Click;
            button2.Click += BtnCopyLeft_Click;
        }

        private void PopulateUserList()
        {
            string usersDirectory = @"C:\Users";
            string[] userDirectories = Directory.GetDirectories(usersDirectory);

            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            foreach (string userDir in userDirectories)
            {
                string userName = new DirectoryInfo(userDir).Name;
                comboBox1.Items.Add(userName);
                comboBox2.Items.Add(userName);
            }

            label1.Text = "Data loaded";
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedUser = comboBox1.SelectedItem.ToString();
            currentUserPath1 = $@"C:\Users\{selectedUser}\AppData\Roaming\1C\1CEStart\ibases.v8i";
            LoadConnectionsFromFile(currentUserPath1, dataGridView1);
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedUser = comboBox2.SelectedItem.ToString();
            currentUserPath2 = $@"C:\Users\{selectedUser}\AppData\Roaming\1C\1CEStart\ibases.v8i";
            LoadConnectionsFromFile(currentUserPath2, dataGridView2);
        }

        private void LoadConnectionsFromFile(string filePath, DataGridView gridView)
        {
            gridView.Rows.Clear();
            gridView.Columns.Clear();

            // Add columns
            gridView.Columns.Add("Name", "Name");
            gridView.Columns.Add("Server", "Server");
            gridView.Columns.Add("Reference", "Reference");
            gridView.Columns.Add("ID", "ID");

            if (!File.Exists(filePath))
            {
                label1.Text = "File not found.";
                return;
            }

            string fileContent = File.ReadAllText(filePath);
            MatchCollection matches = Regex.Matches(fileContent, @"\[([^\]]+)\]([^[]+)(?=\[|$)", RegexOptions.Singleline);

            foreach (Match match in matches)
            {
                string name = match.Groups[1].Value.Trim();
                string details = match.Groups[2].Value;

                string server = ExtractValue(details, @"Srvr=""([^""]+)""");
                string reference = ExtractValue(details, @"Ref=""([^""]+)""");
                string id = ExtractValue(details, @"ID=([^\r\n]+)");

                gridView.Rows.Add(name, server, reference, id);
            }
        }

        private string ExtractValue(string text, string pattern)
        {
            Match match = Regex.Match(text, pattern);
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        private void BtnCopyRight_Click(object sender, EventArgs e)
        {
            CopySelectedRows(dataGridView1, dataGridView2, currentUserPath2);
        }

        private void BtnCopyLeft_Click(object sender, EventArgs e)
        {
            CopySelectedRows(dataGridView2, dataGridView1, currentUserPath1);
        }
        private void CopySelectedRows(DataGridView sourceGrid, DataGridView destGrid, string destFilePath)
        {
            if (sourceGrid.SelectedRows.Count == 0) return;

            try
            {
                using (StreamWriter writer = File.AppendText(destFilePath))
                {
                    foreach (DataGridViewRow row in sourceGrid.SelectedRows)
                    {
                        string connectionBlock = GenerateConnectionBlock(row);
                        writer.WriteLine(connectionBlock);
                    }
                }

                // Reload the destination file
                LoadConnectionsFromFile(destFilePath, destGrid);
                label1.Text = "Connections copied successfully!" ;
            }
            catch (Exception ex)
            {
                label1.Text = $"Error copying connections: {ex.Message}";
            }
        }

        private string GenerateConnectionBlock(DataGridViewRow row)
        {
            string name = row.Cells[0].Value?.ToString() ?? "";
            string server = row.Cells[1].Value?.ToString() ?? "";
            string reference = row.Cells[2].Value?.ToString() ?? "";
            string id = row.Cells[3].Value?.ToString() ?? "";

            return $@"[{name}]
Connect=Srvr=""{server}"";Ref=""{reference}"";
ID={id}
OrderInList=16384
Folder=/
OrderInTree=256
External=0
ClientConnectionSpeed=Normal
App=Auto
WA=1
Version=8.3";
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string url = "https://github.com/e-gaydarzhi";
            Process.Start(url);
        }
    }
}
