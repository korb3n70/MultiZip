using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;

namespace EstrattoreZipAvanzato
{
    public class MainForm : Form
    {
        private TextBox txtSource;
        private TextBox txtDest;
        private TextBox txtLog;
        private ProgressBar progress;
        private CheckBox chkDelete, chkOverwrite, chkRecursive, chkKeepStructure;

        public MainForm()
        {
            Text = "Estrattore ZIP Avanzato";
            Size = new System.Drawing.Size(700, 500);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new System.Drawing.Font("Segoe UI", 10);

            // --- Labels & Inputs ---
            var lblSource = new Label { Text = "Cartella sorgente:", Location = new System.Drawing.Point(10, 20), AutoSize = true };
            Controls.Add(lblSource);

            txtSource = new TextBox { Location = new System.Drawing.Point(150, 18), Size = new System.Drawing.Size(350, 25) };
            Controls.Add(txtSource);

            var btnSource = new Button { Text = "Sfoglia", Location = new System.Drawing.Point(510, 17) };
            btnSource.Click += (s, e) =>
            {
                using var dialog = new FolderBrowserDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                    txtSource.Text = dialog.SelectedPath;
            };
            Controls.Add(btnSource);

            var lblDest = new Label { Text = "Cartella destinazione:", Location = new System.Drawing.Point(10, 60), AutoSize = true };
            Controls.Add(lblDest);

            txtDest = new TextBox { Location = new System.Drawing.Point(150, 58), Size = new System.Drawing.Size(350, 25) };
            Controls.Add(txtDest);

            var btnDest = new Button { Text = "Sfoglia", Location = new System.Drawing.Point(510, 57) };
            btnDest.Click += (s, e) =>
            {
                using var dialog = new FolderBrowserDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                    txtDest.Text = dialog.SelectedPath;
            };
            Controls.Add(btnDest);

            var btnApriDest = new Button { Text = "Apri destinazione", Location = new System.Drawing.Point(510, 90) };
            btnApriDest.Click += (s, e) =>
            {
                if (Directory.Exists(txtDest.Text))
                    System.Diagnostics.Process.Start("explorer.exe", txtDest.Text);
            };
            Controls.Add(btnApriDest);

            // --- Opzioni ---
            chkDelete = new CheckBox { Text = "Elimina ZIP dopo estrazione", Location = new System.Drawing.Point(10, 100), Size = new System.Drawing.Size(300, 25) };
            Controls.Add(chkDelete);

            chkOverwrite = new CheckBox { Text = "Sovrascrivi file esistenti", Location = new System.Drawing.Point(10, 130), Size = new System.Drawing.Size(300, 25) };
            Controls.Add(chkOverwrite);

            chkRecursive = new CheckBox { Text = "Cerca ZIP anche nelle sottocartelle", Location = new System.Drawing.Point(10, 160), Size = new System.Drawing.Size(350, 25) };
            Controls.Add(chkRecursive);

            chkKeepStructure = new CheckBox { Text = "Mantieni la struttura delle sottocartelle", Location = new System.Drawing.Point(10, 190), Size = new System.Drawing.Size(380, 25) };
            Controls.Add(chkKeepStructure);

            // --- Log ---
            var lblLog = new Label { Text = "Log operazioni:", Location = new System.Drawing.Point(10, 230), AutoSize = true };
            Controls.Add(lblLog);

            txtLog = new TextBox { Location = new System.Drawing.Point(10, 260), Size = new System.Drawing.Size(620, 160), Multiline = true, ScrollBars = ScrollBars.Vertical };
            Controls.Add(txtLog);

            // --- Progress Bar ---
            progress = new ProgressBar { Location = new System.Drawing.Point(10, 430), Size = new System.Drawing.Size(620, 20) };
            Controls.Add(progress);

            // --- Bottone Esegui ---
            var btnExtract = new Button
            {
                Text = "ESTRAI ZIP",
                Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(250, 455),
                Size = new System.Drawing.Size(150, 35)
            };
            btnExtract.Click += BtnExtract_Click;
            Controls.Add(btnExtract);
        }

        private void BtnExtract_Click(object sender, EventArgs e)
        {
            string source = txtSource.Text;
            string dest = txtDest.Text;

            txtLog.AppendText("=== Avvio estrazione ===\r\n");

            if (string.IsNullOrWhiteSpace(source) || !Directory.Exists(source))
            {
                MessageBox.Show("La cartella sorgente non esiste o è vuota.", "Errore");
                return;
            }

            if (string.IsNullOrWhiteSpace(dest))
            {
                MessageBox.Show("La cartella destinazione non è valida.", "Errore");
                return;
            }

            Directory.CreateDirectory(dest);

            var zipFiles = Directory.EnumerateFiles(source, "*.zip",
                chkRecursive.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly).ToList();

            if (!zipFiles.Any())
            {
                MessageBox.Show("Nessun file ZIP trovato.", "Info");
                return;
            }

            progress.Minimum = 0;
            progress.Maximum = zipFiles.Count;
            progress.Value = 0;

            foreach (var zip in zipFiles)
            {
                txtLog.AppendText($"Apro ZIP: {zip}\r\n");
                Application.DoEvents();

                string extractPath = chkKeepStructure.Checked
                    ? Path.Combine(dest, Path.GetRelativePath(source, Path.GetDirectoryName(zip)!))
                    : dest;

                Directory.CreateDirectory(extractPath);

                try
                {
                    ZipFile.ExtractToDirectory(zip, extractPath, chkOverwrite.Checked);
                    txtLog.AppendText($"Estratto: {Path.GetFileName(zip)}\r\n");

                    if (chkDelete.Checked)
                    {
                        File.Delete(zip);
                        txtLog.AppendText("ZIP eliminato.\r\n");
                    }
                }
                catch (Exception ex)
                {
                    txtLog.AppendText($"ERRORE durante l'estrazione: {ex.Message}\r\n");
                }

                progress.Value++;
                Application.DoEvents();
            }

            progress.Value = progress.Maximum;
            MessageBox.Show("Estrazione completata!", "Fatto");
            txtLog.AppendText("=== Completato ===\r\n");
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new MainForm());
        }
    }
}
