using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;
using Net.Codecrete.QrCodeGenerator;

namespace Bulk_vCard_QR_Generator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void chooseContactList_Click(object sender, EventArgs e)
        {
            isFinished.Visible = false;
            fileChooser();
            contactPathBox.Text = programState.contacFile;
        }

        private void chooseFolder_Click(object sender, EventArgs e)
        {
            isFinished.Visible = false;
            folderChooser();
            outPutFolderBox.Text = programState.outputPath;
        }

        private void contactPathBox_TextChanged(object sender, EventArgs e)
        {
            isFinished.Visible = false;
            if (!File.Exists(contactPathBox.Text))
                contactPathBox.BackColor = Color.LightCoral;
            else
            {
                contactPathBox.BackColor = SystemColors.Window;
                programState.contacFile = contactPathBox.Text;
            }
        }

        private void outPutFolderBox_TextChanged(object sender, EventArgs e)
        {
            isFinished.Visible = false;
            if (!Directory.Exists(outPutFolderBox.Text))
                outPutFolderBox.BackColor = Color.LightCoral;
            else
            {
                outPutFolderBox.BackColor = SystemColors.Window;
                programState.outputPath = outPutFolderBox.Text;
            }
        }

        private void instructionButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Excel formatý þu þekilde olmalý: ADSOYAD, TELEFON, EMAIL, WEB, ÞÝRKET, ÜNVAN.\n" +
                "Ýsimler tek hücrede tam haliyle yer almalý. QR kodlar UTF-8 desteklidir.",
                "Kullaným Bilgisi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void generateQR_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(programState.outputPath) || string.IsNullOrWhiteSpace(programState.contacFile))
            {
                MessageBox.Show("Lütfen kiþi listesi ve çýktý klasörü seçiniz.", "Uyarý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(programState.contacFile))
            {
                MessageBox.Show("Kiþi dosyasý bulunamadý.", "Uyarý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(programState.outputPath))
            {
                MessageBox.Show("Çýktý klasörü bulunamadý.", "Uyarý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int createdCount = readExcelFile(programState.contacFile);

            isFinished.ForeColor = Color.FromArgb(0, 192, 0);
            isFinished.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold);
            isFinished.Text = $"{createdCount} QR kod baþarýyla oluþturuldu!";
            isFinished.Visible = true;
        }

        private int readExcelFile(string filePath)
        {
            using var workbook = new XLWorkbook(new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed().Skip(1); // Ýlk satýr baþlýk

            int count = 0;

            foreach (var row in rows)
            {
                string fullName = row.Cell(1).GetValue<string>().Trim();
                string tel = row.Cell(2).GetValue<string>().Trim();
                string email = row.Cell(3).GetValue<string>().Trim();
                string web = row.Cell(4).GetValue<string>().Trim();
                string company = row.Cell(5).GetValue<string>().Trim();
                string position = row.Cell(6).GetValue<string>().Trim();

                if (string.IsNullOrWhiteSpace(fullName)) continue;

                string vCard =
                    "BEGIN:VCARD\n" +
                    "VERSION:3.0\n" +
                    $"N;CHARSET=UTF-8:{fullName};;;;\n" +    // sadece tek satýr isim
                    $"FN;CHARSET=UTF-8:{fullName}\n" +
                    $"ORG;CHARSET=UTF-8:{company}\n" +
                    $"TITLE;CHARSET=UTF-8:{position}\n" +
                    $"TEL;TYPE=WORK,VOICE:{tel}\n" +
                    $"EMAIL;CHARSET=UTF-8;TYPE=INTERNET:{email}\n" +
                    $"URL;CHARSET=UTF-8:{web}\n" +
                    $"REV:{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}\n" +
                    "END:VCARD";

                var qr = QrCode.EncodeText(vCard, QrCode.Ecc.Medium, true);

                string safeFileName = CleanFileName(fullName);
                string fullPath = Path.Combine(programState.outputPath, safeFileName + ".png");

                qr.SaveAsPng(fullPath, 10, programState.borderSize, programState.foreGroundColor, programState.backGroundColor);

                count++;

                isFinished.Text = $"{count} QR oluþturuldu...";
                isFinished.Visible = true;
                isFinished.Refresh();
            }

            return count;
        }

        private string CleanFileName(string input)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                input = input.Replace(c.ToString(), "_");
            return input;
        }

        private void fileChooser()
        {
            using OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Excel Dosyasý Seçiniz",
                Filter = "Excel Dosyalarý|*.xls;*.xlsx"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
                programState.contacFile = ofd.FileName;
        }

        private void folderChooser()
        {
            using FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                programState.outputPath = fbd.SelectedPath;
        }
    }

    static class programState
    {
        public static string outputPath = "";
        public static string contacFile = "";
        public static Color foreGroundColor = Color.Black;
        public static Color backGroundColor = Color.White;
        public static int borderSize = 10;
    }
}
