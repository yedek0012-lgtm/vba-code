// =====================================================================
//  mainform.OtomatikKesit.cs  -  "Otomatik Kesit" butonu
// ---------------------------------------------------------------------
//  Kurulum:
//   1) Bu dosyayı ve AutoSectionDetector.cs'i OutlineDrawing projesine ekleyin.
//   2) mainform tasarımcısında btnHeight / btnSelectMultiplePoint'in yanına
//      bir Button ekleyin:  Name = btnAutoSection,  Text = "Otomatik Kesit"
//   3) Butonun Click olayını btnAutoSection_Click'e bağlayın.
//  (mainform "partial class" olmalı; Visual Studio'nun ürettiği formlar zaten öyledir.)
//
//  Akış: .tow okunur -> AutoSectionDetector kesitleri bulur -> grid'e yazılır
//  -> kullanıcı kontrol eder / siler / ekler -> Run (mevcut akış, değişiklik yok).
// =====================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace OutlineDrawing
{
    public partial class mainform
    {
        private static readonly string[] AutoPointColumns =
            { "ColumnFirstPoint", "ColumnSecondPoint", "ColumnThirdPoint", "ColumnFourthPoint" };

        private void btnAutoSection_Click(object sender, EventArgs e)
        {
            string path = (txtTowerPath.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
            {
                PrintErrorLabel(5000, "Tower dosyası seçin.");
                return;
            }

            List<AutoSectionCandidate> found;
            try
            {
                dataprocessing data = dataprocessing.Run(path);
                if (data == null || data.finalMember == null || data.finalMember.Count == 0)
                {
                    PrintErrorLabel(5000, "Tower dosyası okunamadı.");
                    return;
                }
                found = AutoSectionDetector.Detect(data.finalMember, new AutoSectionOptions());
            }
            catch (Exception ex)
            {
                PrintErrorLabel(5000, "Otomatik kesit hatası: " + ex.Message);
                return;
            }

            if (found.Count == 0)
            {
                PrintErrorLabel(5000, "Otomatik kesit bulunamadı.");
                return;
            }

            bool hasRows = dataGridView1.Rows.Cast<DataGridViewRow>().Any(r => !r.IsNewRow && RowHasAnyValue(r));
            if (hasRows)
            {
                DialogResult ans = MessageBox.Show(
                    "Tablodaki mevcut kesitler silinsin mi?\n\nEvet = sil, otomatik kesitleri yaz\nHayır = mevcutların altına ekle",
                    "ŞA-RA ENERJİ", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (ans == DialogResult.Cancel) return;
                if (ans == DialogResult.Yes) dataGridView1.Rows.Clear();
            }

            foreach (AutoSectionCandidate c in found)
            {
                int idx = dataGridView1.Rows.Add();
                DataGridViewRow row = dataGridView1.Rows[idx];
                if (c.Kind == AutoSectionKind.Height)
                {
                    row.Cells["ColumnHeight"].Value = c.Height.Value.ToString("0.###");
                }
                else
                {
                    for (int k = 0; k < AutoPointColumns.Length && k < c.Points.Count; k++)
                        row.Cells[AutoPointColumns[k]].Value = AutoSectionDetector.FormatPoint(c.Points[k]);
                }
                row.Cells["ColumnSection"].ToolTipText = c.Description;   // fareyle üstüne gelince neden önerildiği görünür
            }

            RefreshSectionLetters();
            PrintErrorLabel(5000, string.Format("{0} otomatik kesit eklendi ({1} yatay, {2} eğik). Kontrol edip Run'a basın.",
                found.Count,
                found.Count(c => c.Kind == AutoSectionKind.Height),
                found.Count(c => c.Kind == AutoSectionKind.Plane)));
        }
    }
}
