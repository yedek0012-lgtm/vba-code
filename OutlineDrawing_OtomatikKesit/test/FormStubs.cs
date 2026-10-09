// SADECE DERLEME KONTROLU ICIN: mainform ve dataprocessing'in, butonun kullandigi
// uyelerinin imzalari (OutlineDrawing.dll'den okunmustur). Gercek projeye EKLEMEYIN.
using System.Collections.Generic;
using System.Windows.Forms;

namespace OutlineDrawing
{
    internal class dataprocessing
    {
        public Dictionary<string, FinalMember> finalMember { get; set; }
        public static dataprocessing Run(string path) { return null; }
    }

    public partial class mainform : Form
    {
        private TextBox txtTowerPath = new TextBox();
        private DataGridView dataGridView1 = new DataGridView();
        private void PrintErrorLabel(int durationMilliseconds, string message) { }
        private static bool RowHasAnyValue(DataGridViewRow r) { return false; }
        private void RefreshSectionLetters() { }
    }
}
