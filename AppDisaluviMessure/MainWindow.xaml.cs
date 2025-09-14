using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using ClosedXML.Excel;
using System.IO;


namespace CortesAluminioWpf
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private const double MinResidue = 0.40;

        public enum PieceStatus { InPlay, Used, Excluded }

        public class PieceItem
        {
            public int Id { get; set; }
            public double Length { get; set; }
            public PieceStatus Status { get; set; } = PieceStatus.InPlay;
        }

        public class SavedLine
        {
            public double Original { get; set; }
            public List<double> Cuts { get; set; } = new();
            public string CutsDisplay => string.Join("  |  ", Cuts.Select(c => $"{c:0.##}m"));
            public double Residue { get; set; }
        }

        public ObservableCollection<PieceItem> Pieces { get; } = new();
        public ICollectionView PiecesView { get; private set; }
        public ObservableCollection<SavedLine> Saved { get; } = new();

        private double stockLen = 0;
        private double remaining = 0;
        private readonly List<double> currentCuts = new();

        private int _idSeq = 1;

        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            PiecesView = CollectionViewSource.GetDefaultView(Pieces);
            PiecesView.SortDescriptions.Add(new SortDescription(nameof(PieceItem.Length), ListSortDirection.Descending));

            GridPieces.ItemsSource = PiecesView;
            GridSaved.ItemsSource = Saved;

            RefrescarLabels();
        }

        // Helpers
        private static double ParseDouble(string s, double fallback = 0)
        {
            if (double.TryParse(s, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
                return v;
            if (double.TryParse(s, out v)) return v;
            return fallback;
        }
        private static int ParseInt(string s, int fallback = 1)
        {
            if (int.TryParse(s, out var n) && n > 0) return n;
            return fallback;
        }

        private void RefrescarLabels()
        {
            LblStockTotal.Text = $"{stockLen:0.##} m";
            LblRemaining.Text = $"{remaining:0.##} m";
            LblResumenActual.Text = currentCuts.Count == 0
                ? "—"
                : $"Pieza total = {stockLen:0.##} m   |   Cortes: {string.Join("  |  ", currentCuts.Select(c => $"{c:0.##}m"))}";
        }
        private void ReSortPieces() => PiecesView.Refresh();

        // Agregar (Enter + Cantidad)
        private void AddPieces(double length, int qty)
        {
            for (int k = 0; k < qty; k++)
                Pieces.Add(new PieceItem { Id = _idSeq++, Length = length, Status = PieceStatus.InPlay });

            ReSortPieces();
        }

        private void BtnAgregar_Click(object sender, RoutedEventArgs e)
        {
            var len = ParseDouble(TxtAddPiece.Text, -1);
            if (len <= 0)
            {
                MessageBox.Show("Ingresa un largo válido en metros (ej. 3 o 1.84).");
                return;
            }
            var qty = ParseInt(TxtQty.Text, 1);
            AddPieces(len, qty);

            TxtAddPiece.Clear();
            TxtQty.Text = "1";
        }

        private void TxtAddPiece_KeyDown(object sender, KeyEventArgs e)
        {   
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                BtnAgregar_Click(sender, new RoutedEventArgs());
            }
        }
        private void BtnQtyUp_Click(object sender, RoutedEventArgs e)
        {
            var q = ParseInt(TxtQty.Text, 1);
            TxtQty.Text = Math.Min(q + 1, 999).ToString();
        }
        private void BtnQtyDown_Click(object sender, RoutedEventArgs e)
        {
            var q = ParseInt(TxtQty.Text, 1);
            TxtQty.Text = Math.Max(q - 1, 1).ToString();
        }

        // Eliminar seleccionado (solo piezas InPlay)
        private void BtnEliminar_Click(object sender, RoutedEventArgs e)
        {
            if (GridPieces.SelectedItem is not PieceItem sel)
            {
                MessageBox.Show("Selecciona una pieza en la tabla para eliminarla.");
                return;
            }
            if (sel.Status != PieceStatus.InPlay)
            {
                MessageBox.Show("No puedes eliminar una pieza que ya fue usada en cortes.");
                return;
            }
            Pieces.Remove(sel);
            ReSortPieces();
        }

        // Flujo de cortes
        private void BtnIniciarCorte_Click(object sender, RoutedEventArgs e)
        {
            var stock = ParseDouble(TxtStock.Text, -1);
            if (stock <= 0)
            {
                MessageBox.Show("Ingresa la longitud de la pieza original (m).");
                return;
            }

            if (GridPieces.SelectedItem is not PieceItem sel)
            {
                MessageBox.Show("Selecciona la primera pieza a cortar en la tabla.");
                return;
            }
            if (sel.Status != PieceStatus.InPlay)
            {
                MessageBox.Show("La pieza seleccionada no está en juego.");
                return;
            }
            if (sel.Length > stock)
            {
                MessageBox.Show("La primera pieza seleccionada es mayor que la pieza original.");
                return;
            }

            stockLen = stock;
            remaining = stockLen;

            AplicaCorte(sel);
            LblSugerencia.Text = "—";

            if (remaining < MinResidue)
            {
                LblSugerencia.Text = $"Sobrante {remaining:0.##} m < 0.40 m. La barra se da por terminada.";
            }
            RefrescarLabels();
        }

        // Cortes ilimitados hasta terminar
        private void BtnBestCut_Click(object sender, RoutedEventArgs e)
        {
            if (stockLen <= 0)
            {
                MessageBox.Show("Primero inicia el corte (selecciona una pieza y pulsa 'Iniciar corte').");
                return;
            }

            var log = new List<string>();

            while (true)
            {
                if (remaining <= 0) { log.Add("No queda sobrante."); break; }
                if (remaining < MinResidue) { log.Add($"Sobrante ({remaining:0.##} m) < 0.40 m. La barra se da por terminada."); break; }

                var candidatos = Pieces.Where(p => p.Status == PieceStatus.InPlay && p.Length <= remaining).ToList();
                if (candidatos.Count == 0) { log.Add("No hay piezas ≤ sobrante."); break; }

                var candAsc = candidatos.OrderBy(p => p.Length).ToList();

                PieceItem bestSingle = candAsc.LastOrDefault(p => p.Length <= remaining);
                double singleResidue = bestSingle != null ? remaining - bestSingle.Length : double.PositiveInfinity;

                double bestPairSum = double.NegativeInfinity;
                PieceItem bestA = null, bestB = null;

                int i = 0, j = candAsc.Count - 1;
                while (i < j)
                {
                    double sum = candAsc[i].Length + candAsc[j].Length;
                    if (Math.Abs(sum - remaining) < 1e-9) { bestPairSum = sum; bestA = candAsc[i]; bestB = candAsc[j]; break; }
                    else if (sum < remaining) { if (sum > bestPairSum) { bestPairSum = sum; bestA = candAsc[i]; bestB = candAsc[j]; } i++; }
                    else { j--; }
                }
                double pairResidue = bestPairSum > 0 ? remaining - bestPairSum : double.PositiveInfinity;

                bool applied = false;
                if (pairResidue < singleResidue)
                {
                    if (bestA == null || bestB == null) { log.Add("No se encontró par válido."); break; }
                    log.Add($"Tomar 2 piezas: {bestA.Length:0.##} m + {bestB.Length:0.##} m  (residuo {pairResidue:0.##} m)");
                    AplicaCorte(bestA); AplicaCorte(bestB); applied = true;
                }
                else if (singleResidue < pairResidue)
                {
                    if (bestSingle == null) { log.Add("No se encontró pieza única válida."); break; }
                    log.Add($"Tomar 1 pieza: {bestSingle.Length:0.##} m  (residuo {singleResidue:0.##} m)");
                    AplicaCorte(bestSingle); applied = true;
                }
                else
                {
                    if (bestSingle != null) { log.Add($"Empate: se prefiere SINGLE {bestSingle.Length:0.##} m  (residuo {singleResidue:0.##} m)"); AplicaCorte(bestSingle); applied = true; }
                    else if (bestA != null && bestB != null) { log.Add($"Empate sin single; usar PAR {bestA.Length:0.##} m + {bestB.Length:0.##} m (residuo {pairResidue:0.##} m)"); AplicaCorte(bestA); AplicaCorte(bestB); applied = true; }
                    else { log.Add("No se encontró opción válida."); break; }
                }

                if (!applied) break;
            }

            LblSugerencia.Text = string.Join("  |  ", log);
            RefrescarLabels();
        }

        private void BtnGuardarLinea_Click(object sender, RoutedEventArgs e)
        {
            if (stockLen <= 0 || currentCuts.Count == 0)
            {
                MessageBox.Show("No hay cortes para guardar. Inicia un corte primero.");
                return;
            }

            var saved = new SavedLine
            {
                Original = stockLen,
                Cuts = currentCuts.ToList(),
                Residue = Math.Max(0, stockLen - currentCuts.Sum())
            };
            Saved.Add(saved);

            stockLen = 0;
            remaining = 0;
            currentCuts.Clear();
            LblSugerencia.Text = "—";
            RefrescarLabels();
        }

        private void BtnRestart_Click(object sender, RoutedEventArgs e)
        {
            Pieces.Clear();
            Saved.Clear();
            _idSeq = 1;

            stockLen = 0;
            remaining = 0;
            currentCuts.Clear();

            TxtStock.Text = "6";
            TxtAddPiece.Clear();
            TxtQty.Text = "1";
            LblSugerencia.Text = "—";
            RefrescarLabels();
        }

        // Aplicación de cortes
        private void AplicaCorte(PieceItem item)
        {
            item.Status = PieceStatus.Used;
            remaining = Math.Round(Math.Max(0, remaining - item.Length), 6);
            currentCuts.Add(item.Length);

            ReSortPieces();
            LblResumenActual.Text = $"Pieza total = {stockLen:0.##} m   |   Cortes: {string.Join("  |  ", currentCuts.Select(c => $"{c:0.##}m"))}";
        }

        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportSavedToExcel();
        }

        private void ExportSavedToExcel()
        {
            if (Saved.Count == 0)
            {
                MessageBox.Show("No hay líneas guardadas para exportar.");
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "plan_corte.xlsx"
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Plan");

                // Encabezados
                ws.Cell(1, 1).Value = "Pieza original (m)";
                ws.Cell(1, 2).Value = "Cortes";
                ws.Cell(1, 3).Value = "Residuo (m)";
                ws.Range(1, 1, 1, 3).Style.Font.Bold = true;

                // Filas
                int r = 2;
                foreach (var s in Saved)
                {
                    ws.Cell(r, 1).Value = s.Original;               // numérico
                    ws.Cell(r, 2).Value = s.CutsDisplay;            // texto "1.2m | 0.8m"
                    ws.Cell(r, 3).Value = s.Residue;                // numérico
                    r++;
                }

                // Formato (opcional)
                ws.Column(1).Style.NumberFormat.Format = "0.00";
                ws.Column(3).Style.NumberFormat.Format = "0.00";
                ws.Columns().AdjustToContents();

                wb.SaveAs(sfd.FileName);
                MessageBox.Show("Excel exportado correctamente.");
            }
            catch (IOException)
            {
                MessageBox.Show("El archivo está en uso o no se pudo escribir. Cierra el archivo si está abierto e inténtalo de nuevo.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se pudo exportar: {ex.Message}");
            }
        }

    }
}
