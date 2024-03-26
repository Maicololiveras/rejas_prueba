using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OfficeOpenXml;
using System.Text;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;



namespace CalculadoraCercaDeAcero
{
    public partial class MainWindow : Window
    {
        private const float LargoReja = 2.5f;
        private readonly Dictionary<int, float> alturas = new Dictionary<int, float> { { 1, 1.03f }, { 2, 1.53f }, { 3, 2.03f } };
        private readonly Dictionary<int, int> fijadores = new Dictionary<int, int> { { 1, 3 }, { 2, 4 }, { 3, 6 } };
        private readonly Dictionary<int, string> pinturas = new Dictionary<int, string> { { 0, "Sin pintura" }, { 1, "Blanca" }, { 2, "Negra" }, { 3, "Verde" } };
        private readonly int ancho = 14;
        private List<string> pedidosConfirmados = new List<string>();
        private string excelFilePath = "PedidosConfirmados.xlsx";

        public MainWindow()
        {
            InitializeComponent();
            cmbHeight.ItemsSource = alturas;
            cmbColor.ItemsSource = pinturas;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Calculate_Click(object sender, RoutedEventArgs e)
        {
            if (!float.TryParse(txtLength.Text, out float largoTotal) || largoTotal <= 0)
            {
                MessageBox.Show("Ingrese un largo válido.");
                return;
            }

            int altura = (int)cmbHeight.SelectedValue;
            int color = (int)cmbColor.SelectedValue;

            int cantidadRejas = (int)Math.Ceiling(largoTotal / LargoReja);
            int cantidadPostes = cantidadRejas * 2;
            int cantidadTornillos = cantidadPostes * 4;
            int cantidadFijadores = fijadores[altura];
            string colorPintura = pinturas[color];

            lblCantidadRejas.Text = cantidadRejas.ToString();
            lblCantidadPostes.Text = cantidadPostes.ToString();
            lblCantidadTornillos.Text = cantidadTornillos.ToString();
            lblCantidadFijadores.Text = cantidadFijadores.ToString();
            lblColorPintura.Text = colorPintura;

            MostrarImagenFrente(cantidadRejas, alturas[altura], ancho);
        }

        private void MostrarImagenFrente(int cantidadRejas, float altura, float ancho)
        {
            DrawingVisual drawingVisual = new DrawingVisual();

            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                // Dibujar las rejas
                for (int i = 0; i < cantidadRejas; i++)
                {
                    // Coordenadas de la reja
                    double x1 = i * ancho;
                    double y1 = 0;
                    double x2 = x1;
                    double y2 = altura;

                    // Dibujar la reja como un rectángulo vertical
                    drawingContext.DrawRectangle(Brushes.Black, new Pen(Brushes.Black, 2), new Rect(new Point(x1, y1), new Point(x2 + 10, y2)));
                }

                // Dibujar las ondas sinusoidales
                for (int i = 0; i < cantidadRejas; i++)
                {
                    // Coordenadas de la onda sinusoidal
                    PointCollection points = new PointCollection();
                    for (int j = 0; j <= 100; j++)
                    {
                        // Calcular coordenadas x e y
                        double x = i * ancho + j * ancho / 100;
                        double y = Math.Sin(j * Math.PI / 50) * (altura / 4) + altura / 2;
                        points.Add(new Point(x, y));
                    }

                    // Dibujar la onda sinusoidal como una polilínea
                    Polyline polyline = new Polyline();
                    polyline.Points = points;
                    polyline.Stroke = Brushes.Black;
                    polyline.StrokeThickness = 2;
                    drawingContext.DrawGeometry(null, new Pen(Brushes.Black, 2), polyline.RenderedGeometry);
                }
            }
            RenderTargetBitmap renderBitmap = new RenderTargetBitmap(400, (int)(altura * 100), 96, 96, PixelFormats.Pbgra32);
            renderBitmap.Render(drawingVisual);

            imgCerca.Source = renderBitmap;
        }

        private void VerPedidosConfirmados_Click(object sender, RoutedEventArgs e)
        {
            ExportarTablaExcel_Click(sender, e); // Llamar al método para exportar los datos a Excel

            // Luego, utilizar la ruta del archivo predefinida para abrirlo
            VerPedidosConfirmados(excelFilePath);
        }

        private void ExportarTablaExcel_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = "PedidosConfirmados.xlsx";

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                FileInfo excelFile = new FileInfo("PedidosConfirmados.xlsx"); // Nombre del archivo de Excel

                // Crear un nuevo archivo Excel o abrir el existente
                using (ExcelPackage excelPackage = new ExcelPackage(excelFile))
                {
                    // Crear hoja de cálculo si no existe
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Pedidos");
                    if (worksheet == null)
                    {
                        worksheet = excelPackage.Workbook.Worksheets.Add("Pedidos");

                        // Agregar encabezados de columna
                        worksheet.Cells[1, 1].Value = "Cliente";
                        worksheet.Cells[1, 2].Value = "Producto";
                        worksheet.Cells[1, 3].Value = "Cantidad";
                    }

                    // Obtener el número de la próxima fila disponible
                    int nextRow = worksheet.Dimension?.Rows ?? 1;

                    // Agregar los datos de los pedidos confirmados a la hoja de cálculo
                    for (int i = 0; i < lbPedidosConfirmados.Items.Count; i++)
                    {
                        // Obtener el pedido confirmado actual
                        string[] pedido = lbPedidosConfirmados.Items[i].ToString().Split(',');

                        // Agregar el pedido a la hoja de cálculo
                        worksheet.Cells[nextRow + i + 1, 1].Value = pedido[0].Trim(); // Cliente
                        worksheet.Cells[nextRow + i + 1, 2].Value = pedido[1].Trim(); // Producto
                        worksheet.Cells[nextRow + i + 1, 3].Value = pedido[2].Trim(); // Cantidad
                    }

                    // Guardar el archivo Excel
                    excelPackage.Save();
                }

                MessageBox.Show("Datos exportados correctamente a la hoja de cálculo 'Pedidos'.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar datos a la hoja de cálculo: {ex.Message}");
                excelFilePath = null; // En caso de error, establecer la ruta como nula
            }
        }

        private void ConfirmarPedido_Click(object sender, RoutedEventArgs e)
        {
            float largoTotal;
            if (!float.TryParse(txtLength.Text, out largoTotal) || largoTotal <= 0)
            {
                MessageBox.Show("Ingrese un largo válido.");
                return;
            }

            int altura = (int)cmbHeight.SelectedValue;
            int color = (int)cmbColor.SelectedValue;
            bool confirmado = chkConfirmado.IsChecked ?? false;

            if (confirmado)
            {
                GuardarPedidoConfirmado(largoTotal, alturas[altura], pinturas[color]);
            }
        }

        private void GuardarPedidoConfirmado(float largoTotal, float altura, string colorPintura)
        {
            string pedidoConfirmado = $"Largo: {largoTotal}, Altura: {altura}, Color: {colorPintura}";
            pedidosConfirmados.Add(pedidoConfirmado);
        }

        

        private void VerPedidosConfirmados(string excelFilePath)
        {
            lbPedidosConfirmados.ItemsSource = null; // Limpiar la lista antes de asignarla
            lbPedidosConfirmados.ItemsSource = pedidosConfirmados;

            // Verificar si el archivo existe
            if (File.Exists(excelFilePath))
            {
                // Abrir el archivo Excel con la aplicación predeterminada
                try
                {
                    System.Diagnostics.Process.Start(excelFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al abrir el archivo de Excel: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("El archivo de Excel no existe. Asegúrate de haber Guardado los pedidos confirmados primero.");
            }
        }

    }

}
