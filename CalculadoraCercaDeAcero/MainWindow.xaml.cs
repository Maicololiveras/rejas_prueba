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
        private List<string> pedidosConfirmados = new List<string>();

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

            MostrarImagenCerca(cantidadRejas, alturas[altura]);
        }

        private void MostrarImagenCerca(int cantidadRejas, float altura)
        {
            DrawingVisual drawingVisual = new DrawingVisual();
            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                for (int i = 0; i < cantidadRejas * 2; i++)
                {
                    drawingContext.DrawLine(new Pen(Brushes.Black, 2), new Point(i * LargoReja, 0), new Point(i * LargoReja, altura));
                }

                for (int i = 0; i < cantidadRejas; i++)
                {
                    double[] xValues = new double[101];
                    double[] yValues = new double[101];
                    for (int j = 0; j <= 100; j++)
                    {
                        double x = i * LargoReja + j * LargoReja / 100;
                        double y = Math.Sin(x * Math.PI / LargoReja) * (altura / 4) + altura / 2;
                        xValues[j] = x;
                        yValues[j] = y;
                    }
                    StreamGeometry streamGeometry = new StreamGeometry();
                    using (StreamGeometryContext geometryContext = streamGeometry.Open())
                    {
                        geometryContext.BeginFigure(new Point(xValues[0], yValues[0]), false, false);
                        for (int j = 1; j <= 100; j++)
                        {
                            geometryContext.LineTo(new Point(xValues[j], yValues[j]), true, false);
                        }
                    }
                    streamGeometry.Freeze();
                    drawingContext.DrawGeometry(null, new Pen(Brushes.Black, 2), streamGeometry);

                    for (int j = 0; j <= 100; j++)
                    {
                        yValues[j] -= altura / 2;
                    }
                    streamGeometry = new StreamGeometry();
                    using (StreamGeometryContext geometryContext = streamGeometry.Open())
                    {
                        geometryContext.BeginFigure(new Point(xValues[0], yValues[0]), false, false);
                        for (int j = 1; j <= 100; j++)
                        {
                            geometryContext.LineTo(new Point(xValues[j], yValues[j]), true, false);
                        }
                    }
                    streamGeometry.Freeze();
                    drawingContext.DrawGeometry(null, new Pen(Brushes.Black, 2), streamGeometry);
                }
            }

            RenderTargetBitmap renderBitmap = new RenderTargetBitmap(400, (int)(altura * 100), 96, 96, PixelFormats.Pbgra32);
            renderBitmap.Render(drawingVisual);

            imgCerca.Source = renderBitmap;
        }

        private void ConfirmarPedido_Click(object sender, RoutedEventArgs e)
        {
            float largoTotal = float.Parse(txtLength.Text);
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

        private void VerPedidosConfirmados_Click(object sender, RoutedEventArgs e)
        {
            lbPedidosConfirmados.ItemsSource = pedidosConfirmados;
        }

        private void ExportarTablaExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                FileInfo excelFile = new FileInfo("PedidosConfirmados.xlsx"); // Nombre del archivo de Excel

                // Verificar si el archivo ya existe
                if (!excelFile.Exists)
                {
                    // Crear un nuevo archivo Excel si no existe
                    using (ExcelPackage excelPackage = new ExcelPackage(excelFile))
                    {
                        // Crear hoja de cálculo
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Pedidos");

                        // Agregar encabezados de columna
                        worksheet.Cells[1, 1].Value = "Cliente";
                        worksheet.Cells[1, 2].Value = "Producto";
                        worksheet.Cells[1, 3].Value = "Cantidad";
                    }
                }

                // Abrir el archivo Excel existente
                using (ExcelPackage excelPackage = new ExcelPackage(excelFile))
                {
                    // Obtener la hoja de cálculo "Pedidos"
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Pedidos"];

                    // Obtener el número de la próxima fila disponible
                    int nextRow = worksheet.Dimension?.Rows ?? 1;

                    // Agregar los datos de los pedidos confirmados a la hoja de cálculo
                    for (int i = 0; i < lbPedidosConfirmados.Items.Count; i++)
                    {
                        // Obtener el pedido confirmado actual
                        string[] pedido = lbPedidosConfirmados.Items[i].ToString().Split(',');

                        // Agregar el pedido a la hoja de cálculo
                        worksheet.Cells[nextRow + i, 1].Value = pedido[0].Trim(); // Cliente
                        worksheet.Cells[nextRow + i, 2].Value = pedido[1].Trim(); // Producto
                        worksheet.Cells[nextRow + i, 3].Value = pedido[2].Trim(); // Cantidad
                    }

                    // Guardar el archivo Excel
                    excelPackage.Save();
                }

                MessageBox.Show("Datos exportados correctamente a la hoja de cálculo 'Pedidos'.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al exportar datos a la hoja de cálculo: {ex.Message}");
            }
        }


    }
}
