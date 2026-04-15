using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyExpirationDate("2026-12-31")]
[assembly: ESAPIScript(IsWriteable = true)]

public class AssemblyExpirationDate : Attribute
{
    private readonly string expirationDate;

    public string ExpirationDate
    {
        get { return expirationDate; }
    }

    public AssemblyExpirationDate(string expirationDate)
    {
        this.expirationDate = expirationDate;
    }
}

public static class SimpleLicenseVerifier
{
    private const string HARDCODED_ACCESS_CODE = "2b5d777c";
    private const string APP_SETTINGS_SUBDIR = "MAAS-ProtonSpotSequencer-Settings";

    private static string GetAppSettingsDir()
    {
        string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string appSpecificDir = Path.Combine(appDataDir, APP_SETTINGS_SUBDIR);
        try
        {
            Directory.CreateDirectory(appSpecificDir);
        }
        catch (Exception ex)
        {
            Debug.WriteLine(string.Format(CultureInfo.InvariantCulture, "Failed to create directory {0} in AppData: {1}", APP_SETTINGS_SUBDIR, ex.Message));
        }

        return appSpecificDir;
    }

    private static string GetLicenseFilePath(string projectName)
    {
        return Path.Combine(GetAppSettingsDir(), projectName + "_license.txt");
    }

    public static string GetNoExpireFilePath()
    {
        return Path.Combine(GetAppSettingsDir(), "NOEXPIRE");
    }

    public static string GetNoAgreeFilePath()
    {
        return Path.Combine(GetAppSettingsDir(), "NoAgree.txt");
    }

    public static string GetValidatedFilePath(string projectName)
    {
        return Path.Combine(GetAppSettingsDir(), projectName + "_validated.txt");
    }

    public static bool IsLicenseAccepted(string projectName, string version)
    {
        string licenseFile = GetLicenseFilePath(projectName);
        if (!File.Exists(licenseFile))
        {
            return false;
        }

        try
        {
            string storedCode = File.ReadAllText(licenseFile).Trim();
            return VerifyAccessCode(storedCode);
        }
        catch
        {
            return false;
        }
    }

    public static bool ShowLicenseDialog(string projectName, string version, string licenseUrl)
    {
        Window licenseWindow = new Window();
        licenseWindow.Title = "License Required";
        licenseWindow.Width = 450;
        licenseWindow.Height = 280;
        licenseWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        licenseWindow.ResizeMode = ResizeMode.NoResize;
        licenseWindow.Background = Brushes.White;

        StackPanel mainPanel = new StackPanel();
        mainPanel.Margin = new Thickness(20);

        TextBlock title = new TextBlock();
        title.Text = "License Acceptance Required";
        title.FontSize = 16;
        title.FontWeight = FontWeights.Bold;
        title.Margin = new Thickness(0, 0, 0, 15);
        mainPanel.Children.Add(title);

        TextBlock instructions = new TextBlock();
        instructions.TextWrapping = TextWrapping.Wrap;
        instructions.Margin = new Thickness(0, 0, 0, 20);
        instructions.Inlines.Add(new Run("Please visit the following URL to obtain your access code:\n"));

        Hyperlink licenseLink = new Hyperlink(new Run(licenseUrl));
        licenseLink.NavigateUri = new Uri(licenseUrl, UriKind.Absolute);
        licenseLink.RequestNavigate += OnLicenseLinkRequestNavigate;
        instructions.Inlines.Add(licenseLink);

        instructions.Inlines.Add(new Run("\n\nEnter the access code below to continue."));
        mainPanel.Children.Add(instructions);

        TextBlock codeLabel = new TextBlock();
        codeLabel.Text = "Access Code:";
        codeLabel.FontWeight = FontWeights.Bold;
        codeLabel.Margin = new Thickness(0, 0, 0, 5);
        mainPanel.Children.Add(codeLabel);

        TextBox codeTextBox = new TextBox();
        codeTextBox.Width = 200;
        codeTextBox.Height = 25;
        codeTextBox.HorizontalAlignment = HorizontalAlignment.Left;
        codeTextBox.Margin = new Thickness(0, 0, 0, 20);
        mainPanel.Children.Add(codeTextBox);

        StackPanel buttonPanel = new StackPanel();
        buttonPanel.Orientation = Orientation.Horizontal;
        buttonPanel.HorizontalAlignment = HorizontalAlignment.Right;

        Button okButton = new Button();
        okButton.Content = "OK";
        okButton.Width = 80;
        okButton.Height = 25;
        okButton.Margin = new Thickness(0, 0, 10, 0);
        okButton.IsDefault = true;

        Button cancelButton = new Button();
        cancelButton.Content = "Cancel";
        cancelButton.Width = 80;
        cancelButton.Height = 25;
        cancelButton.IsCancel = true;

        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);
        mainPanel.Children.Add(buttonPanel);

        licenseWindow.Content = mainPanel;

        bool result = false;
        okButton.Click += (sender, e) =>
        {
            string enteredCode = codeTextBox.Text.Trim();
            if (VerifyAccessCode(enteredCode))
            {
                try
                {
                    File.WriteAllText(GetLicenseFilePath(projectName), enteredCode);
                    result = true;
                    licenseWindow.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not save license acceptance: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("Invalid access code. Please try again.", "Invalid Code", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        };

        cancelButton.Click += (sender, e) =>
        {
            result = false;
            licenseWindow.Close();
        };

        licenseWindow.Loaded += (sender, e) => codeTextBox.Focus();
        licenseWindow.ShowDialog();
        return result;
    }

    private static void OnLicenseLinkRequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        e.Handled = true;
    }

    private static bool VerifyAccessCode(string inputCode)
    {
        return string.Equals(inputCode, HARDCODED_ACCESS_CODE, StringComparison.OrdinalIgnoreCase);
    }
}

namespace VMS.TPS
{
    internal sealed class SpotInfo
    {
        public SpotInfo(double x, double y, double monitorUnits, int layerIndex, double nominalEnergy)
        {
            X = x;
            Y = y;
            MonitorUnits = monitorUnits;
            LayerIndex = layerIndex;
            NominalEnergy = nominalEnergy;
        }

        public double X { get; set; }

        public double Y { get; set; }

        public double MonitorUnits { get; set; }

        public int LayerIndex { get; private set; }

        public double NominalEnergy { get; private set; }

        public SpotInfo Copy()
        {
            return new SpotInfo(X, Y, MonitorUnits, LayerIndex, NominalEnergy);
        }
    }

    internal sealed class StructureOutlineInfo
    {
        public StructureOutlineInfo(string name, Color color)
        {
            Name = name;
            Color = color;
            Contours = new List<List<Point>>();
        }

        public string Name { get; private set; }

        public Color Color { get; private set; }

        public List<List<Point>> Contours { get; private set; }
    }

    internal sealed class FieldState
    {
        public FieldState(string id)
        {
            Id = id;
            Structures = new List<StructureOutlineInfo>();
            OriginalSpots = new List<SpotInfo>();
            SequencedSpots = new List<SpotInfo>();
            DefaultXMin = -10.0;
            DefaultXMax = 10.0;
            DefaultYMin = -10.0;
            DefaultYMax = 10.0;
            ResetView();
        }

        public string Id { get; private set; }

        public List<StructureOutlineInfo> Structures { get; private set; }

        public List<SpotInfo> OriginalSpots { get; private set; }

        public List<SpotInfo> SequencedSpots { get; private set; }

        public double DefaultXMin { get; private set; }

        public double DefaultXMax { get; private set; }

        public double DefaultYMin { get; private set; }

        public double DefaultYMax { get; private set; }

        public double ViewXMin { get; set; }

        public double ViewXMax { get; set; }

        public double ViewYMin { get; set; }

        public double ViewYMax { get; set; }

        public bool HasPendingSequence
        {
            get { return SequencedSpots.Count > 0; }
        }

        public void ResetView()
        {
            ViewXMin = DefaultXMin;
            ViewXMax = DefaultXMax;
            ViewYMin = DefaultYMin;
            ViewYMax = DefaultYMax;
        }

        public void UpdateDefaultBounds()
        {
            bool hasBounds = false;
            double minX = double.MaxValue;
            double maxX = double.MinValue;
            double minY = double.MaxValue;
            double maxY = double.MinValue;

            foreach (SpotInfo spot in OriginalSpots)
            {
                ExtendBounds(spot.X, spot.Y, ref minX, ref maxX, ref minY, ref maxY, ref hasBounds);
            }

            if (!hasBounds)
            {
                foreach (StructureOutlineInfo structure in Structures)
                {
                    foreach (List<Point> contour in structure.Contours)
                    {
                        foreach (Point point in contour)
                        {
                            ExtendBounds(point.X, point.Y, ref minX, ref maxX, ref minY, ref maxY, ref hasBounds);
                        }
                    }
                }
            }

            if (!hasBounds)
            {
                minX = -10.0;
                maxX = 10.0;
                minY = -10.0;
                maxY = 10.0;
            }

            if (Math.Abs(maxX - minX) < 0.001)
            {
                minX -= 10.0;
                maxX += 10.0;
            }

            if (Math.Abs(maxY - minY) < 0.001)
            {
                minY -= 10.0;
                maxY += 10.0;
            }

            DefaultXMin = minX - 20.0;
            DefaultXMax = maxX + 20.0;
            DefaultYMin = minY - 20.0;
            DefaultYMax = maxY + 20.0;
            ResetView();
        }

        private static void ExtendBounds(double x, double y, ref double minX, ref double maxX, ref double minY, ref double maxY, ref bool hasBounds)
        {
            if (!hasBounds)
            {
                minX = maxX = x;
                minY = maxY = y;
                hasBounds = true;
                return;
            }

            minX = Math.Min(minX, x);
            maxX = Math.Max(maxX, x);
            minY = Math.Min(minY, y);
            maxY = Math.Max(maxY, y);
        }
    }

    internal sealed class HitTestResult
    {
        public HitTestResult(SpotInfo nearestSpot, int count)
        {
            NearestSpot = nearestSpot;
            Count = count;
        }

        public SpotInfo NearestSpot { get; private set; }

        public int Count { get; private set; }
    }

    internal sealed class BevCanvas : FrameworkElement
    {
        private const double Padding = 12.0;
        private const double HitThresholdPx = 10.0;

        private FieldState fieldState;
        private ISet<string> selectedStructureNames;
        private SpotInfo hoveredSpot;
        private int hoveredSpotCount;

        public BevCanvas()
        {
            Focusable = true;
            ClipToBounds = true;
            Cursor = Cursors.Arrow;
            selectedStructureNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            MouseMove += OnCanvasMouseMove;
            MouseLeave += OnCanvasMouseLeave;
            MouseLeftButtonDown += OnCanvasMouseLeftButtonDown;
            MouseWheel += OnCanvasMouseWheel;
            QueryCursor += OnCanvasQueryCursor;
            SizeChanged += (sender, e) => InvalidateVisual();
        }

        public event Action<SpotInfo> SpotChosen;

        public event Action UndoRequested;

        public event Action<string> StatusChanged;

        public FieldState FieldState
        {
            get { return fieldState; }
            set
            {
                fieldState = value;
                hoveredSpot = null;
                hoveredSpotCount = 0;
                InvalidateVisual();
            }
        }

        public ISet<string> SelectedStructureNames
        {
            get { return selectedStructureNames; }
            set
            {
                selectedStructureNames = value ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                InvalidateVisual();
            }
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);
            drawingContext.DrawRectangle(Brushes.Black, null, new Rect(0, 0, ActualWidth, ActualHeight));

            if (fieldState == null || ActualWidth <= 0 || ActualHeight <= 0)
            {
                return;
            }

            double scale;
            double offsetX;
            double offsetY;
            if (!TryGetTransform(out scale, out offsetX, out offsetY))
            {
                return;
            }

            DrawStructures(drawingContext, scale, offsetX, offsetY);
            DrawOriginalPattern(drawingContext, scale, offsetX, offsetY);
            DrawSequencedPattern(drawingContext, scale, offsetX, offsetY);

            if (hoveredSpot != null && hoveredSpotCount == 1)
            {
                Point hoverPoint = WorldToScreen(hoveredSpot.X, hoveredSpot.Y, scale, offsetX, offsetY);
                drawingContext.DrawEllipse(null, new Pen(Brushes.Yellow, 2.0), hoverPoint, 6.0, 6.0);
            }

        }

        private void DrawStructures(DrawingContext drawingContext, double scale, double offsetX, double offsetY)
        {
            foreach (StructureOutlineInfo structure in fieldState.Structures)
            {
                if (!selectedStructureNames.Contains(structure.Name))
                {
                    continue;
                }

                Pen pen = new Pen(new SolidColorBrush(structure.Color), 1.0);
                foreach (List<Point> contour in structure.Contours)
                {
                    if (contour.Count < 2)
                    {
                        continue;
                    }

                    StreamGeometry geometry = new StreamGeometry();
                    using (StreamGeometryContext context = geometry.Open())
                    {
                        context.BeginFigure(WorldToScreen(contour[0].X, contour[0].Y, scale, offsetX, offsetY), false, true);
                        for (int i = 1; i < contour.Count; i++)
                        {
                            context.LineTo(WorldToScreen(contour[i].X, contour[i].Y, scale, offsetX, offsetY), true, false);
                        }
                    }

                    geometry.Freeze();
                    drawingContext.DrawGeometry(null, pen, geometry);
                }
            }
        }

        private void DrawOriginalPattern(DrawingContext drawingContext, double scale, double offsetX, double offsetY)
        {
            if (fieldState.OriginalSpots.Count > 1)
            {
                Pen linePen = new Pen(Brushes.Red, 1.0);
                linePen.DashStyle = DashStyles.Dot;

                StreamGeometry geometry = new StreamGeometry();
                using (StreamGeometryContext context = geometry.Open())
                {
                    SpotInfo firstSpot = fieldState.OriginalSpots[0];
                    context.BeginFigure(WorldToScreen(firstSpot.X, firstSpot.Y, scale, offsetX, offsetY), false, false);
                    for (int i = 1; i < fieldState.OriginalSpots.Count; i++)
                    {
                        SpotInfo spot = fieldState.OriginalSpots[i];
                        context.LineTo(WorldToScreen(spot.X, spot.Y, scale, offsetX, offsetY), true, false);
                    }
                }

                geometry.Freeze();
                drawingContext.DrawGeometry(null, linePen, geometry);
            }

            foreach (SpotInfo spot in fieldState.OriginalSpots)
            {
                Point screenPoint = WorldToScreen(spot.X, spot.Y, scale, offsetX, offsetY);
                drawingContext.DrawEllipse(Brushes.Red, null, screenPoint, 2.5, 2.5);
            }

            if (fieldState.OriginalSpots.Count > 0)
            {
                SpotInfo firstSpot = fieldState.OriginalSpots[0];
                Point screenPoint = WorldToScreen(firstSpot.X, firstSpot.Y, scale, offsetX, offsetY);
                drawingContext.DrawEllipse(null, new Pen(Brushes.Gold, 2.0), screenPoint, 6.0, 6.0);
            }
        }

        private void DrawSequencedPattern(DrawingContext drawingContext, double scale, double offsetX, double offsetY)
        {
            if (fieldState.SequencedSpots.Count > 1)
            {
                Pen linePen = new Pen(Brushes.LimeGreen, 1.5);
                StreamGeometry geometry = new StreamGeometry();
                using (StreamGeometryContext context = geometry.Open())
                {
                    SpotInfo firstSpot = fieldState.SequencedSpots[0];
                    context.BeginFigure(WorldToScreen(firstSpot.X, firstSpot.Y, scale, offsetX, offsetY), false, false);
                    for (int i = 1; i < fieldState.SequencedSpots.Count; i++)
                    {
                        SpotInfo spot = fieldState.SequencedSpots[i];
                        context.LineTo(WorldToScreen(spot.X, spot.Y, scale, offsetX, offsetY), true, false);
                    }
                }

                geometry.Freeze();
                drawingContext.DrawGeometry(null, linePen, geometry);
            }

            foreach (SpotInfo spot in fieldState.SequencedSpots)
            {
                Point screenPoint = WorldToScreen(spot.X, spot.Y, scale, offsetX, offsetY);
                drawingContext.DrawEllipse(Brushes.LimeGreen, null, screenPoint, 3.0, 3.0);
            }
        }

        private void OnCanvasMouseMove(object sender, MouseEventArgs e)
        {
            if (fieldState == null)
            {
                return;
            }

            Point screenPoint = e.GetPosition(this);
            Point worldPoint = ScreenToWorld(screenPoint);
            HitTestResult hit = FindHit(screenPoint);

            SpotInfo nextHover = hit.Count == 1 ? hit.NearestSpot : null;
            bool hoverChanged = !AreSameSpot(hoveredSpot, nextHover) || hoveredSpotCount != hit.Count;
            hoveredSpot = nextHover;
            hoveredSpotCount = hit.Count;

            if (hoverChanged)
            {
                InvalidateVisual();
            }

            string status = string.Format(CultureInfo.InvariantCulture, "x: {0:0.0}, y: {1:0.0}", worldPoint.X, worldPoint.Y);
            if (hit.Count == 1 && hit.NearestSpot != null)
            {
                status += string.Format(CultureInfo.InvariantCulture, " | MU: {0:0.####}, Layer: {1}, E: {2:0.0} MeV", hit.NearestSpot.MonitorUnits, hit.NearestSpot.LayerIndex + 1, hit.NearestSpot.NominalEnergy);
            }
            else if (hit.Count > 1)
            {
                status += " | Multiple overlapping spots under cursor";
            }

            RaiseStatusChanged(status);
        }

        private void OnCanvasMouseLeave(object sender, MouseEventArgs e)
        {
            if (hoveredSpot != null || hoveredSpotCount != 0)
            {
                hoveredSpot = null;
                hoveredSpotCount = 0;
                InvalidateVisual();
            }

            RaiseStatusChanged("Ready");
        }

        private void OnCanvasMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (fieldState == null)
            {
                return;
            }

            Focus();
            HitTestResult hit = FindHit(e.GetPosition(this));

            if (e.ClickCount > 1)
            {
                if (hit.Count > 0 && UndoRequested != null)
                {
                    UndoRequested();
                }

                e.Handled = true;
                return;
            }

            if (hit.Count == 1 && hit.NearestSpot != null && SpotChosen != null)
            {
                SpotChosen(hit.NearestSpot.Copy());
                e.Handled = true;
            }
        }

        private void OnCanvasMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (fieldState == null)
            {
                return;
            }

            Point worldPoint = ScreenToWorld(e.GetPosition(this));
            ZoomAround(worldPoint, e.Delta / 120.0);
            InvalidateVisual();
            e.Handled = true;
        }

        private void OnCanvasQueryCursor(object sender, QueryCursorEventArgs e)
        {
            e.Cursor = Cursors.Arrow;
            e.Handled = true;
        }

        private void ZoomAround(Point worldPoint, double step)
        {
            double factor = step <= 0.0 ? 1.0 - (0.1 * step) : 1.0 / (1.0 + (0.1 * step));

            double leftX = worldPoint.X - fieldState.ViewXMin;
            double rightX = fieldState.ViewXMax - worldPoint.X;
            double bottomY = worldPoint.Y - fieldState.ViewYMin;
            double topY = fieldState.ViewYMax - worldPoint.Y;

            fieldState.ViewXMin = worldPoint.X - (leftX * factor);
            fieldState.ViewXMax = worldPoint.X + (rightX * factor);
            fieldState.ViewYMin = worldPoint.Y - (bottomY * factor);
            fieldState.ViewYMax = worldPoint.Y + (topY * factor);
        }

        private HitTestResult FindHit(Point screenPoint)
        {
            if (fieldState == null || fieldState.OriginalSpots.Count == 0)
            {
                return new HitTestResult(null, 0);
            }

            double scale;
            double offsetX;
            double offsetY;
            if (!TryGetTransform(out scale, out offsetX, out offsetY))
            {
                return new HitTestResult(null, 0);
            }

            double thresholdSquared = HitThresholdPx * HitThresholdPx;
            SpotInfo nearestSpot = null;
            double nearestDistanceSquared = double.MaxValue;
            int hitCount = 0;

            foreach (SpotInfo spot in fieldState.OriginalSpots)
            {
                Point spotPoint = WorldToScreen(spot.X, spot.Y, scale, offsetX, offsetY);
                double dx = spotPoint.X - screenPoint.X;
                double dy = spotPoint.Y - screenPoint.Y;
                double distanceSquared = (dx * dx) + (dy * dy);

                if (distanceSquared <= thresholdSquared)
                {
                    hitCount++;
                    if (distanceSquared < nearestDistanceSquared)
                    {
                        nearestDistanceSquared = distanceSquared;
                        nearestSpot = spot;
                    }
                }
            }

            return new HitTestResult(nearestSpot, hitCount);
        }

        private bool TryGetTransform(out double scale, out double offsetX, out double offsetY)
        {
            scale = 1.0;
            offsetX = 0.0;
            offsetY = 0.0;

            if (fieldState == null)
            {
                return false;
            }

            double viewWidth = fieldState.ViewXMax - fieldState.ViewXMin;
            double viewHeight = fieldState.ViewYMax - fieldState.ViewYMin;
            double pixelWidth = Math.Max(ActualWidth - (2.0 * Padding), 1.0);
            double pixelHeight = Math.Max(ActualHeight - (2.0 * Padding), 1.0);

            if (viewWidth <= 0.0 || viewHeight <= 0.0)
            {
                return false;
            }

            scale = Math.Min(pixelWidth / viewWidth, pixelHeight / viewHeight);
            double renderedWidth = viewWidth * scale;
            double renderedHeight = viewHeight * scale;

            offsetX = (ActualWidth - renderedWidth) / 2.0;
            offsetY = (ActualHeight - renderedHeight) / 2.0;
            return true;
        }

        private Point WorldToScreen(double x, double y, double scale, double offsetX, double offsetY)
        {
            double screenX = offsetX + ((x - fieldState.ViewXMin) * scale);
            double screenY = ActualHeight - offsetY - ((y - fieldState.ViewYMin) * scale);
            return new Point(screenX, screenY);
        }

        private Point ScreenToWorld(Point screenPoint)
        {
            double scale;
            double offsetX;
            double offsetY;
            if (!TryGetTransform(out scale, out offsetX, out offsetY))
            {
                return new Point();
            }

            double x = fieldState.ViewXMin + ((screenPoint.X - offsetX) / scale);
            double y = fieldState.ViewYMin + ((ActualHeight - offsetY - screenPoint.Y) / scale);
            return new Point(x, y);
        }

        private static bool AreSameSpot(SpotInfo left, SpotInfo right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (left == null || right == null)
            {
                return false;
            }

            return left.LayerIndex == right.LayerIndex
                && Math.Abs(left.X - right.X) < 0.0001
                && Math.Abs(left.Y - right.Y) < 0.0001
                && Math.Abs(left.MonitorUnits - right.MonitorUnits) < 0.0001;
        }

        private void RaiseStatusChanged(string text)
        {
            if (StatusChanged != null)
            {
                StatusChanged(text);
            }
        }
    }

    public sealed class ProtonSpotSequencer : UserControl
    {
        private readonly ScriptContext context;
        private readonly bool isValidated;

        private readonly Dictionary<string, FieldState> fieldsById;
        private readonly HashSet<string> selectedStructureNames;

        private readonly List<FieldState> fields;

        private ComboBox fieldSelector;
        private StackPanel structureCheckPanel;
        private ListBox sequenceList;
        private TextBlock patientText;
        private TextBlock planText;
        private TextBlock fieldStatsText;
        private TextBlock sequenceSummaryText;
        private TextBlock modifiedFieldsText;
        private TextBlock statusText;
        private Button applyButton;
        private Button removeLastButton;
        private Button clearActiveButton;
        private Button clearAllButton;
        private BevCanvas bevCanvas;

        private bool suppressFieldSelectionChanged;
        private bool suppressStructureEvents;
        private bool modificationsBegun;

        private FieldState activeField;

        public ProtonSpotSequencer(ScriptContext context, Window window, ScriptEnvironment environment, bool isValidated)
        {
            if (context.Patient == null)
            {
                throw new InvalidOperationException("No patient currently loaded.");
            }

            if (context.IonPlanSetup == null)
            {
                throw new InvalidOperationException("There are no proton plans opened. Please open a proton plan.");
            }

            if (context.IonPlanSetup.StructureSet == null)
            {
                throw new InvalidOperationException("The active proton plan does not have a structure set.");
            }

            if (!context.IonPlanSetup.IonBeams.Any())
            {
                throw new InvalidOperationException("The active proton plan does not contain any ion beams.");
            }

            this.context = context;
            this.isValidated = isValidated;
            fieldsById = new Dictionary<string, FieldState>(StringComparer.OrdinalIgnoreCase);
            selectedStructureNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            fields = new List<FieldState>();

            BuildUi();
            LoadPlanData();
        }

        private void BuildUi()
        {
            Grid root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = isValidated ? new GridLength(0.0) : GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1.0, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            if (!isValidated)
            {
                Label topBanner = new Label();
                topBanner.Content = "* * * NOT VALIDATED FOR CLINICAL USE * * *";
                topBanner.Background = Brushes.PaleVioletRed;
                topBanner.Foreground = Brushes.Black;
                topBanner.FontWeight = FontWeights.Bold;
                topBanner.FontSize = 14.0;
                topBanner.Padding = new Thickness(0, 2, 0, 2);
                topBanner.HorizontalContentAlignment = HorizontalAlignment.Center;
                Grid.SetRow(topBanner, 0);
                root.Children.Add(topBanner);
            }

            Grid contentGrid = new Grid();
            contentGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1.0, GridUnitType.Star) });
            contentGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Grid.SetRow(contentGrid, 1);
            root.Children.Add(contentGrid);

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.0, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(380.0) });
            headerGrid.Margin = new Thickness(5, 5, 5, 0);
            Grid.SetRow(headerGrid, 0);
            contentGrid.Children.Add(headerGrid);

            Border infoBorder = CreateSectionBorder();
            Grid.SetColumn(infoBorder, 0);
            headerGrid.Children.Add(infoBorder);

            Grid infoGrid = new Grid();
            infoGrid.Margin = new Thickness(10);
            infoGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            infoGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.0, GridUnitType.Star) });
            infoGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            infoGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            infoGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            infoGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            infoBorder.Child = infoGrid;

            infoGrid.Children.Add(CreateCaptionText("Patient:", 0, 0));
            patientText = CreateValueText(0, 1);
            infoGrid.Children.Add(patientText);

            infoGrid.Children.Add(CreateCaptionText("Plan:", 1, 0));
            planText = CreateValueText(1, 1);
            infoGrid.Children.Add(planText);

            infoGrid.Children.Add(CreateCaptionText("Field:", 2, 0));
            fieldSelector = new ComboBox();
            fieldSelector.Margin = new Thickness(0, 0, 0, 6);
            fieldSelector.MinWidth = 240;
            fieldSelector.SelectionChanged += OnFieldSelectionChanged;
            Grid.SetRow(fieldSelector, 2);
            Grid.SetColumn(fieldSelector, 1);
            infoGrid.Children.Add(fieldSelector);

            fieldStatsText = CreateValueText(3, 0);
            Grid.SetColumnSpan(fieldStatsText, 2);
            infoGrid.Children.Add(fieldStatsText);

            Border actionBorder = CreateSectionBorder();
            Grid.SetColumn(actionBorder, 1);
            actionBorder.Margin = new Thickness(5, 0, 0, 0);
            headerGrid.Children.Add(actionBorder);

            StackPanel actionPanel = new StackPanel();
            actionPanel.Margin = new Thickness(10);
            actionBorder.Child = actionPanel;

            WrapPanel buttonPanel = new WrapPanel();
            buttonPanel.Margin = new Thickness(0, 0, 0, 8);
            actionPanel.Children.Add(buttonPanel);

            applyButton = new Button();
            applyButton.Content = "Apply To Plan";
            applyButton.Margin = new Thickness(0, 0, 8, 0);
            applyButton.Padding = new Thickness(10, 4, 10, 4);
            applyButton.Click += (sender, e) => ApplyChangesToPlan();
            buttonPanel.Children.Add(applyButton);

            Button resetViewButton = new Button();
            resetViewButton.Content = "Reset View";
            resetViewButton.Margin = new Thickness(0, 0, 8, 0);
            resetViewButton.Padding = new Thickness(10, 4, 10, 4);
            resetViewButton.Click += (sender, e) => ResetView();
            buttonPanel.Children.Add(resetViewButton);

            Button helpButton = new Button();
            helpButton.Content = "?";
            helpButton.Width = 28;
            helpButton.Height = 28;
            helpButton.Click += (sender, e) => ShowHelp();
            buttonPanel.Children.Add(helpButton);

            sequenceSummaryText = CreateReadOnlyBlock();
            modifiedFieldsText = CreateReadOnlyBlock();
            TextBlock instructionText = CreateReadOnlyBlock();
            instructionText.TextWrapping = TextWrapping.Wrap;
            instructionText.Text = "Left-click a red spot to append it to the green sequence. Use mouse wheel to zoom.";
            actionPanel.Children.Add(sequenceSummaryText);
            actionPanel.Children.Add(modifiedFieldsText);
            actionPanel.Children.Add(instructionText);

            Grid bodyGrid = new Grid();
            bodyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(240.0) });
            bodyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.0, GridUnitType.Star) });
            bodyGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(260.0) });
            bodyGrid.Margin = new Thickness(5);
            Grid.SetRow(bodyGrid, 1);
            contentGrid.Children.Add(bodyGrid);

            Border structureBorder = CreateSectionBorder();
            Grid.SetColumn(structureBorder, 0);
            bodyGrid.Children.Add(structureBorder);

            DockPanel structurePanel = new DockPanel();
            structurePanel.Margin = new Thickness(10);
            structureBorder.Child = structurePanel;

            TextBlock structureHeader = new TextBlock();
            structureHeader.Text = "Structures";
            structureHeader.FontWeight = FontWeights.Bold;
            structureHeader.Margin = new Thickness(0, 0, 0, 8);
            DockPanel.SetDock(structureHeader, Dock.Top);
            structurePanel.Children.Add(structureHeader);

            Button toggleAllStructuresButton = new Button();
            toggleAllStructuresButton.Content = "Toggle All";
            toggleAllStructuresButton.Margin = new Thickness(0, 0, 0, 8);
            toggleAllStructuresButton.Padding = new Thickness(10, 4, 10, 4);
            toggleAllStructuresButton.Click += (sender, e) => ToggleAllStructures();
            DockPanel.SetDock(toggleAllStructuresButton, Dock.Top);
            structurePanel.Children.Add(toggleAllStructuresButton);

            ScrollViewer structureScroll = new ScrollViewer();
            structureScroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            structurePanel.Children.Add(structureScroll);

            structureCheckPanel = new StackPanel();
            structureScroll.Content = structureCheckPanel;

            Border canvasBorder = CreateSectionBorder();
            Grid.SetColumn(canvasBorder, 1);
            canvasBorder.Margin = new Thickness(5, 0, 5, 0);
            bodyGrid.Children.Add(canvasBorder);

            bevCanvas = new BevCanvas();
            bevCanvas.SelectedStructureNames = selectedStructureNames;
            bevCanvas.StatusChanged += SetStatus;
            bevCanvas.SpotChosen += AddSpotToSequence;
            bevCanvas.UndoRequested += RemoveLastSequenceSpot;
            canvasBorder.Child = bevCanvas;

            Border sequenceBorder = CreateSectionBorder();
            Grid.SetColumn(sequenceBorder, 2);
            bodyGrid.Children.Add(sequenceBorder);

            Grid sequenceGrid = new Grid();
            sequenceGrid.Margin = new Thickness(10);
            sequenceGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            sequenceGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            sequenceGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1.0, GridUnitType.Star) });
            sequenceGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            sequenceGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            sequenceBorder.Child = sequenceGrid;

            TextBlock sequenceHeader = new TextBlock();
            sequenceHeader.Text = "Sequenced Spots";
            sequenceHeader.FontWeight = FontWeights.Bold;
            Grid.SetRow(sequenceHeader, 0);
            sequenceGrid.Children.Add(sequenceHeader);

            TextBlock sequenceNote = CreateReadOnlyBlock();
            sequenceNote.Text = "Double-click a red spot to remove the last sequenced spot.";
            sequenceNote.Margin = new Thickness(0, 4, 0, 8);
            Grid.SetRow(sequenceNote, 1);
            sequenceGrid.Children.Add(sequenceNote);

            sequenceList = new ListBox();
            Grid.SetRow(sequenceList, 2);
            sequenceGrid.Children.Add(sequenceList);

            WrapPanel sequenceButtons = new WrapPanel();
            sequenceButtons.Margin = new Thickness(0, 8, 0, 0);
            Grid.SetRow(sequenceButtons, 3);
            sequenceGrid.Children.Add(sequenceButtons);

            removeLastButton = new Button();
            removeLastButton.Content = "Remove Last";
            removeLastButton.Margin = new Thickness(0, 0, 8, 0);
            removeLastButton.Padding = new Thickness(10, 4, 10, 4);
            removeLastButton.Click += (sender, e) => RemoveLastSequenceSpot();
            sequenceButtons.Children.Add(removeLastButton);

            clearActiveButton = new Button();
            clearActiveButton.Content = "Clear Active";
            clearActiveButton.Margin = new Thickness(0, 0, 8, 0);
            clearActiveButton.Padding = new Thickness(10, 4, 10, 4);
            clearActiveButton.Click += (sender, e) => ClearActiveSequence();
            sequenceButtons.Children.Add(clearActiveButton);

            clearAllButton = new Button();
            clearAllButton.Content = "Clear All";
            clearAllButton.Padding = new Thickness(10, 4, 10, 4);
            clearAllButton.Click += (sender, e) => ClearAllSequences();
            sequenceButtons.Children.Add(clearAllButton);

            TextBlock sequenceFooter = CreateReadOnlyBlock();
            sequenceFooter.Text = "Export is intentionally removed in this version so the tool operates fully in memory.";
            sequenceFooter.TextWrapping = TextWrapping.Wrap;
            sequenceFooter.Margin = new Thickness(0, 8, 0, 0);
            Grid.SetRow(sequenceFooter, 4);
            sequenceGrid.Children.Add(sequenceFooter);

            Border statusBorder = new Border();
            statusBorder.Background = Brushes.Black;
            statusBorder.BorderBrush = Brushes.Brown;
            statusBorder.BorderThickness = new Thickness(1, 1, 1, 1);
            statusBorder.Margin = new Thickness(5, 0, 5, 5);
            statusBorder.Padding = new Thickness(8, 4, 8, 4);
            Grid.SetRow(statusBorder, 2);
            contentGrid.Children.Add(statusBorder);

            statusText = new TextBlock();
            statusText.Foreground = Brushes.White;
            statusText.Text = "Ready";
            statusBorder.Child = statusText;

            Border footerBorder = new Border();
            footerBorder.Background = Brushes.PaleVioletRed;
            footerBorder.Padding = new Thickness(8, 4, 8, 4);
            Grid.SetRow(footerBorder, 2);
            root.Children.Add(footerBorder);

            TextBlock footerText = new TextBlock();
            footerText.TextWrapping = TextWrapping.Wrap;
            footerText.Inlines.Add(new Run("Bound by the terms of the "));
            Hyperlink hyperlink = new Hyperlink(new Run("Varian LUSLA"));
            hyperlink.NavigateUri = new Uri("http://medicalaffairs.varian.com/download/VarianLUSLA.pdf");
            hyperlink.RequestNavigate += Hyperlink_RequestNavigate;
            footerText.Inlines.Add(hyperlink);
            if (!isValidated)
            {
                footerText.Inlines.Add(new Run("   * * * NOT VALIDATED FOR CLINICAL USE * * *"));
            }

            footerBorder.Child = footerText;
            Content = root;
        }

        private void LoadPlanData()
        {
            string previousFieldId = activeField != null ? activeField.Id : null;
            HashSet<string> previousSelections = new HashSet<string>(selectedStructureNames, StringComparer.OrdinalIgnoreCase);

            fields.Clear();
            fieldsById.Clear();

            foreach (IonBeam beam in context.IonPlanSetup.IonBeams)
            {
                FieldState field = BuildFieldState(beam);
                fields.Add(field);
                fieldsById[field.Id] = field;
            }

            patientText.Text = BuildPatientLabel();
            planText.Text = context.IonPlanSetup.Id;

            PopulateFieldSelector(previousFieldId);
            PopulateStructurePanel(previousSelections);

            if (fieldSelector.Items.Count > 0)
            {
                if (!string.IsNullOrWhiteSpace(previousFieldId) && fieldsById.ContainsKey(previousFieldId))
                {
                    fieldSelector.SelectedItem = previousFieldId;
                }
                else
                {
                    fieldSelector.SelectedIndex = 0;
                }
            }

            SetStatus("Ready");
        }

        private FieldState BuildFieldState(IonBeam beam)
        {
            FieldState field = new FieldState(beam.Id);
            LoadStructureOutlines(beam, field);
            LoadOriginalSpots(beam, field);
            field.UpdateDefaultBounds();
            return field;
        }

        private void LoadStructureOutlines(IonBeam beam, FieldState field)
        {
            foreach (Structure structure in context.IonPlanSetup.StructureSet.Structures)
            {
                if (!structure.HasSegment || structure.IsEmpty)
                {
                    continue;
                }

                try
                {
                    StructureOutlineInfo outlineInfo = new StructureOutlineInfo(structure.Id, ParseEsapiColor(structure.Color.ToString()));
                    foreach (var outline in beam.GetStructureOutlines(structure, true))
                    {
                        List<Point> contour = new List<Point>();
                        foreach (var point in outline)
                        {
                            contour.Add(new Point(point.X, point.Y));
                        }

                        if (contour.Count > 1)
                        {
                            outlineInfo.Contours.Add(contour);
                        }
                    }

                    if (outlineInfo.Contours.Count > 0)
                    {
                        field.Structures.Add(outlineInfo);
                    }
                }
                catch
                {
                }
            }
        }

        private void LoadOriginalSpots(IonBeam beam, FieldState field)
        {
            double totalMetersetWeight = 0.0;
            if (beam.IonControlPoints.Any())
            {
                totalMetersetWeight = ConvertRequiredToDouble(beam.IonControlPoints.Last().MetersetWeight, "Total meterset weight");
            }

            double muPerWeight = 1.0;
            if (totalMetersetWeight > 0.0)
            {
                double beamMetersetValue = ConvertRequiredToDouble(beam.Meterset.Value, "Beam meterset value");
                muPerWeight = beamMetersetValue / totalMetersetWeight;
            }

            var editableParameters = beam.GetEditableParameters();
            int layerIndex = 0;
            foreach (var controlPointPair in editableParameters.IonControlPointPairs)
            {
                var spotList = controlPointPair.RawSpotList.Count > 0 ? controlPointPair.RawSpotList : controlPointPair.FinalSpotList;
                double nominalBeamEnergy = ConvertRequiredToDouble(controlPointPair.NominalBeamEnergy, "Nominal beam energy");
                foreach (var spot in spotList)
                {
                    field.OriginalSpots.Add(new SpotInfo(spot.X, spot.Y, spot.Weight * muPerWeight, layerIndex, nominalBeamEnergy));
                }

                layerIndex++;
            }
        }

        private void PopulateFieldSelector(string previousFieldId)
        {
            suppressFieldSelectionChanged = true;
            fieldSelector.Items.Clear();
            foreach (FieldState field in fields)
            {
                fieldSelector.Items.Add(field.Id);
            }

            suppressFieldSelectionChanged = false;
        }

        private void PopulateStructurePanel(HashSet<string> previousSelections)
        {
            suppressStructureEvents = true;
            structureCheckPanel.Children.Clear();
            selectedStructureNames.Clear();

            List<string> orderedNames = new List<string>();
            if (fields.Count > 0)
            {
                foreach (string name in fields[0].Structures.Select(structure => structure.Name))
                {
                    if (!orderedNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                    {
                        orderedNames.Add(name);
                    }
                }
            }

            foreach (string name in fields
                .SelectMany(field => field.Structures.Select(structure => structure.Name))
                .Where(name => !orderedNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            {
                orderedNames.Add(name);
            }

            foreach (string name in orderedNames)
            {
                bool isChecked = previousSelections.Contains(name);
                if (isChecked)
                {
                    selectedStructureNames.Add(name);
                }

                CheckBox checkBox = new CheckBox();
                checkBox.Content = name;
                checkBox.Margin = new Thickness(0, 0, 0, 4);
                checkBox.IsChecked = isChecked;
                checkBox.Checked += OnStructureCheckChanged;
                checkBox.Unchecked += OnStructureCheckChanged;
                structureCheckPanel.Children.Add(checkBox);
            }

            suppressStructureEvents = false;
            bevCanvas.InvalidateVisual();
        }

        private void OnFieldSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (suppressFieldSelectionChanged)
            {
                return;
            }

            string fieldId = fieldSelector.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(fieldId) || !fieldsById.ContainsKey(fieldId))
            {
                return;
            }

            activeField = fieldsById[fieldId];
            bevCanvas.FieldState = activeField;
            RefreshSequenceList();
            UpdateSummaries();
        }

        private void OnStructureCheckChanged(object sender, RoutedEventArgs e)
        {
            if (suppressStructureEvents)
            {
                return;
            }

            CheckBox checkBox = sender as CheckBox;
            if (checkBox == null || checkBox.Content == null)
            {
                return;
            }

            string structureName = checkBox.Content.ToString();
            if (checkBox.IsChecked == true)
            {
                selectedStructureNames.Add(structureName);
            }
            else
            {
                selectedStructureNames.Remove(structureName);
            }

            bevCanvas.InvalidateVisual();
        }

        private void ToggleAllStructures()
        {
            bool selectAll = !selectedStructureNames.Any();
            suppressStructureEvents = true;
            selectedStructureNames.Clear();

            foreach (CheckBox checkBox in structureCheckPanel.Children.OfType<CheckBox>())
            {
                checkBox.IsChecked = selectAll;
                if (selectAll && checkBox.Content != null)
                {
                    selectedStructureNames.Add(checkBox.Content.ToString());
                }
            }

            suppressStructureEvents = false;
            bevCanvas.InvalidateVisual();
        }

        private void AddSpotToSequence(SpotInfo spot)
        {
            if (activeField == null || spot == null)
            {
                return;
            }

            activeField.SequencedSpots.Add(spot.Copy());
            RefreshSequenceList();
            UpdateSummaries();
            bevCanvas.InvalidateVisual();
        }

        private void RemoveLastSequenceSpot()
        {
            if (activeField == null || activeField.SequencedSpots.Count == 0)
            {
                return;
            }

            activeField.SequencedSpots.RemoveAt(activeField.SequencedSpots.Count - 1);
            RefreshSequenceList();
            UpdateSummaries();
            bevCanvas.InvalidateVisual();
        }

        private void ClearActiveSequence()
        {
            if (activeField == null || activeField.SequencedSpots.Count == 0)
            {
                return;
            }

            activeField.SequencedSpots.Clear();
            RefreshSequenceList();
            UpdateSummaries();
            bevCanvas.InvalidateVisual();
        }

        private void ClearAllSequences()
        {
            if (!fields.Any(field => field.HasPendingSequence))
            {
                return;
            }

            foreach (FieldState field in fields)
            {
                field.SequencedSpots.Clear();
            }

            RefreshSequenceList();
            UpdateSummaries();
            bevCanvas.InvalidateVisual();
        }

        private void ResetView()
        {
            if (activeField == null)
            {
                return;
            }

            activeField.ResetView();
            bevCanvas.InvalidateVisual();
        }

        private void RefreshSequenceList()
        {
            sequenceList.Items.Clear();
            if (activeField == null)
            {
                return;
            }

            for (int i = 0; i < activeField.SequencedSpots.Count; i++)
            {
                SpotInfo spot = activeField.SequencedSpots[i];
                string label = string.Format(CultureInfo.InvariantCulture, "#{0:000}: X={1:0.00}, Y={2:0.00}, MU={3:0.####}, Layer={4}", i + 1, spot.X, spot.Y, spot.MonitorUnits, spot.LayerIndex + 1);
                sequenceList.Items.Add(label);
            }

            if (sequenceList.Items.Count > 0)
            {
                sequenceList.ScrollIntoView(sequenceList.Items[sequenceList.Items.Count - 1]);
            }
        }

        private void UpdateSummaries()
        {
            if (activeField == null)
            {
                fieldStatsText.Text = "No field selected.";
                sequenceSummaryText.Text = "Active sequence: 0 spot(s)";
                modifiedFieldsText.Text = "Modified fields: none";
                applyButton.IsEnabled = false;
                removeLastButton.IsEnabled = false;
                clearActiveButton.IsEnabled = false;
                clearAllButton.IsEnabled = false;
                return;
            }

            fieldStatsText.Text = string.Format(CultureInfo.InvariantCulture, "Original spots: {0}\nStructures with outlines: {1}", activeField.OriginalSpots.Count, activeField.Structures.Count);
            sequenceSummaryText.Text = string.Format(CultureInfo.InvariantCulture, "Active sequence: {0} spot(s)", activeField.SequencedSpots.Count);

            List<string> modifiedFieldIds = fields.Where(field => field.HasPendingSequence).Select(field => field.Id).ToList();
            modifiedFieldsText.Text = modifiedFieldIds.Count == 0
                ? "Modified fields: none"
                : "Modified fields: " + string.Join(", ", modifiedFieldIds.ToArray());

            bool hasAnyPendingSequence = modifiedFieldIds.Count > 0;
            applyButton.IsEnabled = hasAnyPendingSequence;
            removeLastButton.IsEnabled = activeField.SequencedSpots.Count > 0;
            clearActiveButton.IsEnabled = activeField.SequencedSpots.Count > 0;
            clearAllButton.IsEnabled = hasAnyPendingSequence;
        }

        private void ApplyChangesToPlan()
        {
            List<FieldState> modifiedFields = fields.Where(field => field.HasPendingSequence).ToList();
            if (modifiedFields.Count == 0)
            {
                MessageBox.Show("No pending sequence changes were found.", "Nothing To Apply", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string confirmation = string.Format(CultureInfo.InvariantCulture,
                "Apply the pending sequenced spots to {0} field(s) and recalculate dose?\n\nThis writes the same selected sequence to every energy layer in each modified field.",
                modifiedFields.Count);
            if (MessageBox.Show(confirmation, "Confirm Apply", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                SetStatus("Applying sequence changes to the plan...");

                if (!modificationsBegun)
                {
                    context.Patient.BeginModifications();
                    modificationsBegun = true;
                }

                IonPlanSetup plan = context.IonPlanSetup;
                double planNormalizationValue = ConvertRequiredToDouble(plan.PlanNormalizationValue, "Plan normalization value");
                if (planNormalizationValue == 0.0)
                {
                    throw new InvalidOperationException("Plan normalization value is zero. Cannot convert monitor units back to spot weights.");
                }

                double treatmentPercentage = ConvertRequiredToDouble(plan.TreatmentPercentage, "Treatment percentage");
                if (treatmentPercentage == 0.0)
                {
                    throw new InvalidOperationException("Treatment percentage is zero. Cannot convert monitor units back to spot weights.");
                }

                double numberOfFractions = ConvertRequiredToDouble(plan.NumberOfFractions, "Number of fractions");
                if (numberOfFractions == 0.0)
                {
                    throw new InvalidOperationException("Number of fractions is zero. Cannot convert monitor units back to spot weights.");
                }

                DoseValue? dosePerFraction = plan.DosePerFraction;
                if (!dosePerFraction.HasValue)
                {
                    throw new InvalidOperationException("Dose per fraction is not defined. Cannot convert monitor units back to spot weights.");
                }

                double? dosePerFractionDose = dosePerFraction.Value.Dose;
                if (!dosePerFractionDose.HasValue)
                {
                    throw new InvalidOperationException("Dose per fraction value is not defined. Cannot convert monitor units back to spot weights.");
                }

                double prescribedDose = dosePerFractionDose.Value * numberOfFractions;
                double planNorm = 100.0 / planNormalizationValue;

                Dictionary<string, object> editableParametersByBeamId = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (IonBeam beam in plan.IonBeams)
                {
                    FieldState field;
                    if (!fieldsById.TryGetValue(beam.Id, out field) || !field.HasPendingSequence)
                    {
                        continue;
                    }

                    var editableParameters = beam.GetEditableParameters();
                    double fieldNorm = ConvertRequiredToDouble(editableParameters.WeightFactor, "Beam weight factor");
                    double muPerWeight = (planNorm * prescribedDose * fieldNorm) / treatmentPercentage;
                    if (muPerWeight <= 0.0)
                    {
                        throw new InvalidOperationException("Derived MU-per-weight conversion factor is invalid for beam '" + beam.Id + "'.");
                    }

                    foreach (var controlPointPair in editableParameters.IonControlPointPairs)
                    {
                        if (controlPointPair.RawSpotList.Count > 0)
                        {
                            controlPointPair.ResizeRawSpotList(field.SequencedSpots.Count);
                            for (int i = 0; i < field.SequencedSpots.Count; i++)
                            {
                                SpotInfo spot = field.SequencedSpots[i];
                                controlPointPair.RawSpotList[i].Weight = (float)(spot.MonitorUnits / muPerWeight);
                                controlPointPair.RawSpotList[i].X = (float)spot.X;
                                controlPointPair.RawSpotList[i].Y = (float)spot.Y;
                            }
                        }
                        else
                        {
                            controlPointPair.ResizeFinalSpotList(field.SequencedSpots.Count);
                            for (int i = 0; i < field.SequencedSpots.Count; i++)
                            {
                                SpotInfo spot = field.SequencedSpots[i];
                                controlPointPair.FinalSpotList[i].Weight = (float)(spot.MonitorUnits / muPerWeight);
                                controlPointPair.FinalSpotList[i].X = (float)spot.X;
                                controlPointPair.FinalSpotList[i].Y = (float)spot.Y;
                            }
                        }
                    }

                    editableParametersByBeamId[beam.Id] = editableParameters;
                }

                foreach (IonBeam beam in plan.IonBeams)
                {
                    object editableParameters;
                    if (!editableParametersByBeamId.TryGetValue(beam.Id, out editableParameters))
                    {
                        continue;
                    }

                    InvokeApplyParameters(beam, editableParameters);
                }

                string doseError;
                if (!TryRecalculateDose(plan, out doseError))
                {
                    throw new InvalidOperationException(doseError);
                }

                LoadPlanData();
                MessageBox.Show("The sequenced spots were applied successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Apply Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                SetStatus("Apply failed.");
            }
            finally
            {
                Mouse.OverrideCursor = null;
                bevCanvas.InvalidateVisual();
                UpdateSummaries();
            }
        }

        private void ShowHelp()
        {
            string info = string.Join(Environment.NewLine, new[]
            {
                "Left-click a red spot to append it to the green sequence.",
                "Double-click a red spot, or use Remove Last, to remove the last sequenced spot.",
                "Use the mouse wheel over the canvas to zoom.",
                "Toggle structures in the left panel to show or hide outlines.",
                "Apply To Plan writes pending sequences back into the active plan and recalculates dose.",
            });

            MessageBox.Show(info, "Proton Spot Sequencer Help", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SetStatus(string text)
        {
            statusText.Text = text;
        }

        private string BuildPatientLabel()
        {
            string firstName = string.IsNullOrWhiteSpace(context.Patient.FirstName) ? string.Empty : context.Patient.FirstName.Trim();
            string lastName = string.IsNullOrWhiteSpace(context.Patient.LastName) ? string.Empty : context.Patient.LastName.Trim();
            string displayName = (firstName + " " + lastName).Trim();
            if (displayName.Length == 0)
            {
                displayName = "Unknown";
            }

            return displayName + " (ID: " + context.Patient.Id + ")";
        }

        private static void InvokeApplyParameters(IonBeam beam, object editableParameters)
        {
            MethodInfo method = beam.GetType()
                .GetMethods()
                .FirstOrDefault(candidate => candidate.Name == "ApplyParameters" && candidate.GetParameters().Length == 1);

            if (method == null)
            {
                throw new MissingMethodException("Could not locate IonBeam.ApplyParameters.");
            }

            method.Invoke(beam, new[] { editableParameters });
        }

        private static bool TryRecalculateDose(IonPlanSetup plan, out string errorMessage)
        {
            errorMessage = null;
            try
            {
                MethodInfo calculateDoseMethod = plan.GetType().GetMethod("CalculateDose", Type.EmptyTypes);
                if (calculateDoseMethod == null)
                {
                    return true;
                }

                object result = calculateDoseMethod.Invoke(plan, null);
                if (result == null)
                {
                    return true;
                }

                PropertyInfo successProperty = result.GetType().GetProperty("Success");
                if (successProperty != null && successProperty.PropertyType == typeof(bool))
                {
                    bool success = (bool)successProperty.GetValue(result, null);
                    if (!success)
                    {
                        errorMessage = "Dose calculation failed after applying the sequenced spots.";
                        return false;
                    }
                }

                return true;
            }
            catch (TargetInvocationException ex)
            {
                errorMessage = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static double ConvertRequiredToDouble(object value, string valueName)
        {
            if (value == null)
            {
                throw new InvalidOperationException(valueName + " is not defined.");
            }

            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        private static Border CreateSectionBorder()
        {
            return new Border
            {
                BorderBrush = Brushes.Brown,
                BorderThickness = new Thickness(1, 1, 2, 2),
                CornerRadius = new CornerRadius(3),
                Background = Brushes.WhiteSmoke
            };
        }

        private static TextBlock CreateCaptionText(string caption, int row, int column)
        {
            TextBlock textBlock = new TextBlock();
            textBlock.Text = caption;
            textBlock.FontWeight = FontWeights.Bold;
            textBlock.Margin = new Thickness(0, 0, 10, 6);
            Grid.SetRow(textBlock, row);
            Grid.SetColumn(textBlock, column);
            return textBlock;
        }

        private static TextBlock CreateValueText(int row, int column)
        {
            TextBlock textBlock = new TextBlock();
            textBlock.Margin = new Thickness(0, 0, 0, 6);
            textBlock.TextWrapping = TextWrapping.Wrap;
            Grid.SetRow(textBlock, row);
            Grid.SetColumn(textBlock, column);
            return textBlock;
        }

        private static TextBlock CreateReadOnlyBlock()
        {
            return new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 6),
                TextWrapping = TextWrapping.Wrap
            };
        }

        private static Color ParseEsapiColor(string colorText)
        {
            if (!string.IsNullOrWhiteSpace(colorText) && colorText[0] == '#')
            {
                string hex = colorText.Substring(1);
                if (hex.Length == 8)
                {
                    byte alpha = byte.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte red = byte.Parse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte green = byte.Parse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte blue = byte.Parse(hex.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    return Color.FromArgb(alpha, red, green, blue);
                }

                if (hex.Length == 6)
                {
                    byte red = byte.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte green = byte.Parse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    byte blue = byte.Parse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    return Color.FromRgb(red, green, blue);
                }
            }

            return Colors.White;
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }
    }

    public class Script
    {
        private const string PROJECT_NAME = "ProtonSpotSequencer";
        private const string PROJECT_VERSION = "1.0.0";
        private const string LICENSE_URL = "https://varian-medicalaffairsappliedsolutions.github.io/MAAS-ProtonSpotSequencer";
        private const string GITHUB_URL = "https://github.com/Varian-MedicalAffairsAppliedSolutions/MAAS-ProtonSpotSequencer";

        private bool IsValidated = false;

        private static MessageBoxResult ShowLinkedDialog(Window owner, string title, string textBeforeLink, string linkUrl, string textAfterLink, MessageBoxButton buttons)
        {
            Window dialog = new Window();
            dialog.Title = title;
            dialog.Width = 620;
            dialog.Height = 260;
            dialog.MinWidth = 620;
            dialog.MinHeight = 260;
            dialog.WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            dialog.ResizeMode = ResizeMode.NoResize;
            dialog.Background = Brushes.White;
            dialog.ShowInTaskbar = false;

            if (owner != null)
            {
                dialog.Owner = owner;
            }

            Grid root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1.0, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBlock messageText = new TextBlock();
            messageText.Margin = new Thickness(20, 20, 20, 12);
            messageText.TextWrapping = TextWrapping.Wrap;
            messageText.Inlines.Add(new Run(textBeforeLink ?? string.Empty));

            if (!string.IsNullOrWhiteSpace(linkUrl))
            {
                Hyperlink hyperlink = new Hyperlink(new Run(linkUrl));
                hyperlink.NavigateUri = new Uri(linkUrl, UriKind.Absolute);
                hyperlink.RequestNavigate += OnAgreementLinkRequestNavigate;
                messageText.Inlines.Add(hyperlink);
            }

            messageText.Inlines.Add(new Run(textAfterLink ?? string.Empty));
            Grid.SetRow(messageText, 0);
            root.Children.Add(messageText);

            StackPanel buttonPanel = new StackPanel();
            buttonPanel.Orientation = Orientation.Horizontal;
            buttonPanel.HorizontalAlignment = HorizontalAlignment.Right;
            buttonPanel.Margin = new Thickness(20, 0, 20, 20);

            MessageBoxResult result = buttons == MessageBoxButton.YesNo ? MessageBoxResult.No : MessageBoxResult.OK;

            if (buttons == MessageBoxButton.YesNo)
            {
                Button yesButton = new Button();
                yesButton.Content = "Yes";
                yesButton.Width = 90;
                yesButton.Height = 28;
                yesButton.Margin = new Thickness(0, 0, 10, 0);
                yesButton.IsDefault = true;
                yesButton.Click += (sender, e) =>
                {
                    result = MessageBoxResult.Yes;
                    dialog.DialogResult = true;
                    dialog.Close();
                };
                buttonPanel.Children.Add(yesButton);

                Button noButton = new Button();
                noButton.Content = "No";
                noButton.Width = 90;
                noButton.Height = 28;
                noButton.IsCancel = true;
                noButton.Click += (sender, e) =>
                {
                    result = MessageBoxResult.No;
                    dialog.DialogResult = false;
                    dialog.Close();
                };
                buttonPanel.Children.Add(noButton);
            }
            else
            {
                Button okButton = new Button();
                okButton.Content = "OK";
                okButton.Width = 90;
                okButton.Height = 28;
                okButton.IsDefault = true;
                okButton.IsCancel = true;
                okButton.Click += (sender, e) =>
                {
                    result = MessageBoxResult.OK;
                    dialog.DialogResult = true;
                    dialog.Close();
                };
                buttonPanel.Children.Add(okButton);
            }

            Grid.SetRow(buttonPanel, 1);
            root.Children.Add(buttonPanel);

            dialog.Content = root;
            dialog.ShowDialog();
            return result;
        }

        private static void OnAgreementLinkRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context, Window window, ScriptEnvironment environment)
        {
            try
            {
                bool noExpire = File.Exists(SimpleLicenseVerifier.GetNoExpireFilePath());
                bool skipAgreement = File.Exists(SimpleLicenseVerifier.GetNoAgreeFilePath());

                if (!skipAgreement && !SimpleLicenseVerifier.IsLicenseAccepted(PROJECT_NAME, PROJECT_VERSION))
                {
                    if (!SimpleLicenseVerifier.ShowLicenseDialog(PROJECT_NAME, PROJECT_VERSION, LICENSE_URL))
                    {
                        MessageBox.Show("License acceptance is required to use this application.", "License Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                CustomAttributeData expirationAttribute = typeof(Script).Assembly.CustomAttributes
                    .FirstOrDefault(attribute => attribute.AttributeType == typeof(AssemblyExpirationDate));

                DateTime endDate;
                if (expirationAttribute != null && DateTime.TryParse(expirationAttribute.ConstructorArguments.FirstOrDefault().Value as string, new CultureInfo("en-US"), DateTimeStyles.None, out endDate))
                {
                    if (DateTime.Now > endDate && !noExpire)
                    {
                        ShowLinkedDialog(
                            window,
                            "Application Expired",
                            "Application has expired. Newer builds with future expiration dates can be found here: ",
                            GITHUB_URL,
                            string.Empty,
                            MessageBoxButton.OK);
                        return;
                    }

                    string agreementMessageSuffix = "\n\nSee the FAQ for more information on how to remove this pop-up and expiration.";
                    string firstUseMessagePrefix = "The current " + PROJECT_NAME + " application is provided AS IS as a non-clinical, research only tool in evaluation only. The current application will only be available until " + endDate.ToShortDateString() + " after which the application will be unavailable. By clicking 'Yes' you agree that this application will be evaluated and not utilized in providing planning decision support.\n\nNewer builds with future expiration dates can be found here: ";
                    string repeatUseMessagePrefix = "Application will only be available until " + endDate.ToShortDateString() + " after which the application will be unavailable. By clicking 'Yes' you agree that this application will be evaluated and not utilized in providing planning decision support.\n\nNewer builds with future expiration dates can be found here: ";

                    bool isValidatedForUser = IsValidated;
                    string validatedFilePath = SimpleLicenseVerifier.GetValidatedFilePath(PROJECT_NAME);
                    bool hasValidatedBefore = File.Exists(validatedFilePath);

                    if (!noExpire && !skipAgreement)
                    {
                        if (!hasValidatedBefore)
                        {
                            if (ShowLinkedDialog(window, "Agreement", firstUseMessagePrefix, GITHUB_URL, agreementMessageSuffix, MessageBoxButton.YesNo) == MessageBoxResult.No)
                            {
                                return;
                            }

                            try
                            {
                                File.WriteAllText(validatedFilePath, DateTime.Now.ToString(CultureInfo.InvariantCulture));
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine("Error saving validation status to AppData: " + ex.Message);
                            }

                            isValidatedForUser = true;
                        }
                        else
                        {
                            if (ShowLinkedDialog(window, "Agreement", repeatUseMessagePrefix, GITHUB_URL, agreementMessageSuffix, MessageBoxButton.YesNo) == MessageBoxResult.No)
                            {
                                return;
                            }

                            isValidatedForUser = true;
                        }
                    }
                    else
                    {
                        isValidatedForUser = IsValidated || hasValidatedBefore;
                    }

                    IsValidated = isValidatedForUser;
                }

                if (context.Patient == null)
                {
                    MessageBox.Show("No patient currently loaded.");
                    return;
                }

                if (context.IonPlanSetup == null)
                {
                    MessageBox.Show("No proton plan currently loaded.");
                    return;
                }

                ProtonSpotSequencer sequencer = new ProtonSpotSequencer(context, window, environment, IsValidated);
                window.Content = sequencer;
                window.Title = "MAAS - ProtonSpotSequencer" + (IsValidated ? string.Empty : " NOT VALIDATED FOR CLINICAL USE");
                window.Width = 1400;
                window.Height = 900;

                RoutedEventHandler loadedHandler = null;
                loadedHandler = (sender, e) =>
                {
                    window.WindowState = WindowState.Normal;
                    window.Activate();
                    window.Loaded -= loadedHandler;
                };
                window.Loaded += loadedHandler;
            }
            catch (Exception ex)
            {
                string errorText = ex.InnerException != null
                    ? ex.Message + "\nInner: " + ex.InnerException.Message
                    : ex.Message;
                MessageBox.Show("Error: " + errorText);
            }
        }
    }
}