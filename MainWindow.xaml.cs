using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using System.Windows.Media.Imaging;

// EXIF
using MetadataExtractor;
using MetadataExtractor.Formats.Exif;

namespace RSS_Image_Interoperability_Tool
{
    public partial class MainWindow : Window
    {
        private List<CameraRow> _csvRows = new();
        // filename -> full path
        private readonly Dictionary<string, string> _imagePaths = new(StringComparer.OrdinalIgnoreCase);
        // cache of EXIF/file props keyed by full path
        private readonly Dictionary<string, Dictionary<string, object>> _exifCache =
            new(StringComparer.OrdinalIgnoreCase);

        // UI list + preview
        public class ImageItem { public string Name { get; set; } = ""; public string FullPath { get; set; } = ""; }
        private readonly ObservableCollection<ImageItem> _imageItems = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        // ===== UI actions =====
        private void BrowseCsv_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "CSV files|*.csv|All files|*.*" };
            if (dlg.ShowDialog() == true)
            {
                CsvPathBox.Text = dlg.FileName;
                LoadCsv(dlg.FileName);
                UpdateStatsAndPreview();
            }
        }

        private void AddImageFiles_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Images|*.jpg;*.jpeg;*.png;*.webp|All files|*.*",
                Multiselect = true
            };
            if (dlg.ShowDialog() == true)
            {
                foreach (var p in dlg.FileNames)
                    _imagePaths[System.IO.Path.GetFileName(p)] = p;

                ImagesPathBox.Text = $"{_imagePaths.Count} images (files)";
                UpdateStatsAndPreview();
            }
        }

        private void AddImageFolder_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".jpg", ".jpeg", ".png", ".webp" };
                var files = System.IO.Directory.EnumerateFiles(dlg.SelectedPath, "*.*", SearchOption.AllDirectories)
                               .Where(p => exts.Contains(System.IO.Path.GetExtension(p)));
                foreach (var p in files)
                    _imagePaths[System.IO.Path.GetFileName(p)] = p;

                ImagesPathBox.Text = $"{_imagePaths.Count} images (folder)";
                UpdateStatsAndPreview();
            }
        }

        private void ClearImages_Click(object sender, RoutedEventArgs e)
        {
            _imagePaths.Clear();
            ImagesPathBox.Text = "";
            UpdateStatsAndPreview();
        }

        private void ChooseOutput_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "GeoJSON|*.geojson|JSON|*.json|All files|*.*",
                FileName = string.IsNullOrWhiteSpace(SingleFileNameBox.Text) ? "cameras.geojson" : SingleFileNameBox.Text
            };
            if (dlg.ShowDialog() == true) OutPathBox.Text = dlg.FileName;
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (_csvRows.Count == 0) { MessageBox.Show("Load a RealityCapture camera CSV first."); return; }
            if (string.IsNullOrWhiteSpace(OutPathBox.Text)) { MessageBox.Show("Choose an output file path."); return; }

            bool localMode = ModeLocalRadio.IsChecked == true;

            double originLon = 0, originLat = 0, originH = 0;
            if (localMode)
            {
                if (!double.TryParse(OriginLonBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out originLon) ||
                    !double.TryParse(OriginLatBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out originLat) ||
                    !double.TryParse(OriginHgtBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out originH))
                {
                    MessageBox.Show("Enter valid project origin lon/lat/height.");
                    return;
                }
            }

            // Filter rows if needed
            var rows = _csvRows;
            if (OnlyMatchedCheck.IsChecked == true)
                rows = _csvRows.Where(r => r.HasImage).ToList();

            if (rows.Count == 0)
            {
                MessageBox.Show("No rows to export with current filters.");
                return;
            }

            string baseImg = ImagesBaseUrlBox.Text?.Trim() ?? "";
            string baseThumb = ThumbsBaseUrlBox.Text?.Trim() ?? "";
            string pinIcon = PinIconUrlBox.Text?.Trim() ?? "/icons/camera-pin.png";

            var features = new List<object>(rows.Count);
            foreach (var row in rows)
            {
                double lon, lat, h;
                if (localMode)
                    (lon, lat, h) = GeoMath.LocalEnuToLonLatH(row.X, row.Y, row.Alt, originLon, originLat, originH);
                else
                    (lon, lat, h) = (row.X, row.Y, row.Alt);

                string imgUrl = BuildImageUrl(row.Name, baseImg); // file:/// if enabled & path known
                string thumbUrl = CombineUrl(baseThumb, row.Name);

                Dictionary<string, object>? exifProps = null;
                if (IncludeExifCheck.IsChecked == true && row.HasImage)
                    exifProps = GetExifAndFilePropsFor(row.Name);

                features.Add(new
                {
                    type = "Feature",
                    geometry = new { type = "Point", coordinates = new[] { lon, lat, h } },
                    properties = new
                    {
                        id = row.Name,
                        heading = row.Heading,
                        pitch = row.Pitch,
                        roll = row.Roll,
                        imageUrl = imgUrl,
                        thumbUrl = thumbUrl,
                        pinIcon = pinIcon,
                        hasImage = row.HasImage,
                        exif = exifProps
                    }
                });
            }

            var fc = new { type = "FeatureCollection", features };
            File.WriteAllText(OutPathBox.Text, JsonSerializer.Serialize(fc, new JsonSerializerOptions { WriteIndented = true }));
            MessageBox.Show($"Exported GeoJSON:\n{OutPathBox.Text}");
        }

        private void CoordModeChanged(object sender, RoutedEventArgs e)
        {
            OriginPanel.IsEnabled = ModeLocalRadio.IsChecked == true;
            UpdateStatsAndPreview();
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);

            var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".jpg", ".jpeg", ".png", ".webp" };

            foreach (var p in paths)
            {
                if (System.IO.Directory.Exists(p))
                {
                    foreach (var f in System.IO.Directory.EnumerateFiles(p, "*.*", SearchOption.AllDirectories)
                                                         .Where(f => exts.Contains(System.IO.Path.GetExtension(f))))
                        _imagePaths[System.IO.Path.GetFileName(f)] = f;
                }
                else if (File.Exists(p))
                {
                    var ext = System.IO.Path.GetExtension(p);
                    if (ext.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        CsvPathBox.Text = p;
                        LoadCsv(p);
                    }
                    else if (exts.Contains(ext))
                    {
                        _imagePaths[System.IO.Path.GetFileName(p)] = p;
                    }
                }
            }
            ImagesPathBox.Text = _imagePaths.Count > 0 ? $"{_imagePaths.Count} images" : ImagesPathBox.Text;
            UpdateStatsAndPreview();
        }

        private void ImagesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var item = ImagesList.SelectedItem as ImageItem;
            PreviewImage.Source = null;
            PreviewNameText.Text = "";
            PreviewDimText.Text = "";
            PreviewPathText.Text = "";

            if (item == null || !File.Exists(item.FullPath)) return;

            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new Uri(item.FullPath);
                bmp.EndInit();
                bmp.Freeze();

                PreviewImage.Source = bmp;
                PreviewNameText.Text = item.Name;
                PreviewDimText.Text = $"{bmp.PixelWidth} × {bmp.PixelHeight}";
                PreviewPathText.Text = item.FullPath;
            }
            catch
            {
                // ignore preview errors
            }
        }

        // ===== helpers =====
        private void LoadCsv(string path)
        {
            _csvRows.Clear();
            using var sr = new StreamReader(path);
            string? header = sr.ReadLine();
            if (header == null) { MessageBox.Show("Empty CSV."); return; }

            var cols = header.Split(',');
            int idxName = Array.FindIndex(cols, c => c.Trim().Equals("#name", StringComparison.OrdinalIgnoreCase) || c.Trim().Equals("name", StringComparison.OrdinalIgnoreCase));
            int idxX = Array.FindIndex(cols, c => c.Trim().Equals("x", StringComparison.OrdinalIgnoreCase));
            int idxY = Array.FindIndex(cols, c => c.Trim().Equals("y", StringComparison.OrdinalIgnoreCase));
            int idxAlt = Array.FindIndex(cols, c => c.Trim().Equals("alt", StringComparison.OrdinalIgnoreCase));
            int idxHeading = Array.FindIndex(cols, c => c.Trim().Equals("heading", StringComparison.OrdinalIgnoreCase));
            int idxPitch = Array.FindIndex(cols, c => c.Trim().Equals("pitch", StringComparison.OrdinalIgnoreCase));
            int idxRoll = Array.FindIndex(cols, c => c.Trim().Equals("roll", StringComparison.OrdinalIgnoreCase));

            if (idxName < 0 || idxX < 0 || idxY < 0 || idxAlt < 0 || idxHeading < 0 || idxPitch < 0 || idxRoll < 0)
            {
                MessageBox.Show("CSV is missing required columns (#name, x, y, alt, heading, pitch, roll).");
                return;
            }

            string? line;
            while ((line = sr.ReadLine()) != null)
            {
                var parts = SplitCsvLine(line, cols.Length);
                if (parts.Length != cols.Length) continue;

                _csvRows.Add(new CameraRow
                {
                    Name = parts[idxName].Trim(),
                    X = ParseInv(parts[idxX]),
                    Y = ParseInv(parts[idxY]),
                    Alt = ParseInv(parts[idxAlt]),
                    Heading = ParseInv(parts[idxHeading]),
                    Pitch = ParseInv(parts[idxPitch]),
                    Roll = ParseInv(parts[idxRoll]),
                });
            }
        }

        private static double ParseInv(string s) =>
            double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0;

        private static string[] SplitCsvLine(string line, int expected)
        {
            var list = new List<string>(expected);
            bool inQuotes = false;
            var cur = new System.Text.StringBuilder();
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '\"') { cur.Append('\"'); i++; }
                    else inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    list.Add(cur.ToString());
                    cur.Clear();
                }
                else cur.Append(c);
            }
            list.Add(cur.ToString());
            return list.ToArray();
        }

        private void UpdateStatsAndPreview()
        {
            int matched = 0;
            foreach (var r in _csvRows)
            {
                r.HasImage = _imagePaths.ContainsKey(r.Name);
                if (r.HasImage) matched++;
            }

            CsvCountText.Text = _csvRows.Count.ToString();
            ImgCountText.Text = _imagePaths.Count.ToString();
            MatchCountText.Text = matched.ToString();
            MissingCountText.Text = (_csvRows.Count - matched).ToString();

            CoordSummaryText.Text = ModeLocalRadio.IsChecked == true
                ? "Coordinate mode: LOCAL ENU (meters) → will convert using Project Origin."
                : "Coordinate mode: WGS84 (x=lon, y=lat, alt=height) → will export as-is.";

            PreviewList.ItemsSource = null;
            PreviewList.ItemsSource = _csvRows;

            // unmatched rows highlighted
            PreviewList.ItemContainerStyle = new Style(typeof(ListViewItem))
            {
                Setters = { new Setter(ListViewItem.BackgroundProperty, Brushes.Transparent) },
                Triggers = {
                    new DataTrigger {
                        Binding = new System.Windows.Data.Binding("HasImage"),
                        Value = false,
                        Setters = { new Setter(ListViewItem.BackgroundProperty,
                                   new SolidColorBrush(Color.FromRgb(255,244,244))) }
                    }
                }
            };

            // rebuild image list
            _imageItems.Clear();
            foreach (var kvp in _imagePaths.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                _imageItems.Add(new ImageItem { Name = kvp.Key, FullPath = kvp.Value });
            ImagesList.ItemsSource = _imageItems;
        }

        private static string CombineUrl(string baseUrl, string tail)
        {
            if (string.IsNullOrEmpty(baseUrl)) return tail;
            if (!baseUrl.EndsWith("/")) baseUrl += "/";
            return baseUrl + tail;
        }

        private string BuildImageUrl(string fileName, string baseUrl)
        {
            // If user chose file:/// URLs and we know the full path, use it
            if (UseFileSchemeCheck.IsChecked == true &&
                _imagePaths.TryGetValue(fileName, out var fullPath))
            {
                try { return new Uri(fullPath).AbsoluteUri; }
                catch { /* fall back */ }
            }
            return CombineUrl(baseUrl, fileName);
        }

        private Dictionary<string, object> GetExifAndFilePropsFor(string fileName)
        {
            if (!_imagePaths.TryGetValue(fileName, out var fullPath) || !File.Exists(fullPath))
                return new Dictionary<string, object>();

            if (_exifCache.TryGetValue(fullPath, out var cached))
                return cached;

            var props = new Dictionary<string, object>();

            // File properties
            var fi = new FileInfo(fullPath);
            props["fileName"] = fi.Name;
            props["fileSizeBytes"] = fi.Length;
            props["createdUtc"] = fi.CreationTimeUtc;
            props["modifiedUtc"] = fi.LastWriteTimeUtc;

            // EXIF
            try
            {
                var dirs = ImageMetadataReader.ReadMetadata(fullPath);

                var ifd0 = dirs.OfType<ExifIfd0Directory>().FirstOrDefault();
                var subIfd = dirs.OfType<ExifSubIfdDirectory>().FirstOrDefault();
                var ifd = dirs.OfType<MetadataExtractor.Formats.Exif.ExifImageDirectory>().FirstOrDefault();
                var gps = dirs.OfType<GpsDirectory>().FirstOrDefault();

                int? width = ifd?.GetInt32(ExifDirectoryBase.TagImageWidth)
                          ?? subIfd?.GetInt32(ExifDirectoryBase.TagExifImageWidth);
                int? height = ifd?.GetInt32(ExifDirectoryBase.TagImageHeight)
                           ?? subIfd?.GetInt32(ExifDirectoryBase.TagExifImageHeight);

                if (width.HasValue) props["imageWidth"] = width.Value;
                if (height.HasValue) props["imageHeight"] = height.Value;

                string? make = ifd0?.GetDescription(ExifDirectoryBase.TagMake);
                string? model = ifd0?.GetDescription(ExifDirectoryBase.TagModel);
                if (!string.IsNullOrEmpty(make)) props["cameraMake"] = make;
                if (!string.IsNullOrEmpty(model)) props["cameraModel"] = model;

                var fnum = subIfd?.GetRational(ExifDirectoryBase.TagFNumber);
                if (fnum != null) props["fNumber"] = Math.Round(fnum.Value.ToDouble(), 3);

                var flen = subIfd?.GetRational(ExifDirectoryBase.TagFocalLength);
                if (flen != null) props["focalLengthMm"] = Math.Round(flen.Value.ToDouble(), 3);

                var exp = subIfd?.GetRational(ExifDirectoryBase.TagExposureTime);
                if (exp != null)
                {
                    var seconds = exp.Value.ToDouble();
                    props["exposureSeconds"] = seconds;
                    if (seconds > 0) props["exposureDisplay"] = seconds >= 1 ? $"{seconds:0.###} s" : $"1/{Math.Round(1 / seconds)} s";
                }

                var iso = subIfd?.GetInt32(ExifDirectoryBase.TagIsoEquivalent);
                if (iso.HasValue) props["iso"] = iso.Value;

                var dateTaken = subIfd?.GetDateTime(ExifDirectoryBase.TagDateTimeOriginal);
                if (dateTaken.HasValue) props["dateTimeOriginal"] = dateTaken.Value.ToUniversalTime();

                if (gps != null)
                {
                    var loc = gps.GetGeoLocation();
                    if (loc != null && !double.IsNaN(loc.Latitude) && !double.IsNaN(loc.Longitude))
                    {
                        props["gpsLon"] = loc.Longitude;
                        props["gpsLat"] = loc.Latitude;
                        if (gps.TryGetRational(GpsDirectory.TagAltitude, out var altRat))
                            props["gpsAlt"] = altRat.ToDouble();
                    }
                }
            }
            catch { /* ignore */ }

            _exifCache[fullPath] = props;
            return props;
        }
    }

    public class CameraRow
    {
        public string Name { get; set; } = "";
        public double X { get; set; }     // lon or local Easting (m)
        public double Y { get; set; }     // lat or local Northing (m)
        public double Alt { get; set; }   // height (m)
        public double Heading { get; set; } // deg
        public double Pitch { get; set; }   // deg
        public double Roll { get; set; }    // deg
        public bool HasImage { get; set; }
    }

    internal static class GeoMath
    {
        const double a = 6378137.0;                 // WGS84 semimajor
        const double f = 1.0 / 298.257223563;
        const double b = a * (1 - f);
        const double e2 = 1 - (b * b) / (a * a);

        public static (double lonDeg, double latDeg, double h) LocalEnuToLonLatH(
            double east, double north, double up,
            double originLonDeg, double originLatDeg, double originH)
        {
            var (x0, y0, z0) = LlaToEcef(originLonDeg, originLatDeg, originH);

            double lon = DegToRad(originLonDeg);
            double lat = DegToRad(originLatDeg);
            double sinLon = Math.Sin(lon), cosLon = Math.Cos(lon);
            double sinLat = Math.Sin(lat), cosLat = Math.Cos(lat);

            // ENU→ECEF rotation
            double r11 = -sinLon; double r12 = cosLon; double r13 = 0;
            double r21 = -sinLat * cosLon; double r22 = -sinLat * sinLon; double r23 = cosLat;
            double r31 = cosLat * cosLon; double r32 = cosLat * sinLon; double r33 = sinLat;

            double dx = r11 * east + r21 * north + r31 * up;
            double dy = r12 * east + r22 * north + r32 * up;
            double dz = r13 * east + r23 * north + r33 * up;

            double x = x0 + dx, y = y0 + dy, z = z0 + dz;
            return EcefToLla(x, y, z);
        }

        public static (double x, double y, double z) LlaToEcef(double lonDeg, double latDeg, double h)
        {
            double lon = DegToRad(lonDeg);
            double lat = DegToRad(latDeg);
            double cosLat = Math.Cos(lat);
            double sinLat = Math.Sin(lat);
            double cosLon = Math.Cos(lon);
            double sinLon = Math.Sin(lon);

            double N = a / Math.Sqrt(1 - e2 * sinLat * sinLat);
            double x = (N + h) * cosLat * cosLon;
            double y = (N + h) * cosLat * sinLon;
            double z = (N * (1 - e2) + h) * sinLat;
            return (x, y, z);
        }

        public static (double lonDeg, double latDeg, double h) EcefToLla(double x, double y, double z)
        {
            double lon = Math.Atan2(y, x);
            double p = Math.Sqrt(x * x + y * y);
            double theta = Math.Atan2(z * a, p * b);
            double sinTheta = Math.Sin(theta);
            double cosTheta = Math.Cos(theta);
            double e2p = (a * a - b * b) / (b * b);

            double lat = Math.Atan2(z + e2p * b * Math.Pow(sinTheta, 3),
                                    p - e2 * a * Math.Pow(cosTheta, 3));

            double N = a / Math.Sqrt(1 - e2 * Math.Sin(lat) * Math.Sin(lat));
            double h = p / Math.Cos(lat) - N;

            return (RadToDeg(lon), RadToDeg(lat), h);
        }

        private static double DegToRad(double d) => d * Math.PI / 180.0;
        private static double RadToDeg(double r) => r * 180.0 / Math.PI;
    }
}
