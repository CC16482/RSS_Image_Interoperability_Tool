using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows.Media;

// EXIF
using MetadataExtractor;
using MetadataExtractor.Formats.Exif;

namespace RSS_Image_Interoperability_Tool
{
    public partial class MainWindow : Window
    {
        private static readonly JsonSerializerOptions s_jsonOpts = new() { WriteIndented = true };

        private readonly List<CameraRow> _csvRows = new();
        private bool _csvAnglesAreOpk = false;

        // filename -> full path
        private readonly Dictionary<string, string> _imagePaths =
            new(StringComparer.OrdinalIgnoreCase);

        // EXIF/file-props cache keyed by full path
        private readonly Dictionary<string, Dictionary<string, object>> _exifCache =
            new(StringComparer.OrdinalIgnoreCase);

        public class ImageItem
        {
            public string Name { get; set; } = "";
            public string FullPath { get; set; } = "";
            public string Modified { get; set; } = "";
            public long SizeKb { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
        }
        private readonly ObservableCollection<ImageItem> _imageItems = new();

        public MainWindow()
        {
            InitializeComponent();

            Loaded += (_, __) =>
            {
                // Toggle-like radio checkboxes
                ModeWgs84Radio.Checked += (_, __2) => { ModeLocalRadio.IsChecked = false; ApplyCoordMode(); };
                ModeLocalRadio.Checked += (_, __2) => { ModeWgs84Radio.IsChecked = false; ApplyCoordMode(); };
                ModeWgs84Radio.Unchecked += (_, __2) => ApplyCoordMode();
                ModeLocalRadio.Unchecked += (_, __2) => ApplyCoordMode();

                ApplyCoordMode();

                // Ensure context menus exist (in case XAML didn't define them)
                EnsureImagesContextMenu();
                EnsureCsvContextMenu();

                // Keep angle summary in sync with toggle
                if (ConvertOpkCheck != null)
                {
                    ConvertOpkCheck.Checked += (_, __2) => UpdateStatsAndPreview();
                    ConvertOpkCheck.Unchecked += (_, __2) => UpdateStatsAndPreview();
                }
                if (IncludeHprCheck != null)
                {
                    IncludeHprCheck.Checked += (_, __2) => UpdateStatsAndPreview();
                    IncludeHprCheck.Unchecked += (_, __2) => UpdateStatsAndPreview();
                }
            };
        }

        private void ApplyCoordMode()
        {
            OriginPanel.IsEnabled = ModeLocalRadio.IsChecked == true;
            if (IsLoaded) UpdateStatsAndPreview();
        }

        /* ======================= IMPORTS ======================= */

        private void BrowseCsv_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "CSV files|*.csv|All files|*.*" };
            if (dlg.ShowDialog() == true)
            {
                CsvPathBox.Text = dlg.FileName;
                LoadCsv2(dlg.FileName);
                UpdateStatsAndPreview();
            }
        }

        private void ClearCsv_Click(object sender, RoutedEventArgs e)
        {
            _csvRows.Clear();
            CsvPathBox.Text = string.Empty;
            CsvRowsGrid.ItemsSource = null;
            UpdateStatsAndPreview();
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
                var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { ".jpg", ".jpeg", ".png", ".webp" };

                foreach (var f in System.IO.Directory
                             .EnumerateFiles(dlg.SelectedPath, "*.*", SearchOption.AllDirectories)
                             .Where(p => exts.Contains(System.IO.Path.GetExtension(p))))
                {
                    _imagePaths[System.IO.Path.GetFileName(f)] = f;
                }

                ImagesPathBox.Text = $"{_imagePaths.Count} images (folder)";
                UpdateStatsAndPreview();
            }
        }

        private void ClearImages_Click(object sender, RoutedEventArgs e)
        {
            _imagePaths.Clear();
            ImagesPathBox.Text = "";
            _imageItems.Clear();
            PreviewImage.Source = null;
            PreviewNameText.Text = "";
            PreviewDimText.Text = "";
            ExifDetailsPanel.ItemsSource = null;
            UpdateStatsAndPreview();
        }

        /* ======================= EXPORT ======================= */

        private void ChooseOutput_Click(object sender, RoutedEventArgs e)
        {
            var suggested = string.IsNullOrWhiteSpace(SingleFileNameBox.Text)
                ? "cameras.geojson"
                : SingleFileNameBox.Text;

            var dlg = new SaveFileDialog
            {
                Filter = "GeoJSON|*.geojson|JSON|*.json|All files|*.*",
                FileName = suggested
            };
            if (dlg.ShowDialog() == true) OutPathBox.Text = dlg.FileName;
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (_csvRows.Count == 0)
            {
                MessageBox.Show("Load a RealityCapture camera CSV first.");
                return;
            }
            if (string.IsNullOrWhiteSpace(OutPathBox.Text))
            {
                MessageBox.Show("Choose an output file path.");
                return;
            }

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

                // Choose angles to export
                double outHeading = row.Heading;
                double outPitch = row.Pitch;
                double outRoll = row.Roll;
                double? qx = null, qy = null, qz = null, qw = null; // ECEF quaternion

                // If CSV angles are OPK (RealityCapture), always compute the robust ECEF quaternion
                // from OPK + position (lon/lat). Optionally convert HPR if the toggle is on.
                bool haveRadHpr = false;
                if (_csvAnglesAreOpk)
                {
                    if (ConvertOpkCheck?.IsChecked == true)
                    {
                        (outHeading, outPitch, outRoll) = GeoMath.OpkDegToCesiumHprRad(row.Omega, row.Phi, row.Kappa, lon, lat);
                        haveRadHpr = true;
                    }

                    try
                    {
                        var q = GeoMath.OpkDegToCesiumQuaternion(row.Omega, row.Phi, row.Kappa, lon, lat);
                        qx = q.qx; qy = q.qy; qz = q.qz; qw = q.qw;
                    }
                    catch { /* fall back silently if quaternion calc fails */ }
                }

                string imgUrl = BuildImageUrl(row.Name, baseImg);
                string thumbUrl = CombineUrl(baseThumb, row.Name);

                Dictionary<string, object>? exifProps = null;
                if (IncludeExifCheck.IsChecked == true && row.HasImage)
                    exifProps = GetExifAndFilePropsFor(row.Name);

                var props = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["id"] = row.Name,
                    ["imageUrl"] = imgUrl,
                    ["thumbUrl"] = thumbUrl,
                    ["pinIcon"] = pinIcon,
                    ["hasImage"] = row.HasImage,
                    ["exif"] = exifProps,
                };

                if (_csvAnglesAreOpk)
                {
                    props["opk"] = new { omegaDeg = row.Omega, phiDeg = row.Phi, kappaDeg = row.Kappa };
                }

                if (qx.HasValue && qy.HasValue && qz.HasValue && qw.HasValue)
                {
                    props["quaternion"] = new[] { qx.Value, qy.Value, qz.Value, qw.Value };
                    props["quaternionFrame"] = "ECEF";
                }

                bool includeHpr = IncludeHprCheck?.IsChecked == true;
                if (includeHpr)
                {
                    if (haveRadHpr)
                    {
                        // Best: export Cesium-convention values in radians with explicit tags
                        props["cesiumHeading"] = outHeading; // radians CCW from local East
                        props["pitch"] = outPitch;           // radians
                        props["roll"] = outRoll;             // radians
                        props["headingUnit"] = "rad";
                        props["pitchUnit"] = "rad";
                        props["rollUnit"] = "rad";
                        props["headingConvention"] = "east-radians";
                    }
                    else
                    {
                        // As-is CSV HPR; tag explicitly as degrees and north-based heading (most common)
                        var csvConv = CsvHprConventionCombo?.SelectedValue as string;
                        if (string.IsNullOrWhiteSpace(csvConv)) csvConv = "north-degrees";
                        props["heading"] = outHeading;       // degrees (assumed)
                        props["pitch"] = outPitch;           // degrees (assumed)
                        props["roll"] = outRoll;             // degrees (assumed)
                        props["headingUnit"] = "deg";
                        props["pitchUnit"] = "deg";
                        props["rollUnit"] = "deg";
                        props["headingConvention"] = csvConv; // "north-degrees" or "east-degrees"
                    }
                }

                features.Add(new
                {
                    type = "Feature",
                    geometry = new { type = "Point", coordinates = new[] { lon, lat, h } },
                    properties = props
                });
            }

            var fc = new { type = "FeatureCollection", features };
            File.WriteAllText(OutPathBox.Text, JsonSerializer.Serialize(fc, s_jsonOpts));
            MessageBox.Show($"Exported GeoJSON:\n{OutPathBox.Text}");
        }

        /* ======================= DRAG & DROP ======================= */

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);

            var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { ".jpg", ".jpeg", ".png", ".webp" };

            foreach (var p in paths)
            {
                if (System.IO.Directory.Exists(p))
                {
                    foreach (var f in System.IO.Directory
                                 .EnumerateFiles(p, "*.*", SearchOption.AllDirectories)
                                 .Where(f => exts.Contains(System.IO.Path.GetExtension(f))))
                        _imagePaths[System.IO.Path.GetFileName(f)] = f;
                }
                else if (File.Exists(p))
                {
                    var ext = System.IO.Path.GetExtension(p);
                    if (ext.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        CsvPathBox.Text = p;
                        LoadCsv2(p);
                    }
                    else if (exts.Contains(ext))
                    {
                        _imagePaths[System.IO.Path.GetFileName(p)] = p;
                    }
                }
            }

            ImagesPathBox.Text = _imagePaths.Count > 0
                ? $"{_imagePaths.Count} images"
                : ImagesPathBox.Text;

            UpdateStatsAndPreview();
        }

        /* ======================= CSV LOAD ======================= */

        private void LoadCsv2(string path)
        {
            try
            {
                _csvRows.Clear();

                using var sr = new StreamReader(path);

                // 1) Find header row (skip comment/count lines)
                string? header = null;
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    if (line.StartsWith("#") && !(line.Contains(',') || line.Contains('\t'))) continue;
                    header = line;
                    break;
                }
                if (header is null)
                {
                    MessageBox.Show("CSV appears empty.");
                    return;
                }

                // 2) Detect delimiter
                char delim = header.Contains(',') ? ',' : (header.Contains('\t') ? '\t' : ',');
                var cols = header.Split(delim);
                string Norm(string s) => s.Trim().TrimStart('#').ToLowerInvariant();

                int idxName = Array.FindIndex(cols, c => Norm(c) is "name" or "#name");
                int idxX = Array.FindIndex(cols, c => new[] { "x", "lon", "longitude" }.Contains(Norm(c)));
                int idxY = Array.FindIndex(cols, c => new[] { "y", "lat", "latitude" }.Contains(Norm(c)));
                int idxAlt = Array.FindIndex(cols, c => new[] { "alt", "z", "height", "h" }.Contains(Norm(c)));
                int idxHeading = Array.FindIndex(cols, c => new[] { "heading", "yaw", "kappa" }.Contains(Norm(c)));
                int idxPitch = Array.FindIndex(cols, c => new[] { "pitch", "phi" }.Contains(Norm(c)));
                int idxRoll = Array.FindIndex(cols, c => new[] { "roll", "omega" }.Contains(Norm(c)));

                if (idxName < 0 || idxX < 0 || idxY < 0 || idxAlt < 0 || idxHeading < 0 || idxPitch < 0 || idxRoll < 0)
                {
                    MessageBox.Show(
                        "CSV is missing required columns.\n\nAccepted names:\n" +
                        "name — #name/name\nx — x/lon/longitude\ny — y/lat/latitude\nalt — alt/z/height/h\n" +
                        "heading — heading/yaw/kappa\npitch — pitch/phi\nroll — roll/omega");
                    return;
                }

                // 3) Mark whether angles are OPK
                _csvAnglesAreOpk = false;
                try
                {
                    var hName = Norm(cols[idxHeading]);
                    var pName = Norm(cols[idxPitch]);
                    var rName = Norm(cols[idxRoll]);
                    _csvAnglesAreOpk = (hName == "kappa" && pName == "phi" && rName == "omega");
                }
                catch { _csvAnglesAreOpk = false; }

                // 4) Data rows
                string? line2;
                while ((line2 = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line2)) continue;
                    if (line2.StartsWith("#")) continue;

                    var parts = SplitSeparatedLine(line2, delim, cols.Length);
                    if (parts.Length < cols.Length) continue;

                    var row = new CameraRow
                    {
                        Name = parts[idxName].Trim(),
                        X = ParseInv(parts[idxX]),
                        Y = ParseInv(parts[idxY]),
                        Alt = ParseInv(parts[idxAlt]),
                        Heading = ParseInv(parts[idxHeading]),
                        Pitch = ParseInv(parts[idxPitch]),
                        Roll = ParseInv(parts[idxRoll])
                    };

                    if (_csvAnglesAreOpk)
                    {
                        row.Omega = ParseInv(parts[idxRoll]);   // omega
                        row.Phi   = ParseInv(parts[idxPitch]);  // phi
                        row.Kappa = ParseInv(parts[idxHeading]); // kappa
                    }

                    _csvRows.Add(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to read CSV: {ex.Message}");
            }
        }

        private static string[] SplitSeparatedLine(string line, char delimiter, int expected)
        {
            if (delimiter == ',') return SplitCsvLine(line, expected);
            if (delimiter == '\t') return line.Split('\t');
            return line.Split(delimiter);
        }
        private void LoadCsv(string path)
        {
            _csvRows.Clear();

            using var sr = new StreamReader(path);

            // 1) Find header row
            string? header = null;
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;
                if (line.StartsWith("#") && !line.Contains(',')) continue;
                header = line;
                break;
            }
            if (header is null)
            {
                MessageBox.Show("CSV appears empty.");
                return;
            }

            var cols = header.Split(',');
            string Norm(string s) => s.Trim().TrimStart('#').ToLowerInvariant();

            int idxName = Array.FindIndex(cols, c => Norm(c) is "name" or "#name");
            int idxX = Array.FindIndex(cols, c => new[] { "x", "lon", "longitude" }.Contains(Norm(c)));
            int idxY = Array.FindIndex(cols, c => new[] { "y", "lat", "latitude" }.Contains(Norm(c)));
            int idxAlt = Array.FindIndex(cols, c => new[] { "alt", "z", "height", "h" }.Contains(Norm(c)));
            int idxHeading = Array.FindIndex(cols, c => new[] { "heading", "yaw", "kappa" }.Contains(Norm(c)));
            int idxPitch = Array.FindIndex(cols, c => new[] { "pitch", "phi" }.Contains(Norm(c)));
            int idxRoll = Array.FindIndex(cols, c => new[] { "roll", "omega" }.Contains(Norm(c)));

            if (idxName < 0 || idxX < 0 || idxY < 0 || idxAlt < 0 || idxHeading < 0 || idxPitch < 0 || idxRoll < 0)
            {
                MessageBox.Show(
                    "CSV is missing required columns.\n\nAccepted names:\n" +
                    "name → #name/name\nx → x/lon/longitude\ny → y/lat/latitude\nalt → alt/z/height/h\n" +
                    "heading → heading/yaw/kappa\npitch → pitch/phi\nroll → roll/omega");
                return;
            }

            // Determine if rotations are OPK (omega/phi/kappa from RealityCapture)
            _csvAnglesAreOpk = false;
            try
            {
                if (idxHeading >= 0 && idxPitch >= 0 && idxRoll >= 0)
                {
                    var hName = Norm(cols[idxHeading]);
                    var pName = Norm(cols[idxPitch]);
                    var rName = Norm(cols[idxRoll]);
                    _csvAnglesAreOpk = (hName == "kappa" && pName == "phi" && rName == "omega");
                }
            }
            catch { _csvAnglesAreOpk = false; }

            // 2) data rows
            string? line2;
            while ((line2 = sr.ReadLine()) != null)
            {
                if (string.IsNullOrWhiteSpace(line2)) continue;
                if (line2.StartsWith("#")) continue;

                var parts = SplitCsvLine(line2, cols.Length);
                if (parts.Length < cols.Length) continue;

                var row = new CameraRow
                {
                    Name = parts[idxName].Trim(),
                    X = ParseInv(parts[idxX]),
                    Y = ParseInv(parts[idxY]),
                    Alt = ParseInv(parts[idxAlt]),
                    // store raw CSV angles as read (units from CSV)
                    Heading = ParseInv(parts[idxHeading]),
                    Pitch = ParseInv(parts[idxPitch]),
                    Roll = ParseInv(parts[idxRoll])
                };

                if (_csvAnglesAreOpk)
                {
                    row.Omega = ParseInv(parts[idxRoll]);   // omega
                    row.Phi   = ParseInv(parts[idxPitch]);  // phi
                    row.Kappa = ParseInv(parts[idxHeading]); // kappa
                }

                _csvRows.Add(row);
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

        /* ======================= STATS / BIND / PREVIEW ======================= */

        private void UpdateStatsAndPreview()
        {
            if (ModeWgs84Radio.IsChecked == true && ModeLocalRadio.IsChecked == true)
                ModeLocalRadio.IsChecked = false;

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
                : "Coordinate mode: WGS84 (x=lon, y=lat, alt=height) — will export as-is.";

            // Bind CSV rows (row coloring handled by XAML RowStyle triggers)
            CsvRowsGrid.ItemsSource = null;
            CsvRowsGrid.ItemsSource = _csvRows;

            // Angles summary
            if (AnglesSummaryText != null)
            {
                if (_csvAnglesAreOpk)
                {
                    bool convert = ConvertOpkCheck?.IsChecked == true;
                    AnglesSummaryText.Text = convert
                        ? "Angles: CSV has omega/phi/kappa - converting to Cesium heading/pitch/roll (radians)."
                        : "Angles: CSV has omega/phi/kappa - exporting raw OPK values as heading/pitch/roll (no conversion).";
                }
                else
                {
                    AnglesSummaryText.Text = "Angles: using CSV heading/pitch/roll as-is.";
                }
            }

            // Enable/disable CSV HPR convention selector
            bool includeHpr = IncludeHprCheck?.IsChecked == true;
            bool opkConvert = (_csvAnglesAreOpk && (ConvertOpkCheck?.IsChecked == true));
            if (CsvHprConventionCombo != null)
                CsvHprConventionCombo.IsEnabled = includeHpr && !opkConvert;
            if (CsvHprConventionHint != null)
                CsvHprConventionHint.Text = opkConvert
                    ? "Not used when converting OPK → Cesium (using cesiumHeading)."
                    : "Used to tag CSV HPR as north-degrees or east-degrees.";

            // Rebuild image items off-UI thread
            Task.Run(() =>
            {
                var items = new List<ImageItem>(_imagePaths.Count);
                foreach (var kvp in _imagePaths.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                {
                    var fi = new FileInfo(kvp.Value);
                    int w = 0, h = 0;
                    try
                    {
                        var dirs = ImageMetadataReader.ReadMetadata(kvp.Value);
                        var sub = dirs.OfType<ExifSubIfdDirectory>().FirstOrDefault();
                        if (sub != null)
                        {
                            w = sub.GetInt32(ExifDirectoryBase.TagExifImageWidth);
                            h = sub.GetInt32(ExifDirectoryBase.TagExifImageHeight);
                        }
                    }
                    catch { /* leave 0x0 */ }

                    items.Add(new ImageItem
                    {
                        Name = kvp.Key,
                        FullPath = kvp.Value,
                        Modified = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"),
                        SizeKb = fi.Length / 1024,
                        Width = w,
                        Height = h
                    });
                }

                Dispatcher.Invoke(() =>
                {
                    _imageItems.Clear();
                    foreach (var it in items) _imageItems.Add(it);
                    ImagesList.ItemsSource = _imageItems;

                    // Keep the first row selected so preview shows something
                    if (_imageItems.Count > 0 && ImagesList.SelectedItem == null)
                        ImagesList.SelectedIndex = 0;
                });
            });
        }

        private static string CombineUrl(string baseUrl, string tail)
        {
            if (string.IsNullOrEmpty(baseUrl)) return tail;
            if (!baseUrl.EndsWith('/')) baseUrl += "/";
            return baseUrl + tail;
        }

        private string BuildImageUrl(string fileName, string baseUrl)
        {
            if (UseFileSchemeCheck.IsChecked == true &&
                _imagePaths.TryGetValue(fileName, out var fullPath))
            {
                try { return new Uri(fullPath).AbsoluteUri; }
                catch { /* fall back */ }
            }
            return CombineUrl(baseUrl, fileName);
        }

        private void ImagesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ImagesList.SelectedItem is not ImageItem item || !File.Exists(item.FullPath))
            {
                PreviewImage.Source = null;
                PreviewNameText.Text = "";
                PreviewDimText.Text = "";
                ExifDetailsPanel.ItemsSource = null;
                return;
            }

            try
            {
                // Fast, non-locking preview
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new Uri(item.FullPath);
                bmp.DecodePixelWidth = 2000; // tune as you like
                bmp.EndInit();
                bmp.Freeze();

                PreviewImage.Source = bmp;
                PreviewNameText.Text = item.Name;
                PreviewDimText.Text = $"{bmp.PixelWidth} × {bmp.PixelHeight}";

                var exif = GetExifAndFilePropsFor(item.Name);
                ExifDetailsPanel.ItemsSource =
                    BuildFixedExifList(item.FullPath, bmp.PixelWidth, bmp.PixelHeight, exif);
            }
            catch
            {
                PreviewImage.Source = null;
                ExifDetailsPanel.ItemsSource = null;
            }
        }

        /* ======================= EXIF ======================= */

        private Dictionary<string, object> GetExifAndFilePropsFor(string fileName)
        {
            if (!_imagePaths.TryGetValue(fileName, out var fullPath) || !File.Exists(fullPath))
                return new Dictionary<string, object>();

            if (_exifCache.TryGetValue(fullPath, out var cached))
                return cached;

            var props = new Dictionary<string, object>();

            // File props
            var fi = new FileInfo(fullPath);
            props["fileName"] = fi.Name;
            props["fileSizeBytes"] = fi.Length;
            props["createdUtc"] = fi.CreationTimeUtc;
            props["modifiedUtc"] = fi.LastWriteTimeUtc;

            try
            {
                var dirs = ImageMetadataReader.ReadMetadata(fullPath);

                var ifd0 = dirs.OfType<ExifIfd0Directory>().FirstOrDefault();
                var subIfd = dirs.OfType<ExifSubIfdDirectory>().FirstOrDefault();
                var ifd = dirs.OfType<ExifImageDirectory>().FirstOrDefault();
                var gps = dirs.OfType<GpsDirectory>().FirstOrDefault();

                int? width = ifd?.GetInt32(ExifDirectoryBase.TagImageWidth)
                          ?? subIfd?.GetInt32(ExifDirectoryBase.TagExifImageWidth);
                int? height = ifd?.GetInt32(ExifDirectoryBase.TagImageHeight)
                           ?? subIfd?.GetInt32(ExifDirectoryBase.TagExifImageHeight);
                if (width.HasValue) props["imageWidth"] = width.Value;
                if (height.HasValue) props["imageHeight"] = height.Value;

                string? make = ifd0?.GetDescription(ExifDirectoryBase.TagMake);
                string? model = ifd0?.GetDescription(ExifDirectoryBase.TagModel);
                if (!string.IsNullOrWhiteSpace(make)) props["cameraMake"] = make!;
                if (!string.IsNullOrWhiteSpace(model)) props["cameraModel"] = model!;

                string? lensModel = subIfd?.GetDescription(ExifDirectoryBase.TagLensModel);
                if (!string.IsNullOrWhiteSpace(lensModel)) props["lensModel"] = lensModel!;

                var fnum = subIfd?.GetRational(ExifDirectoryBase.TagFNumber);
                if (fnum != null) props["fNumber"] = Math.Round(fnum.Value.ToDouble(), 3);

                var flen = subIfd?.GetRational(ExifDirectoryBase.TagFocalLength);
                if (flen != null) props["focalLengthMm"] = Math.Round(flen.Value.ToDouble(), 3);

                var exp = subIfd?.GetRational(ExifDirectoryBase.TagExposureTime);
                if (exp != null)
                {
                    var seconds = exp.Value.ToDouble();
                    props["exposureSeconds"] = seconds;
                    if (seconds > 0)
                        props["exposureDisplay"] = seconds >= 1 ? $"{seconds:0.###} s" : $"1/{Math.Round(1 / seconds)} s";
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
            catch { /* ignore EXIF errors */ }

            _exifCache[fullPath] = props;
            return props;
        }

        private static List<KeyValuePair<string, string>> BuildFixedExifList(
            string fullPath,
            int width,
            int height,
            IReadOnlyDictionary<string, object> exif)
        {
            string GetS(string key)
            {
                if (exif.TryGetValue(key, out var v) && v != null)
                    return v switch
                    {
                        DateTime dt => dt.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                        double d => d.ToString("0.###", CultureInfo.InvariantCulture),
                        float f => f.ToString("0.###", CultureInfo.InvariantCulture),
                        _ => v.ToString() ?? "—"
                    };
                return "—";
            }

            var fi = new FileInfo(fullPath);
            string sizeMb = (fi.Length / 1024d / 1024d).ToString("0.00") + " MB";
            string dims = (width > 0 && height > 0) ? $"{width} × {height}" : "—";

            return new List<KeyValuePair<string, string>>
            {
                new("File name",         fi.Name),
                new("File size",         sizeMb),
                new("Dimensions",        dims),
                new("Camera make",       GetS("cameraMake")),
                new("Camera model",      GetS("cameraModel")),
                new("Lens model",        GetS("lensModel")),
                new("F-number",          GetS("fNumber") is string f && f != "—" ? $"f/{f}" : "—"),
                new("Exposure",          GetS("exposureDisplay") != "—" ? GetS("exposureDisplay") : GetS("exposureSeconds") + " s"),
                new("ISO",               GetS("iso")),
                new("Date taken",        GetS("dateTimeOriginal") != "—" ? GetS("dateTimeOriginal") : GetS("modifiedUtc"))
            };
        }

        /* ======================= CONTEXT MENUS & RIGHT-CLICK SELECTION ======================= */

        private void EnsureImagesContextMenu()
        {
            if (ImagesList.ContextMenu != null) return;

            var cm = new ContextMenu();

            cm.Items.Add(MenuItem("Open image", (_, __) =>
            {
                if (ImagesList.SelectedItem is ImageItem it && File.Exists(it.FullPath))
                    _ = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(it.FullPath) { UseShellExecute = true });
            }));

            cm.Items.Add(MenuItem("Open containing folder", (_, __) =>
            {
                if (ImagesList.SelectedItem is ImageItem it && File.Exists(it.FullPath))
                    _ = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", $"/select,\"{it.FullPath}\""));
            }));

            cm.Items.Add(new Separator());

            cm.Items.Add(MenuItem("Copy file path", (_, __) =>
            {
                if (ImagesList.SelectedItem is ImageItem it) Clipboard.SetText(it.FullPath);
            }));

            cm.Items.Add(MenuItem("Copy file:// URL", (_, __) =>
            {
                if (ImagesList.SelectedItem is ImageItem it)
                {
                    try { Clipboard.SetText(new Uri(it.FullPath).AbsoluteUri); }
                    catch { }
                }
            }));

            cm.Items.Add(MenuItem("Copy file name", (_, __) =>
            {
                if (ImagesList.SelectedItem is ImageItem it) Clipboard.SetText(it.Name);
            }));

            ImagesList.ContextMenu = cm;
        }

        private void EnsureCsvContextMenu()
        {
            if (CsvRowsGrid.ContextMenu != null) return;

            var cm = new ContextMenu();

            cm.Items.Add(MenuItem("Copy row (CSV)", (_, __) =>
            {
                if (CsvRowsGrid.SelectedItem is CameraRow r)
                {
                    var s = $"{r.Name},{r.X.ToString(CultureInfo.InvariantCulture)},{r.Y.ToString(CultureInfo.InvariantCulture)},{r.Alt.ToString(CultureInfo.InvariantCulture)},{r.Heading.ToString(CultureInfo.InvariantCulture)},{r.Pitch.ToString(CultureInfo.InvariantCulture)},{r.Roll.ToString(CultureInfo.InvariantCulture)}";
                    Clipboard.SetText(s);
                }
            }));

            cm.Items.Add(MenuItem("Copy row (JSON)", (_, __) =>
            {
                if (CsvRowsGrid.SelectedItem is CameraRow r)
                    Clipboard.SetText(JsonSerializer.Serialize(r, s_jsonOpts));
            }));

            cm.Items.Add(new Separator());

            cm.Items.Add(MenuItem("Copy image name", (_, __) =>
            {
                if (CsvRowsGrid.SelectedItem is CameraRow r)
                    Clipboard.SetText(r.Name);
            }));

            CsvRowsGrid.ContextMenu = cm;
        }

        // XAML wires this:
        private void ImagesList_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            SelectRowUnderMouse(ImagesList, e.OriginalSource as DependencyObject);
            EnsureImagesContextMenu();
            // let the bubbling continue so ContextMenu opens normally
        }

        // XAML wires this:
        private void CsvRowsGrid_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            SelectRowUnderMouse(CsvRowsGrid, e.OriginalSource as DependencyObject);
            EnsureCsvContextMenu();
        }

        private static void SelectRowUnderMouse(DataGrid grid, DependencyObject? origin)
        {
            if (origin == null) return;

            // Walk up the tree to find the DataGridRow
            var row = FindAncestor<DataGridRow>(origin);
            if (row != null)
            {
                grid.SelectedItem = row.Item;
                row.IsSelected = true;
                row.Focus();
            }
            else
            {
                // Fallback: find a cell then its row
                var cell = FindAncestor<DataGridCell>(origin);
                if (cell?.DataContext != null)
                {
                    grid.SelectedItem = cell.DataContext;
                }
            }
        }

        private static T? FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T t) return t;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        private static MenuItem MenuItem(string text, RoutedEventHandler onClick)
        {
            var mi = new MenuItem { Header = text };
            mi.Click += onClick;
            return mi;
        }
    }

    public class CameraRow
    {
        public string Name { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double Alt { get; set; }
        public double Heading { get; set; }
        public double Pitch { get; set; }
        public double Roll { get; set; }
        public bool HasImage { get; set; }
        // Preserve OPK when CSV contains omega/phi/kappa (kept internal so it doesn't clutter the DataGrid)
        internal double Omega { get; set; }
        internal double Phi { get; set; }
        internal double Kappa { get; set; }
    }

    internal static class GeoMath
    {
        const double a = 6378137.0;
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

        // Converts RealityCapture OPK in degrees to Cesium heading/pitch/roll in radians.
        // OPK convention (photogrammetry): R_cam2world = Rz(kappa) * Ry(phi) * Rx(omega).
        // Cesium HPR: R_body2world = Rz(-heading) * Ry(-pitch) * Rx(roll), with body axes X fwd, Y right, Z up.
        public static (double heading, double pitch, double roll) OpkDegToCesiumHprRad(double omegaDeg, double phiDeg, double kappaDeg)
        {
            double w = DegToRad(omegaDeg);
            double p = DegToRad(phiDeg);
            double k = DegToRad(kappaDeg);

            // Basic rotation matrices (right-handed)
            static void Rx(double a, out double r00, out double r01, out double r02,
                                     out double r10, out double r11, out double r12,
                                     out double r20, out double r21, out double r22)
            {
                double ca = Math.Cos(a), sa = Math.Sin(a);
                r00 = 1; r01 = 0;  r02 = 0;
                r10 = 0; r11 = ca; r12 = -sa;
                r20 = 0; r21 = sa; r22 = ca;
            }

            static void Ry(double a, out double r00, out double r01, out double r02,
                                     out double r10, out double r11, out double r12,
                                     out double r20, out double r21, out double r22)
            {
                double ca = Math.Cos(a), sa = Math.Sin(a);
                r00 = ca;  r01 = 0; r02 = sa;
                r10 = 0;   r11 = 1; r12 = 0;
                r20 = -sa; r21 = 0; r22 = ca;
            }

            static void Rz(double a, out double r00, out double r01, out double r02,
                                     out double r10, out double r11, out double r12,
                                     out double r20, out double r21, out double r22)
            {
                double ca = Math.Cos(a), sa = Math.Sin(a);
                r00 = ca; r01 = -sa; r02 = 0;
                r10 = sa; r11 = ca;  r12 = 0;
                r20 = 0;  r21 = 0;   r22 = 1;
            }

            // R_cam2world = Rz(k) * Ry(p) * Rx(w)
            Rz(k, out var az00, out var az01, out var az02, out var az10, out var az11, out var az12, out var az20, out var az21, out var az22);
            Ry(p, out var ay00, out var ay01, out var ay02, out var ay10, out var ay11, out var ay12, out var ay20, out var ay21, out var ay22);
            Rx(w, out var ax00, out var ax01, out var ax02, out var ax10, out var ax11, out var ax12, out var ax20, out var ax21, out var ax22);

            // temp = Rz(k) * Ry(p)
            double t00 = az00 * ay00 + az01 * ay10 + az02 * ay20;
            double t01 = az00 * ay01 + az01 * ay11 + az02 * ay21;
            double t02 = az00 * ay02 + az01 * ay12 + az02 * ay22;
            double t10 = az10 * ay00 + az11 * ay10 + az12 * ay20;
            double t11 = az10 * ay01 + az11 * ay11 + az12 * ay21;
            double t12 = az10 * ay02 + az11 * ay12 + az12 * ay22;
            double t20 = az20 * ay00 + az21 * ay10 + az22 * ay20;
            double t21 = az20 * ay01 + az21 * ay11 + az22 * ay21;
            double t22 = az20 * ay02 + az21 * ay12 + az22 * ay22;

            // R_cam2world = temp * Rx(w)
            double rc00 = t00 * ax00 + t01 * ax10 + t02 * ax20;
            double rc01 = t00 * ax01 + t01 * ax11 + t02 * ax21;
            double rc02 = t00 * ax02 + t01 * ax12 + t02 * ax22;
            double rc10 = t10 * ax00 + t11 * ax10 + t12 * ax20;
            double rc11 = t10 * ax01 + t11 * ax11 + t12 * ax21;
            double rc12 = t10 * ax02 + t11 * ax12 + t12 * ax22;
            double rc20 = t20 * ax00 + t21 * ax10 + t22 * ax20;
            double rc21 = t20 * ax01 + t21 * ax11 + t22 * ax21;
            double rc22 = t20 * ax02 + t21 * ax12 + t22 * ax22;

            // Basis change: camera -> body where body X fwd, Y right, Z up
            // C = body<-camera = [[0,0,1],[1,0,0],[0,-1,0]], so R_body2world = R_cam2world * C^T
            // C^T = [[0,1,0],[0,0,-1],[1,0,0]]
            // Multiply rc * Ct
            double r00 = rc00 * 0 + rc01 * 0 + rc02 * 1;
            double r01 = rc00 * 1 + rc01 * 0 + rc02 * 0;
            double r02 = rc00 * 0 + rc01 * -1 + rc02 * 0;
            double r10 = rc10 * 0 + rc11 * 0 + rc12 * 1;
            double r11 = rc10 * 1 + rc11 * 0 + rc12 * 0;
            double r12 = rc10 * 0 + rc11 * -1 + rc12 * 0;
            double r20 = rc20 * 0 + rc21 * 0 + rc22 * 1;
            double r21 = rc20 * 1 + rc21 * 0 + rc22 * 0;
            double r22 = rc20 * 0 + rc21 * -1 + rc22 * 0;

            // Extract Cesium heading/pitch/roll from R_body2world
            double heading = -Math.Atan2(r10, r00);
            double s = Math.Max(-1.0, Math.Min(1.0, r20));
            double pitch = Math.Asin(s);
            double roll = Math.Atan2(r21, r22);

            // Canonicalise Euler near-nadir to ensure down-facing (negative pitch) branch only when clearly vertical.
            // Forward vector in world is the first column of R_body2world: f = [r00, r10, r20]^T, so f.z = r20.
            const double tau = 0.5; // ~= cos(60°). Use 0.7 for stricter if desired.
            if (r20 > tau)
            {
                // "Up" near-nadir branch. Flip to the equivalent down-facing solution.
                heading += Math.PI;
                pitch = -pitch;
                roll += Math.PI;

                // Wrap to [-pi, pi] for clean outputs
                heading = WrapPi(heading);
                roll = WrapPi(roll);
                // pitch is in [-pi/2, pi/2] already via Asin and sign flip maintains range
            }

            return (heading, pitch, roll);
        }

        // Overload with lon/lat for API parity with quaternion builder. Values are ENU-relative
        // and independent of lon/lat, so we delegate to the core method.
        public static (double heading, double pitch, double roll) OpkDegToCesiumHprRad(
            double omegaDeg, double phiDeg, double kappaDeg, double lonDeg, double latDeg)
            => OpkDegToCesiumHprRad(omegaDeg, phiDeg, kappaDeg);

        // Converts RealityCapture OPK (degrees) + lon/lat (degrees) into a Cesium-ready quaternion (x,y,z,w) in ECEF.
        // Steps (matches Cesium reference):
        // 1) R_cam->ENU = Rz(k) * Ry(p) * Rx(w)    [photogrammetry order]
        // 2) Camera->body remap (camera: X right, Y down, Z fwd) -> (body: +X fwd, +Y right, +Z up)
        //    C_body<-cam = [[0,0,1],[1,0,0],[0,-1,0]], so R_body->ENU = R_cam->ENU * (C_body<-cam)^T
        // 3) ENU->ECEF for lon/lat: pre-multiply R_body->ECEF = (ENU->ECEF) * R_body->ENU
        // 4) Convert to quaternion [x,y,z,w]
        public static (double qx, double qy, double qz, double qw) OpkDegToCesiumQuaternion(
            double omegaDeg, double phiDeg, double kappaDeg, double lonDeg, double latDeg)
        {
            double w = DegToRad(omegaDeg);
            double p = DegToRad(phiDeg);
            double k = DegToRad(kappaDeg);

            static double[] RotX(double a)
            {
                double c = Math.Cos(a), s = Math.Sin(a);
                return new double[] { 1,0,0, 0,c,-s, 0,s,c };
            }
            static double[] RotY(double a)
            {
                double c = Math.Cos(a), s = Math.Sin(a);
                return new double[] { c,0,s, 0,1,0, -s,0,c };
            }
            static double[] RotZ(double a)
            {
                double c = Math.Cos(a), s = Math.Sin(a);
                return new double[] { c,-s,0, s,c,0, 0,0,1 };
            }
            static double[] Mul(double[] A, double[] B)
            {
                // Row-major 3x3 multiply: C = A * B
                return new double[]
                {
                    A[0]*B[0] + A[1]*B[3] + A[2]*B[6],
                    A[0]*B[1] + A[1]*B[4] + A[2]*B[7],
                    A[0]*B[2] + A[1]*B[5] + A[2]*B[8],

                    A[3]*B[0] + A[4]*B[3] + A[5]*B[6],
                    A[3]*B[1] + A[4]*B[4] + A[5]*B[7],
                    A[3]*B[2] + A[4]*B[5] + A[5]*B[8],

                    A[6]*B[0] + A[7]*B[3] + A[8]*B[6],
                    A[6]*B[1] + A[7]*B[4] + A[8]*B[7],
                    A[6]*B[2] + A[7]*B[5] + A[8]*B[8],
                };
            }
            static double[] Transpose(double[] M)
            {
                return new double[]
                {
                    M[0], M[3], M[6],
                    M[1], M[4], M[7],
                    M[2], M[5], M[8]
                };
            }

            // 1) camera->ENU
            var Rcw = Mul(RotZ(k), Mul(RotY(p), RotX(w)));

            // 2) camera->body remap (RealityCapture camera ~ NED: X fwd, Y right, Z down)
            // Cesium body: X fwd, Y right, Z up => simple Z flip
            double[] C_body_cam = new double[] { 1,0,0, 0,1,0, 0,0,-1 };
            var Ct = Transpose(C_body_cam);
            var Rbw_ENU = Mul(Rcw, Ct);

            // 3) ENU->ECEF at lon/lat (columns are E,N,U)
            double lon = DegToRad(lonDeg);
            double lat = DegToRad(latDeg);
            double sinLon = Math.Sin(lon), cosLon = Math.Cos(lon);
            double sinLat = Math.Sin(lat), cosLat = Math.Cos(lat);
            var E = new double[]
            {
                -sinLon,             -sinLat * cosLon,  cosLat * cosLon,
                 cosLon,             -sinLat * sinLon,  cosLat * sinLon,
                 0,                   cosLat,            sinLat
            };

            // 4) body->ECEF
            var R = Mul(E, Rbw_ENU);

            // Matrix3 -> quaternion (x,y,z,w)
            double r00 = R[0], r01 = R[1], r02 = R[2];
            double r10 = R[3], r11 = R[4], r12 = R[5];
            double r20 = R[6], r21 = R[7], r22 = R[8];
            double trace = r00 + r11 + r22;
            double qx, qy, qz, qw;
            if (trace > 0)
            {
                double s = Math.Sqrt(trace + 1.0) * 2.0; // s = 4*qw
                qw = 0.25 * s;
                qx = (r21 - r12) / s;
                qy = (r02 - r20) / s;
                qz = (r10 - r01) / s;
            }
            else if (r00 > r11 && r00 > r22)
            {
                double s = Math.Sqrt(1.0 + r00 - r11 - r22) * 2.0; // s = 4*qx
                qw = (r21 - r12) / s;
                qx = 0.25 * s;
                qy = (r01 + r10) / s;
                qz = (r02 + r20) / s;
            }
            else if (r11 > r22)
            {
                double s = Math.Sqrt(1.0 + r11 - r00 - r22) * 2.0; // s = 4*qy
                qw = (r02 - r20) / s;
                qx = (r01 + r10) / s;
                qy = 0.25 * s;
                qz = (r12 + r21) / s;
            }
            else
            {
                double s = Math.Sqrt(1.0 + r22 - r00 - r11) * 2.0; // s = 4*qz
                qw = (r10 - r01) / s;
                qx = (r02 + r20) / s;
                qy = (r12 + r21) / s;
                qz = 0.25 * s;
            }

            // Normalize for safety
            double norm = Math.Sqrt(qx * qx + qy * qy + qz * qz + qw * qw);
            if (norm > 0)
            {
                qx /= norm; qy /= norm; qz /= norm; qw /= norm;
            }
            return (qx, qy, qz, qw);
        }

        private static double WrapPi(double a)
        {
            const double pi = Math.PI;
            const double twoPi = 2 * Math.PI;
            a = (a + pi) % twoPi;
            if (a < 0) a += twoPi;
            return a - pi;
        }
    }
}
