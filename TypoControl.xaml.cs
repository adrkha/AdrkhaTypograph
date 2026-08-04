using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using SkiaSharp;
using System.Windows.Media;

namespace AdrkhaTypograph
{
    public partial class TypoControl : UserControl
    {
        public class FeatureViewModel : DependencyObject
        {
            public string Tag { get; set; }
            public string Description { get; set; }

            public static readonly DependencyProperty IsEnabledProperty =
                DependencyProperty.Register("IsEnabled", typeof(bool), typeof(FeatureViewModel),
                    new PropertyMetadata(false, OnFeatureChanged));

            public bool IsEnabled
            {
                get { return (bool)GetValue(IsEnabledProperty); }
                set { SetValue(IsEnabledProperty, value); }
            }

            private static void OnFeatureChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
            {
                var vm = d as FeatureViewModel;
                vm?.Changed?.Invoke();
            }

            public Action Changed { get; set; }
        }

        public class FontItemViewModel
        {
            public string DisplayName { get; set; } // اسم عائلة الخط + اسم الملف للتوضيح
            public string FilePath { get; set; }
            public string SearchKey { get; set; } // مفتاح البحث بالصيغة الصغيرة للفلترة
        }

        private string _selectedFontPath;
        private ObservableCollection<FeatureViewModel> _features = new ObservableCollection<FeatureViewModel>();
        private DispatcherTimer _previewTimer;
        
        // تتبع الشكل المحدد حالياً للتعديل
        private PowerPoint.Shape _selectedShape = null;
        private bool _isUpdatingUiFromSelection = false;

        // القائمة الرئيسية للخطوط الممسوحة
        private List<FontItemViewModel> _masterFontList = new List<FontItemViewModel>();

        // لون المعاينة المتغير ديناميكياً مع الثيم
        private SKColor _previewTextColor = new SKColor(17, 24, 39);

        // مؤقت ومؤشر لتتبع وتطبيق ثيم بوربوينت تلقائياً
        private DispatcherTimer _themeTimer;
        private bool? _currentThemeIsDark = null;

        // نظام التحديث التلقائي
        private UpdateChecker _updateChecker;

        public TypoControl()
        {
            // تهيئة مؤقت تحديث المعاينة لمنع التحميل الزائد (Debouncing 250ms)
            // نقوم بالتهيئة قبل InitializeComponent لتفادي NullReferenceException عند إطلاق الأحداث أثناء تحميل واجهة XAML
            _previewTimer = new DispatcherTimer();
            _previewTimer.Interval = TimeSpan.FromMilliseconds(250);
            _previewTimer.Tick += PreviewTimer_Tick;

            InitializeComponent();
            FeaturesList.ItemsSource = _features;

            // الاشتراك في حدث تغيير التحديد في بوربوينت
            try
            {
                Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to subscribe to selection changes: " + ex.Message);
            }

            // تطبيق ثيم بوربوينت الحالي تلقائياً
            ApplyPowerPointTheme();

            // إعداد مؤقت للتحقق من تغيير ثيم البرنامج كل ثانية لتحديث الألوان تلقائياً
            _themeTimer = new DispatcherTimer();
            _themeTimer.Interval = TimeSpan.FromSeconds(1);
            _themeTimer.Tick += (s, ev) => ApplyPowerPointTheme();
            _themeTimer.Start();

            // تحميل المجلد الأخير المختار تلقائياً إن وُجد
            string lastFolder = LoadSelectedFolder();
            if (!string.IsNullOrEmpty(lastFolder))
            {
                ScanFolderAsync(lastFolder);
            }

            // بدء فحص التحديثات في الخلفية (لا يعلّق الواجهة)
            _updateChecker = new UpdateChecker();
            _updateChecker.UpdateDetected += OnUpdateDetected;
            _ = _updateChecker.CheckAsync();
        }

        // كشف وتطبيق ثيم بوربوينت تلقائياً لتوافق الوضع النهاري والليلي
        private void ApplyPowerPointTheme()
        {
            bool isDark = false;
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Common"))
                {
                    if (key != null)
                    {
                        object themeObj = key.GetValue("UI Theme");
                        if (themeObj != null)
                        {
                            int themeVal = Convert.ToInt32(themeObj);
                            
                            // قيم ثيم مايكروسوفت أوفيس:
                            // 3 = داكن رمادي، 4 = أسود
                            if (themeVal == 3 || themeVal == 4)
                            {
                                isDark = true;
                            }
                            // 6 = استخدام إعدادات النظام
                            else if (themeVal == 6)
                            {
                                using (var winKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                                {
                                    if (winKey != null)
                                    {
                                        isDark = Convert.ToInt32(winKey.GetValue("AppsUseLightTheme", 1)) == 0;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to detect Office theme from registry: " + ex.Message);
            }

            // نقوم بالتحديث فقط إذا تغير المظهر فعلياً لمنع التكرار وتجميد المعاينة الحية
            if (_currentThemeIsDark == null || _currentThemeIsDark.Value != isDark)
            {
                _currentThemeIsDark = isDark;
                UpdateThemeResources(isDark);
            }
        }

        private void UpdateThemeResources(bool isDark)
        {
            try
            {
                if (isDark)
                {
                    // الوضع الليلي (Dark Mode)
                    this.Resources["ThemeBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#1E1E1E");
                    this.Resources["ThemeForegroundColor"] = (Color)ColorConverter.ConvertFromString("#F0F0F0");
                    this.Resources["ThemeCardBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#2D2D2D");
                    this.Resources["ThemeTextBoxBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#252526");
                    this.Resources["ThemeBorderBrushColor"] = (Color)ColorConverter.ConvertFromString("#3F3F46");
                    this.Resources["ThemeTextMutedColor"] = (Color)ColorConverter.ConvertFromString("#A0A0A0");
                    this.Resources["ThemeHighlightColor"] = (Color)ColorConverter.ConvertFromString("#3E3E42");
                    this.Resources["PreviewBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#111111");
                    this.Resources["ComboSelectedBackground"] = (Color)ColorConverter.ConvertFromString("#0066FF");
                    this.Resources["ComboSelectedForeground"] = (Color)ColorConverter.ConvertFromString("#FFFFFF");
                    
                    _previewTextColor = new SKColor(255, 215, 0); // رسم المعاينة بالذهبي المضيء
                }
                else
                {
                    // الوضع النهاري (Light Mode)
                    this.Resources["ThemeBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#F9FAFB");
                    this.Resources["ThemeForegroundColor"] = (Color)ColorConverter.ConvertFromString("#111827");
                    this.Resources["ThemeCardBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#FFFFFF");
                    this.Resources["ThemeTextBoxBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#FFFFFF");
                    this.Resources["ThemeBorderBrushColor"] = (Color)ColorConverter.ConvertFromString("#D1D5DB");
                    this.Resources["ThemeTextMutedColor"] = (Color)ColorConverter.ConvertFromString("#4B5563");
                    this.Resources["ThemeHighlightColor"] = (Color)ColorConverter.ConvertFromString("#F3F4F6");
                    this.Resources["PreviewBackgroundColor"] = (Color)ColorConverter.ConvertFromString("#FFFFFF");
                    this.Resources["ComboSelectedBackground"] = (Color)ColorConverter.ConvertFromString("#E0F2FE");
                    this.Resources["ComboSelectedForeground"] = (Color)ColorConverter.ConvertFromString("#0066FF");
                    
                    _previewTextColor = new SKColor(17, 24, 39); // رسم المعاينة بالكحلي/الأسود الداكن الفاخر
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to update theme resources: " + ex.Message);
            }
        }

        private void BtnBrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "اختر مجلد الخطوط العربية (.otf / .ttf)";
                dialog.ShowNewFolderButton = false;

                string lastFolder = LoadSelectedFolder();
                if (!string.IsNullOrEmpty(lastFolder))
                {
                    dialog.SelectedPath = lastFolder;
                }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string folderPath = dialog.SelectedPath;
                    SaveSelectedFolder(folderPath);
                    ScanFolderAsync(folderPath);
                }
            }
        }

        private async void ScanFolderAsync(string folderPath)
        {
            TxtFolderPath.Text = folderPath;
            TxtFolderPath.ToolTip = folderPath;
            GridFontSelection.Visibility = Visibility.Collapsed;
            
            _features.Clear();
            TxtNoFeatures.Visibility = Visibility.Visible;
            TxtNoFeatures.Text = "جاري مسح المجلد واستخراج أسماء عائلات الخطوط العربية...";

            var fontList = await Task.Run(() =>
            {
                var list = new List<FontItemViewModel>();
                try
                {
                    if (Directory.Exists(folderPath))
                    {
                        var files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
                        foreach (var file in files)
                        {
                            string ext = Path.GetExtension(file).ToLower();
                            if (ext == ".otf" || ext == ".ttf")
                            {
                                string familyName = string.Empty;
                                try
                                {
                                    using (var fs = File.OpenRead(file))
                                    using (var tf = SKTypeface.FromStream(fs))
                                    {
                                        if (tf != null)
                                            familyName = tf.FamilyName;
                                    }
                                }
                                catch { }

                                if (string.IsNullOrEmpty(familyName))
                                {
                                    familyName = Path.GetFileNameWithoutExtension(file);
                                }

                                list.Add(new FontItemViewModel
                                {
                                    DisplayName = string.Format("{0} ({1})", familyName, Path.GetFileName(file)),
                                    FilePath = file,
                                    SearchKey = string.Format("{0} {1}", familyName, Path.GetFileName(file)).ToLower()
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Folder scan error: " + ex.Message);
                }
                return list;
            });

            _masterFontList = fontList;

            if (_masterFontList.Count > 0)
            {
                GridFontSelection.Visibility = Visibility.Visible;
                TxtNoFeatures.Visibility = Visibility.Visible;
                TxtNoFeatures.Text = "يرجى اختيار خط من القائمة المنسدلة أعلاه لبدء الكتابة.";
                PopulateFontsComboBox(_masterFontList);
            }
            else
            {
                TxtNoFeatures.Visibility = Visibility.Visible;
                TxtNoFeatures.Text = "لم يتم العثور على أي خطوط بصيغة .otf أو .ttf داخل المجلد المحدد.";
            }
        }

        private void PopulateFontsComboBox(List<FontItemViewModel> list)
        {
            string selectedPath = _selectedFontPath;

            ComboFonts.SelectionChanged -= ComboFonts_SelectionChanged; // تعطيل مؤقت للأحداث
            ComboFonts.Items.Clear();

            Style itemStyle = FindResource("ModernComboBoxItemStyle") as Style;
            foreach (var font in list)
            {
                ComboFonts.Items.Add(new ComboBoxItem
                {
                    Content = font.DisplayName,
                    Tag = font.FilePath,
                    ToolTip = font.FilePath,
                    Style = itemStyle
                });
            }

            bool restored = false;
            for (int i = 0; i < ComboFonts.Items.Count; i++)
            {
                var item = ComboFonts.Items[i] as ComboBoxItem;
                if (item != null && (string)item.Tag == selectedPath)
                {
                    ComboFonts.SelectedIndex = i;
                    restored = true;
                    break;
                }
            }

            if (!restored && ComboFonts.Items.Count > 0 && string.IsNullOrEmpty(selectedPath))
            {
                ComboFonts.SelectedIndex = 0;
                var item = ComboFonts.Items[0] as ComboBoxItem;
                _selectedFontPath = (string)item?.Tag;
                LoadFontFeatures();
            }

            ComboFonts.SelectionChanged += ComboFonts_SelectionChanged; // إعادة التفعيل
        }

        private void TxtSearchFont_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TxtSearchPlaceholder != null)
            {
                TxtSearchPlaceholder.Visibility = string.IsNullOrEmpty(TxtSearchFont.Text) ? Visibility.Visible : Visibility.Collapsed;
            }
            FilterFonts(TxtSearchFont.Text);
        }

        private void FilterFonts(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                PopulateFontsComboBox(_masterFontList);
                return;
            }

            string q = query.ToLower().Trim();
            var filtered = _masterFontList
                .Where(f => f.SearchKey.Contains(q))
                .ToList();

            PopulateFontsComboBox(filtered);
        }

        private void ComboFonts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComboFonts.SelectedItem is ComboBoxItem item)
            {
                string path = item.Tag as string;
                if (!string.IsNullOrEmpty(path) && File.Exists(path) && path != _selectedFontPath)
                {
                    _selectedFontPath = path;
                    LoadFontFeatures();
                    TriggerPreviewUpdate();
                }
            }
        }

        private void LoadFontFeatures()
        {
            _features.Clear();

            if (string.IsNullOrEmpty(_selectedFontPath) || !File.Exists(_selectedFontPath))
            {
                TxtNoFeatures.Visibility = Visibility.Visible;
                TxtNoFeatures.Text = "يرجى اختيار خط صالح لعرض ميزات التنسيق.";
                return;
            }

            var allTags = OpenTypeParser.GetFeatures(_selectedFontPath);
            
            var togglableTags = allTags
                .Where(IsUserTogglableFeature)
                .OrderBy(t => t)
                .ToList();

            if (togglableTags.Count > 0)
            {
                TxtNoFeatures.Visibility = Visibility.Collapsed;
                foreach (var tag in togglableTags)
                {
                    _features.Add(new FeatureViewModel
                    {
                        Tag = tag,
                        Description = GetFeatureDescription(tag),
                        IsEnabled = (tag == "calt" || tag == "liga"),
                        Changed = TriggerPreviewUpdate
                    });
                }
            }
            else
            {
                TxtNoFeatures.Visibility = Visibility.Visible;
                TxtNoFeatures.Text = "لم يتم العثور على ميزات OpenType اختيارية في هذا الخط.";
            }
        }

        private bool IsUserTogglableFeature(string tag)
        {
            tag = tag.ToLower().Trim();

            if (tag == "init" || tag == "medi" || tag == "fina" || tag == "isol" ||
                tag == "ccmp" || tag == "rlig" || tag == "locl" || tag == "mset" ||
                tag == "kern" || tag == "mark" || tag == "mkmk")
            {
                return false;
            }

            if (tag.StartsWith("ss") || tag.StartsWith("cv") ||
                tag == "swsh" || tag == "dlig" || tag == "hlig" ||
                tag == "salt" || tag == "titl" || tag == "liga" || tag == "calt")
            {
                return true;
            }

            return false;
        }

        private string GetFeatureDescription(string tag)
        {
            tag = tag.ToLower().Trim();

            if (tag.StartsWith("ss") && tag.Length == 4 && int.TryParse(tag.Substring(2), out int ssNum))
            {
                return string.Format("مجموعة أنماط {0:00}", ssNum);
            }
            if (tag.StartsWith("cv") && tag.Length == 4 && int.TryParse(tag.Substring(2), out int cvNum))
            {
                return string.Format("بديل حرفي {0:00}", cvNum);
            }

            switch (tag)
            {
                case "swsh": return "كشيدة / زينة جمالية (Swash)";
                case "calt": return "بدائل سياقية (Contextual)";
                case "liga": return "تربيطات قياسية (Ligatures)";
                case "dlig": return "تربيطات جمالية (Discretionary)";
                case "hlig": return "تربيطات تاريخية (Historical)";
                case "salt": return "بدائل تشكيلية (Stylistic)";
                case "titl": return "بدائل العناوين (Titling)";
                default: return string.Format("ميزة ({0})", tag.ToUpper());
            }
        }

        private void TriggerPreviewUpdate()
        {
            ApplyPowerPointTheme();
            if (_previewTimer != null)
            {
                _previewTimer.Stop();
                _previewTimer.Start();
            }
        }

        private void PreviewTimer_Tick(object sender, EventArgs e)
        {
            _previewTimer.Stop();
            UpdatePreview();
        }

        private TypoAlignment GetSelectedAlignment()
        {
            if (RadAlignLeft.IsChecked == true) return TypoAlignment.Left;
            if (RadAlignCenter.IsChecked == true) return TypoAlignment.Center;
            return TypoAlignment.Right; // افتراضي
        }

        private void UpdatePreview()
        {
            string text = TxtInputText.Text;
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(_selectedFontPath))
            {
                ImgPreview.Source = null;
                TxtPreviewPlaceholder.Visibility = Visibility.Visible;
                TxtPreviewPlaceholder.Text = string.IsNullOrEmpty(_selectedFontPath) ? "اختر خطاً لعرض المعاينة..." : "اكتب نصاً للمعاينة...";
                return;
            }

            TxtPreviewPlaceholder.Visibility = Visibility.Collapsed;

            var activeFeatures = _features
                .Where(f => f.IsEnabled)
                .Select(f => f.Tag)
                .ToList();

            var previewSource = TypoEngine.RenderPreviewImage(text, _selectedFontPath, activeFeatures, _previewTextColor, GetSelectedAlignment());
            if (previewSource != null)
            {
                ImgPreview.Source = previewSource;
            }
            else
            {
                ImgPreview.Source = null;
                TxtPreviewPlaceholder.Visibility = Visibility.Visible;
                TxtPreviewPlaceholder.Text = "خطأ في إنشاء المعاينة.";
            }
        }

        private void TxtInputText_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isUpdatingUiFromSelection)
                TriggerPreviewUpdate();
        }

        private void AlignmentButton_Checked(object sender, RoutedEventArgs e)
        {
            TriggerPreviewUpdate();
        }

        private void BtnInsertUpdate_Click(object sender, RoutedEventArgs e)
        {
            string text = TxtInputText.Text;
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("يرجى كتابة نص أولاً.", "تنبيه", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(_selectedFontPath) || !File.Exists(_selectedFontPath))
            {
                MessageBox.Show("يرجى اختيار خط صالح أولاً من القائمة.", "تنبيه", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var activeFeatures = _features
                .Where(f => f.IsEnabled)
                .Select(f => f.Tag)
                .ToList();

            try
            {
                string svgFile = TypoEngine.CreateVectorText(text, _selectedFontPath, activeFeatures, GetSelectedAlignment());

                if (_selectedShape != null)
                {
                    // وضع التحديث لشكل محدد مسبقاً
                    PowerPoint.Slide slide = null;
                    try { slide = _selectedShape.Parent as PowerPoint.Slide; } catch { }

                    if (slide == null)
                    {
                        slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                    }

                    float left = _selectedShape.Left;
                    float top = _selectedShape.Top;

                    // حذف الشكل القديم
                    _selectedShape.Delete();
                    _selectedShape = null;

                    // إدراج الشكل الجديد في نفس موقعه
                    var newShape = slide.Shapes.AddPicture(svgFile,
                        Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                        left, top, -1, -1);

                    // إضافة البيانات الوصفية للشكل الجديد
                    newShape.Tags.Add("OriginalText", text);
                    newShape.Tags.Add("FontPath", _selectedFontPath);
                    newShape.Tags.Add("Alignment", GetSelectedAlignment().ToString());
                    newShape.Tags.Add("ActiveFeatures", string.Join(",", activeFeatures));

                    newShape.Select();
                }
                else
                {
                    // وضع الإدراج العادي لشكل جديد
                    PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                    if (pptApp.Presentations.Count == 0 || pptApp.ActiveWindow == null)
                    {
                        MessageBox.Show("لا يوجد عرض تقديمي مفتوح حالياً في بوربوينت.", "تنبيه", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    PowerPoint.Slide activeSlide = pptApp.ActiveWindow.View.Slide;
                    if (activeSlide == null)
                    {
                        MessageBox.Show("لا توجد شريحة نشطة لإدراج النص بها.", "تنبيه", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var newShape = activeSlide.Shapes.AddPicture(svgFile,
                        Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                        100, 100, -1, -1);

                    newShape.Tags.Add("OriginalText", text);
                    newShape.Tags.Add("FontPath", _selectedFontPath);
                    newShape.Tags.Add("Alignment", GetSelectedAlignment().ToString());
                    newShape.Tags.Add("ActiveFeatures", string.Join(",", activeFeatures));

                    if (ChkInsertEditableText.IsChecked == true)
                    {
                        string fontFamilyName = string.Empty;
                        try
                        {
                            using (var fs = File.OpenRead(_selectedFontPath))
                            using (var tf = SKTypeface.FromStream(fs))
                            {
                                if (tf != null)
                                    fontFamilyName = tf.FamilyName;
                            }
                        }
                        catch { }

                        if (string.IsNullOrEmpty(fontFamilyName))
                        {
                            fontFamilyName = Path.GetFileNameWithoutExtension(_selectedFontPath);
                        }

                        // إدراج صندوق نص قابل للتعديل أسفل الـ SVG مباشرة
                        float textLeft = newShape.Left;
                        float textTop = newShape.Top + newShape.Height + 12;
                        float textWidth = newShape.Width;
                        if (textWidth < 200) textWidth = 200;

                        PowerPoint.Shape txtShape = activeSlide.Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            textLeft, textTop, textWidth, 40);

                        txtShape.TextFrame.TextRange.Text = text;

                        try
                        {
                            txtShape.TextFrame.TextRange.Font.Name = fontFamilyName;
                            txtShape.TextFrame.TextRange.Font.Size = 24;

                            var align = GetSelectedAlignment();
                            if (align == TypoAlignment.Left)
                                txtShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                            else if (align == TypoAlignment.Center)
                                txtShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                            else
                                txtShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("Failed to set basic text style: " + ex.Message);
                        }

                        // تطبيق خصائص OpenType عبر dynamic late binding
                        // (المرجعية Office15 لا تعرّف OpenType، لكن بوربوينت 2016+ يدعمها وقت التشغيل)
                        try
                        {
                            dynamic font2 = txtShape.TextFrame2.TextRange.Font;
                            dynamic openType = font2.OpenType;

                            bool hasLiga = activeFeatures.Contains("liga");
                            bool hasDlig = activeFeatures.Contains("dlig");
                            bool hasClig = activeFeatures.Contains("clig");
                            bool hasCalt = activeFeatures.Contains("calt");

                            // --- التراكيب (Ligatures): None=0, Standard=1, Contextual=2, HistDisc=4, All=7 ---
                            if (hasLiga && hasDlig)
                                openType.Ligatures = 7;   // msoLigaturesAll
                            else if (hasLiga && hasClig)
                                openType.Ligatures = 3;   // Standard | Contextual
                            else if (hasDlig)
                                openType.Ligatures = 4;   // msoLigaturesHistoricalAndDiscretionary
                            else if (hasClig)
                                openType.Ligatures = 2;   // msoLigaturesContextual
                            else if (hasLiga)
                                openType.Ligatures = 1;   // msoLigaturesStandard

                            // --- البدائل السياقية (calt) ---
                            if (hasCalt)
                                openType.ContextualAlternates = 1;

                            // --- المجموعات التنسيقية (ss01–ss20): القيمة الرقمية = رقم المجموعة ---
                            for (int ssNum = 1; ssNum <= 20; ssNum++)
                            {
                                string tag = "ss" + ssNum.ToString("D2");
                                if (activeFeatures.Contains(tag))
                                {
                                    openType.StylisticSets = ssNum;
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("OpenType late binding failed: " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطأ أثناء الإدراج أو التحديث في بوربوينت:\n" + ex.Message, "خطأ", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // معالجة حدث تغير التحديد في بوربوينت لتحميل الأشكال المدرجة مسبقاً وتعديلها
        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    // تحديث المظهر تلقائياً لمطابقة PowerPoint في حال تغييره أثناء العمل
                    ApplyPowerPointTheme();

                    if (Sel != null && Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && Sel.ShapeRange.Count == 1)
                    {
                        var shape = Sel.ShapeRange[1];
                        if (shape.Tags["OriginalText"] != null && shape.Tags["FontPath"] != null)
                        {
                            _selectedShape = shape;
                            _isUpdatingUiFromSelection = true;

                            // تعبئة البيانات في الواجهة
                            TxtInputText.Text = shape.Tags["OriginalText"];

                            string fontPath = shape.Tags["FontPath"];
                            if (File.Exists(fontPath))
                            {
                                _selectedFontPath = fontPath;
                                
                                bool fontInList = _masterFontList.Any(f => f.FilePath == fontPath);
                                if (!fontInList)
                                {
                                    string folder = Path.GetDirectoryName(fontPath);
                                    SaveSelectedFolder(folder);
                                    ScanFolderAsync(folder);
                                }
                                else
                                {
                                    for (int i = 0; i < ComboFonts.Items.Count; i++)
                                    {
                                        var item = ComboFonts.Items[i] as ComboBoxItem;
                                        if (item != null && (string)item.Tag == fontPath)
                                            ComboFonts.SelectedIndex = i;
                                    }
                                    LoadFontFeatures();
                                }
                            }

                            // محاذاة النص المخزنة
                            string alignStr = shape.Tags["Alignment"];
                            if (!string.IsNullOrEmpty(alignStr))
                            {
                                if (alignStr == "Left") RadAlignLeft.IsChecked = true;
                                else if (alignStr == "Center") RadAlignCenter.IsChecked = true;
                                else RadAlignRight.IsChecked = true;
                            }

                            // تفعيل الميزات التنسيقية المخزنة
                            string activeFeaturesStr = shape.Tags["ActiveFeatures"] ?? "";
                            var activeFeaturesList = activeFeaturesStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                            
                            foreach (var feature in _features)
                            {
                                feature.IsEnabled = activeFeaturesList.Contains(feature.Tag);
                            }

                            BtnInsertUpdate.Content = "تحديث الشكل المحدد";
                            _isUpdatingUiFromSelection = false;
                            
                            UpdatePreview();
                            return;
                        }
                    }

                    // في حال لم يتم تحديد شكل صالح للتحرير، ارجع لوضع الإدراج العادي
                    _selectedShape = null;
                    BtnInsertUpdate.Content = "إدراج في بوربوينت";
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Selection change tracking error: " + ex.Message);
                }
            }));
        }

        // --- منطق حفظ المجلد واستعادته ---
        private string GetFolderConfigFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "AdrkhaTypograph");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
            return Path.Combine(folder, "selected_folder.txt");
        }

        private string LoadSelectedFolder()
        {
            try
            {
                string file = GetFolderConfigFilePath();
                if (File.Exists(file))
                {
                    string path = File.ReadAllText(file, Encoding.UTF8).Trim();
                    if (Directory.Exists(path)) return path;
                }
            }
            catch { }
            return string.Empty;
        }

        private void SaveSelectedFolder(string path)
        {
            try
            {
                string file = GetFolderConfigFilePath();
                File.WriteAllText(file, path, Encoding.UTF8);
            }
            catch { }
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.Uri.AbsoluteUri));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to open URL: " + ex.Message);
            }
            e.Handled = true;
        }

        // ── نظام التحديث التلقائي ───────────────────────────────────────

        /// <summary>يُستدعى في الخلفية عند اكتشاف تحديث جديد.</summary>
        private void OnUpdateDetected(UpdateChecker checker)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                TxtUpdateMessage.Text = $"تحديث جديد متاح! (v{checker.LatestVersion})";
                UpdateBannerGrid.Visibility = Visibility.Visible;
            }));
        }

        /// <summary>زر "تحديث الآن": يحمّل المثبت الجديد ويشغّله.</summary>
        private async void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            BtnUpdate.IsEnabled = false;
            BtnDismissUpdate.IsEnabled = false;
            TxtUpdateMessage.Text = "جارٍ تحميل التحديث...";
            UpdateProgressBar.Visibility = Visibility.Visible;
            TxtDownloadStatus.Visibility = Visibility.Visible;

            _updateChecker.DownloadProgressChanged += pct =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    UpdateProgressBar.Value = pct;
                    TxtDownloadStatus.Text = $"جارٍ التحميل... {pct}%";
                }));
            };

            _updateChecker.DownloadCompleted += path =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (path != null)
                    {
                        TxtDownloadStatus.Text = "اكتمل التحميل. سيبدأ التثبيت الآن...";
                        UpdateChecker.LaunchInstaller(path);
                    }
                    else
                    {
                        TxtUpdateMessage.Text = "فشل التحميل. حاول مرة أخرى لاحقاً.";
                        BtnUpdate.IsEnabled = true;
                        BtnDismissUpdate.IsEnabled = true;
                    }
                }));
            };

            await _updateChecker.DownloadUpdateAsync();
        }

        /// <summary>زر "لاحقاً": يخفي الشريط حتى الجلسة التالية.</summary>
        private void BtnDismissUpdate_Click(object sender, RoutedEventArgs e)
        {
            UpdateBannerGrid.Visibility = Visibility.Collapsed;
        }
    }
}
