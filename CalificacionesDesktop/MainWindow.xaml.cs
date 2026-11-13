using ClosedXML.Excel;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace CalificacionesDesktop;

public partial class MainWindow : Window
{
    // Fuerza los estilos nativos de Windows para que la barra del sistema
    // aparezca incluso si WPF conserva un estilo anterior en la ventana.
    private const int GWL_STYLE = -16;
    private const long WS_CAPTION = 0x00C00000L;
    private const long WS_SYSMENU = 0x00080000L;
    private const long WS_MINIMIZEBOX = 0x00020000L;
    private const long WS_MAXIMIZEBOX = 0x00010000L;
    private const long WS_THICKFRAME = 0x00040000L;
    private const long WS_POPUP = unchecked((long)0x80000000);
    private const long WS_CHILD = 0x40000000L;
    private const long WS_OVERLAPPEDWINDOW = 0x00CF0000L;
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_FRAMECHANGED = 0x0020;

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);
    private readonly ObservableCollection<SourceItem> _sources = new();
    private readonly ObservableCollection<MatchRow> _rows = new();
    private string? _masterPath;
    private XLWorkbook? _masterBook;
    private MasterInfo? _masterInfo;
    private bool _analyzed;

    // Estado independiente del módulo de notas generales.
    private readonly ObservableCollection<SourceItem> _generalSources = new();
    private readonly ObservableCollection<MatchRow> _generalRows = new();
    private XLWorkbook? _generalMasterBook;
    private MasterInfo? _generalMasterInfo;
    private string? _generalMasterPath;
    private bool _generalAnalyzed;
    private int? _generalAnalyzedTrimester;
    private bool _generalHasBlockedSource;
    // Conserva todas las notas aplicadas durante la sesión, aunque cambie el trimestre
    // y se limpie la vista de análisis. La clave es fila y columna del maestro.
    private readonly Dictionary<(int Row, int Column), double> _generalPendingGrades = new();
    private bool _isResettingGeneral;

    public MainWindow()
    {
        InitializeComponent();
        // Mantener la barra nativa de Windows con minimizar, maximizar y cerrar.
        WindowStyle = WindowStyle.SingleBorderWindow;
        ResizeMode = ResizeMode.CanResize;
        WindowState = WindowState.Normal;
        ShowInTaskbar = true;
        SourcesList.ItemsSource = _sources;
        ResultsGrid.ItemsSource = _rows;
        GeneralSourcesList.ItemsSource = _generalSources;
        GeneralResultsGrid.ItemsSource = _generalRows;
        ShowMainMenu();
        SourceInitialized += (_, _) =>
        {
            // Reaplicar cuando WPF ya creó el HWND y después de terminar el Loaded.
            ApplyNativeWindowFrame();
            Dispatcher.BeginInvoke(new Action(ApplyNativeWindowFrame),
                System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        };
    }

    private void ApplyNativeWindowFrame()
    {
        var hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero) return;

        WindowStyle = WindowStyle.SingleBorderWindow;
        ResizeMode = ResizeMode.CanResize;
        ShowInTaskbar = true;

        // Normalizar el estilo completo: no basta con agregar WS_CAPTION si
        // el HWND conserva WS_POPUP/WS_CHILD de una ventana sin marco.
        long style = GetWindowLongPtr(hwnd, GWL_STYLE).ToInt64() & 0xFFFFFFFFL;
        style &= ~(WS_POPUP | WS_CHILD);
        style |= WS_OVERLAPPEDWINDOW | WS_CAPTION | WS_SYSMENU |
                 WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_THICKFRAME;
        SetWindowLongPtr(hwnd, GWL_STYLE, new IntPtr(style));
        SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    }

    private void ShowMainMenu()
    {
        MainMenuView.Visibility = Visibility.Visible;
        ProjectsModuleView.Visibility = Visibility.Collapsed;
        GeneralModuleView.Visibility = Visibility.Collapsed;
    }

    private void OpenProjects_Click(object sender, RoutedEventArgs e)
    {
        MainMenuView.Visibility = Visibility.Collapsed;
        GeneralModuleView.Visibility = Visibility.Collapsed;
        ProjectsModuleView.Visibility = Visibility.Visible;
    }

    private void OpenGeneral_Click(object sender, RoutedEventArgs e)
    {
        // Cada entrada al módulo comienza una sesión nueva y sin datos residuales.
        ResetGeneralModule();
        MainMenuView.Visibility = Visibility.Collapsed;
        ProjectsModuleView.Visibility = Visibility.Collapsed;
        GeneralModuleView.Visibility = Visibility.Visible;
    }

    private void BackToMenu_Click(object sender, RoutedEventArgs e)
    {
        // Al volver al menú se descarta todo el estado temporal del módulo general.
        ResetGeneralModule();
        ShowMainMenu();
    }

    private void ResetGeneralModule()
    {
        _isResettingGeneral = true;
        try
        {
            _generalMasterBook?.Dispose();
            _generalMasterBook = null;
            _generalMasterInfo = null;
            _generalMasterPath = null;
            _generalSources.Clear();
            _generalRows.Clear();
            _generalAnalyzed = false;
            _generalAnalyzedTrimester = null;
            _generalHasBlockedSource = false;
            _generalPendingGrades.Clear();

            if (GeneralMasterPath != null) GeneralMasterPath.Clear();
            if (GeneralTrimesterCombo != null) GeneralTrimesterCombo.SelectedIndex = 0;
            if (GeneralStatusText != null)
                GeneralStatusText.Text = "Selecciona el maestro, el trimestre y los archivos de los docentes.";
            GeneralSourcesList?.UnselectAll();
            GeneralResultsGrid?.Items.Refresh();
        }
        finally
        {
            _isResettingGeneral = false;
        }
    }

    private void ChooseGeneralMaster_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel maestro (*.xlsx)|*.xlsx" };
        if (dialog.ShowDialog() != true) return;
        try
        {
            var book = new XLWorkbook(dialog.FileName);
            var info = DetectGeneralMaster(book);
            _generalMasterBook?.Dispose();
            _generalMasterBook = book;
            _generalMasterInfo = info;
            _generalMasterPath = dialog.FileName;
            GeneralMasterPath.Text = dialog.FileName;
            // Un maestro nuevo inicia una sesión de cambios independiente.
            _generalPendingGrades.Clear();
            _generalRows.Clear();
            _generalAnalyzed = false;
            _generalAnalyzedTrimester = null;
            _generalHasBlockedSource = false;
            GeneralStatusText.Text = $"Maestro cargado: {info.StudentRows.Count} estudiantes. Hoja: {info.Worksheet.Name}.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"No se pudo leer el Excel maestro.\n\n{ex.Message}", "Error al cargar", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static MasterInfo DetectGeneralMaster(XLWorkbook book)
    {
        var ws = book.Worksheets.FirstOrDefault(x => x.Name.Equals("INGRESO-NOTAS", StringComparison.OrdinalIgnoreCase));
        if (ws == null) throw new InvalidOperationException("El archivo maestro no contiene la hoja INGRESO-NOTAS.");
        int headerRow = 6;
        int nameColumn = 3;
        string header = NormalizeIdentity(ws.Cell(headerRow, nameColumn).GetString());
        if (!header.Contains("APELLIDOS") || !header.Contains("NOMBRES"))
            throw new InvalidOperationException("No se encontró APELLIDOS/NOMBRES en C6 de INGRESO-NOTAS.");
        var students = new List<StudentRow>();
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRow;
        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            string name = ws.Cell(r, nameColumn).GetString().Trim();
            if (LooksLikeStudentName(name)) students.Add(new StudentRow(r, name));
        }
        if (students.Count == 0) throw new InvalidOperationException("No se encontraron estudiantes en INGRESO-NOTAS.");
        return new MasterInfo(ws, headerRow, nameColumn, students);
    }

    private void ChooseGeneralSources_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Archivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls", Multiselect = true };
        if (dialog.ShowDialog() != true) return;
        int added = 0;
        foreach (string path in dialog.FileNames)
        {
            if (_generalSources.Any(x => string.Equals(x.Path, path, StringComparison.OrdinalIgnoreCase))) continue;
            _generalSources.Add(new SourceItem { Path = path, FileName = Path.GetFileName(path), SheetName = "Pendiente", MatterName = "Por detectar", DetectionStatus = "En espera" });
            added++;
        }
        _generalAnalyzed = false;
        GeneralStatusText.Text = $"Se agregaron {added} archivo(s). Selecciona el trimestre y analiza.";
    }

    private void RemoveGeneralSources_Click(object sender, RoutedEventArgs e)
    {
        var selected = GeneralSourcesList.SelectedItems.Cast<SourceItem>().ToList();
        foreach (var item in selected) _generalSources.Remove(item);
        _generalRows.Clear();
        _generalAnalyzed = false;
        GeneralStatusText.Text = $"Archivos retirados: {selected.Count}.";
    }

    private void GeneralTrimesterCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_isResettingGeneral) return;
        // Las columnas de destino se calculan durante el análisis. Si cambia el trimestre,
        // invalida el análisis anterior para impedir que se escriba en la columna equivocada.
        if (_generalRows == null) return;
        _generalAnalyzed = false;
        _generalAnalyzedTrimester = null;
        _generalRows.Clear();
        if (GeneralStatusText != null && GeneralTrimesterCombo?.SelectedItem != null)
            GeneralStatusText.Text = $"Trimestre cambiado a {SelectedGeneralTrimester}. Pulsa 'Analizar archivos' nuevamente antes de aplicar.";
    }

    private int SelectedGeneralTrimester => (GeneralTrimesterCombo.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Tag is string tag && int.TryParse(tag, out int n) ? n : 1;

    private void AnalyzeGeneral_Click(object sender, RoutedEventArgs e)
    {
        if (_generalMasterBook == null || _generalMasterInfo == null)
        {
            MessageBox.Show("Primero selecciona el Excel maestro.", "Falta el maestro", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (_generalSources.Count == 0)
        {
            MessageBox.Show("Agrega al menos un archivo de los docentes.", "Faltan archivos", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _generalRows.Clear();
        _generalAnalyzed = false;
        _generalHasBlockedSource = false;
        int validGrades = 0, issues = 0;
        foreach (var source in _generalSources)
        {
            try
            {
                var result = AnalyzeGeneralSource(source, SelectedGeneralTrimester);
                source.SheetName = result.SheetName;
                source.MatterName = result.MatterName;
                source.DetectionStatus = result.ValidGradeCount == 0 ? "BLOQUEADO · Sin notas válidas" : $"{result.ValidGradeCount} nota(s) válida(s)";
                foreach (var row in result.Rows) _generalRows.Add(row);
                validGrades += result.ValidGradeCount;
                issues += result.Rows.Count(r => !r.Status.StartsWith("COINCIDENCIA", StringComparison.Ordinal));
                if (result.ValidGradeCount == 0) _generalHasBlockedSource = true;
            }
            catch (Exception ex)
            {
                source.DetectionStatus = "BLOQUEADO · " + ex.Message;
                _generalRows.Add(new MatchRow { SourceFile = source.FileName, Matter = source.MatterName, Status = "TRANSFERENCIA BLOQUEADA: " + ex.Message });
                _generalHasBlockedSource = true;
                issues++;
            }
        }
        GeneralSourcesList.Items.Refresh();
        GeneralResultsGrid.Items.Refresh();
        _generalAnalyzed = true;
        _generalAnalyzedTrimester = SelectedGeneralTrimester;
        GeneralStatusText.Text = $"Análisis terminado para el trimestre {SelectedGeneralTrimester}. Notas válidas: {validGrades}. Observaciones: {issues}.";
        if (validGrades == 0 || _generalHasBlockedSource)
            MessageBox.Show("Se detectaron archivos sin notas válidas o con errores. Revisa el reporte. Los archivos bloqueados no se transferirán.", "Validación de calificaciones", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private GeneralDetectedSource AnalyzeGeneralSource(SourceItem source, int trimester)
    {
        using var workbook = OpenSource(source.Path);
        string expectedSheet = trimester switch { 1 => "1ERTRIMESTRE", 2 => "2DOTRIMESTRE", _ => "3ERTRIMESTRE" };
        ISheet? sheet = null;
        for (int i = 0; i < workbook.NumberOfSheets; i++)
        {
            ISheet candidate = workbook.GetSheetAt(i);
            if (NormalizeIdentity(candidate.SheetName).Replace(" ", "").StartsWith(expectedSheet, StringComparison.Ordinal))
            {
                sheet = candidate;
                break;
            }
        }
        if (sheet == null) throw new InvalidOperationException($"No existe la hoja del {trimester}° trimestre en el archivo.");

        int gradeHeaderRow = -1;
        for (int r = 0; r <= Math.Min(sheet.LastRowNum, 20); r++)
        {
            string header = NormalizeIdentity(GetCellText(sheet.GetRow(r)?.GetCell(36)));
            if (header.Contains("TOTALTRIMESTRE")) { gradeHeaderRow = r; break; }
        }
        if (gradeHeaderRow < 0) throw new InvalidOperationException("No se encontró TOTAL TRIMESTRE (100%) en la columna AK.");
        int nameHeaderRow = -1;
        for (int r = Math.Max(0, gradeHeaderRow); r <= Math.Min(sheet.LastRowNum, gradeHeaderRow + 5); r++)
        {
            string header = NormalizeIdentity(GetCellText(sheet.GetRow(r)?.GetCell(2)));
            if (header.Contains("APELLIDOS") || header.Contains("NOMBRES")) { nameHeaderRow = r; break; }
        }
        if (nameHeaderRow < 0) throw new InvalidOperationException("No se encontró la columna APELLIDOS/NOMBRES en C.");

        string fileSubject = Path.GetFileNameWithoutExtension(source.Path);
        var ws = _generalMasterInfo!.Worksheet;
        int subjectStartColumn = -1;
        string subjectName = "";
        int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 1;
        for (int c = 4; c <= lastCol; c += 8)
        {
            string candidate = ws.Cell(6, c).GetString().Trim();
            if (string.IsNullOrWhiteSpace(candidate)) candidate = ws.Cell(6, c).Value.ToString().Trim();
            if (string.IsNullOrWhiteSpace(candidate)) continue;
            string a = NormalizeIdentity(candidate), b = NormalizeIdentity(fileSubject);
            if (a == b || a.Contains(b) || b.Contains(a) || EditDistance(a, b) <= 2)
            { subjectStartColumn = c; subjectName = candidate; break; }
        }
        if (subjectStartColumn < 0) throw new InvalidOperationException("No se pudo identificar la asignatura en el archivo maestro. Verifica el nombre del archivo.");
        int targetColumn = subjectStartColumn + trimester - 1;
        int dataStart = nameHeaderRow + 1;
        var masterGroups = _generalMasterInfo.StudentRows.GroupBy(x => FirstSurnameKey(x.Name), StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);
        var matchedRows = new HashSet<int>();
        var resultRows = new List<MatchRow>();
        var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
        var formatter = new DataFormatter(CultureInfo.InvariantCulture);
        int valid = 0, studentsFound = 0;

        for (int r = dataStart; r <= sheet.LastRowNum; r++)
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;
            string sourceName = ReadCell(row.GetCell(2), formatter, evaluator);
            if (string.IsNullOrWhiteSpace(sourceName) || !LooksLikeStudentName(sourceName)) continue;
            studentsFound++;
            var gradeCell = row.GetCell(36);
            string gradeText = ReadCell(gradeCell, formatter, evaluator);
            bool gradeOk = TryReadGeneralGrade(gradeCell, formatter, evaluator, out double grade);
            if (!gradeOk)
            {
                resultRows.Add(new MatchRow { SourceFile = source.FileName, Matter = subjectName, SourceName = sourceName, Grade = string.IsNullOrWhiteSpace(gradeText) ? "—" : gradeText, Status = string.IsNullOrWhiteSpace(gradeText) ? "SIN NOTA · NO SE TRANSFIERE" : "NOTA CON ERROR / NO NUMÉRICA · NO SE TRANSFIERE" });
                continue;
            }
            if (!masterGroups.TryGetValue(FirstSurnameKey(sourceName), out var candidates))
            {
                resultRows.Add(new MatchRow { SourceFile = source.FileName, Matter = subjectName, SourceName = sourceName, Grade = grade.ToString("0.##", CultureInfo.InvariantCulture), Status = "ADICIONAL / NO ENCONTRADO EN MAESTRO" });
                continue;
            }
            var available = candidates.Where(x => !matchedRows.Contains(x.Row)).ToList();
            var chosen = ResolveStudent(sourceName, available, out string matchStatus);
            if (chosen == null)
            {
                resultRows.Add(new MatchRow { SourceFile = source.FileName, Matter = subjectName, SourceName = sourceName, Grade = grade.ToString("0.##", CultureInfo.InvariantCulture), Student = string.Join(" / ", available.Select(x => x.Name)), Status = available.Count == 0 ? "DUPLICADO EN ARCHIVO" : "COINCIDENCIA AMBIGUA · REVISAR" });
                continue;
            }
            matchedRows.Add(chosen.Row);
            var target = ws.Cell(chosen.Row, targetColumn);
            string existingValue = _generalPendingGrades.TryGetValue((chosen.Row, targetColumn), out double pendingGrade)
                ? pendingGrade.ToString("0.##", CultureInfo.InvariantCulture)
                : target.Value.ToString();
            resultRows.Add(new MatchRow { SourceFile = source.FileName, Matter = subjectName, Student = chosen.Name, SourceName = sourceName, Grade = grade.ToString("0.##", CultureInfo.InvariantCulture), Existing = existingValue, MasterRow = chosen.Row, TargetColumn = targetColumn, Status = matchStatus });
            valid++;
        }
        if (studentsFound == 0) throw new InvalidOperationException("La hoja del trimestre no contiene estudiantes.");
        if (valid == 0) throw new InvalidOperationException("La hoja seleccionada no contiene ninguna calificación numérica válida en TOTAL TRIMESTRE (100%). No se transferirá nada.");
        foreach (var student in _generalMasterInfo.StudentRows.Where(x => !matchedRows.Contains(x.Row)))
            resultRows.Add(new MatchRow { SourceFile = source.FileName, Matter = subjectName, Student = student.Name, Status = "FALTANTE EN ARCHIVO DEL DOCENTE", MasterRow = student.Row, TargetColumn = targetColumn });
        return new GeneralDetectedSource(sheet.SheetName, subjectName, resultRows, valid);
    }

    private static bool TryReadGeneralGrade(ICell? cell, DataFormatter formatter, IFormulaEvaluator evaluator, out double grade)
    {
        grade = 0;
        if (cell == null || cell.CellType == CellType.Blank) return false;
        try
        {
            string value = formatter.FormatCellValue(cell, evaluator).Trim();
            if (string.IsNullOrWhiteSpace(value) || value.StartsWith("#", StringComparison.Ordinal)) return false;
            if (!double.TryParse(value.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out grade)) return false;
            return grade >= 0 && grade <= 10;
        }
        catch
        {
            try
            {
                if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric)
                { grade = cell.NumericCellValue; return grade >= 0 && grade <= 10; }
            }
            catch { }
            return false;
        }
    }

    private void ApplyGeneral_Click(object sender, RoutedEventArgs e)
    {
        if (!_generalAnalyzed || _generalMasterBook == null || _generalMasterInfo == null)
        { MessageBox.Show("Primero analiza los archivos."); return; }
        if (_generalAnalyzedTrimester != SelectedGeneralTrimester)
        {
            _generalAnalyzed = false;
            MessageBox.Show("Cambiaste el trimestre después del análisis. Pulsa 'Analizar archivos' nuevamente para calcular las columnas correctas.", "Debes volver a analizar", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        int applied = 0, skipped = 0;
        foreach (var item in _generalRows.Where(x => x.MasterRow > 0 && x.TargetColumn > 0))
        {
            if (!item.Status.StartsWith("COINCIDENCIA", StringComparison.Ordinal) || !TryParseGrade(item.Grade, out double grade) || grade < 0 || grade > 10) { skipped++; continue; }
            var cell = _generalMasterInfo.Worksheet.Cell(item.MasterRow, item.TargetColumn);
            string old = cell.Value.ToString();
            if (!string.IsNullOrWhiteSpace(old))
            {
                var answer = MessageBox.Show($"Asignatura: {item.Matter}\nTrimestre: {SelectedGeneralTrimester}\nEstudiante: {item.Student}\nNota actual: {old}\nNueva nota: {grade}\n\n¿Deseas reemplazarla?", "Nota existente", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                if (answer == MessageBoxResult.Cancel) break;
                if (answer == MessageBoxResult.No) { skipped++; continue; }
            }
            cell.Value = grade;
            _generalPendingGrades[(item.MasterRow, item.TargetColumn)] = grade;
            item.Existing = grade.ToString(CultureInfo.InvariantCulture);
            item.Status = "NOTA APLICADA";
            applied++;
        }
        GeneralResultsGrid.Items.Refresh();
        if (applied > 0)
        {
            GeneralStatusText.Text = "Notas aplicadas correctamente.";
            MessageBox.Show("Notas aplicadas correctamente.", "Johana Software", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else
        {
            GeneralStatusText.Text = "No se aplicaron notas.";
            MessageBox.Show("No se aplicó ninguna nota. Verifica que existan coincidencias y calificaciones válidas.", "Sin cambios", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void SaveGeneral_Click(object sender, RoutedEventArgs e)
    {
        if (_generalMasterBook == null || string.IsNullOrWhiteSpace(_generalMasterPath))
        {
            MessageBox.Show("Carga primero el Excel maestro.");
            return;
        }

        // Guardar el conjunto completo de notas aplicadas en todos los trimestres
        // de esta sesión, no solamente las filas visibles del último análisis.
        var pending = _generalPendingGrades.ToList();
        if (pending.Count == 0)
        {
            MessageBox.Show("No hay notas aplicadas para guardar.", "Sin cambios", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx",
            FileName = "Maestro_notas_generales_actualizado.xlsx",
            AddExtension = true,
            DefaultExt = ".xlsx"
        };
        if (dialog.ShowDialog() != true) return;
        if (string.Equals(Path.GetFullPath(dialog.FileName), Path.GetFullPath(_generalMasterPath), StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("Selecciona una ubicación diferente para guardar el archivo.", "Ubicación de guardado", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? worksheet = null;
        try
        {
            // Copiar el archivo original sin reconstruirlo: conserva imágenes, impresión,
            // saltos de página, márgenes, áreas de impresión, fórmulas y demás elementos.
            File.Copy(_generalMasterPath, dialog.FileName, true);

            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("No se encontró Microsoft Excel instalado. Para conservar exactamente el formato del maestro, esta operación requiere Excel de escritorio.");

            excel = Activator.CreateInstance(excelType);
            excel.Visible = false;
            excel.DisplayAlerts = false;
            workbook = excel.Workbooks.Open(Path.GetFullPath(dialog.FileName), ReadOnly: false);
            worksheet = workbook.Worksheets["INGRESO-NOTAS"];

            foreach (var item in pending)
            {
                if (item.Value < 0 || item.Value > 10) continue;
                worksheet.Cells[item.Key.Row, item.Key.Column].Value2 = item.Value;
            }

            workbook.Save();
            workbook.Close(true);
            workbook = null;
            excel.Quit();
            excel = null;

            // El libro en memoria ya contiene las notas aplicadas. Se actualiza la ruta
            // para que el siguiente guardado parta del archivo recién generado.
            _generalMasterPath = dialog.FileName;
            GeneralMasterPath.Text = dialog.FileName;
            GeneralStatusText.Text = string.Empty;
            MessageBox.Show("Archivo descargado con éxito");
        }
        catch (Exception ex)
        {
            try { if (workbook != null) workbook.Close(false); } catch { }
            try { if (excel != null) excel.Quit(); } catch { }
            try { if (File.Exists(dialog.FileName)) File.Delete(dialog.FileName); } catch { }
            MessageBox.Show("No se pudo guardar el archivo.\n\n" + ex.Message, "Error al guardar", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            try { if (worksheet != null) Marshal.FinalReleaseComObject(worksheet); } catch { }
            try { if (workbook != null) Marshal.FinalReleaseComObject(workbook); } catch { }
            try { if (excel != null) Marshal.FinalReleaseComObject(excel); } catch { }
        }
    }

    private sealed record GeneralDetectedSource(string SheetName, string MatterName, List<MatchRow> Rows, int ValidGradeCount);

    private void ChooseMaster_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel maestro (*.xlsx)|*.xlsx" };
        if (dialog.ShowDialog() != true) return;

        try
        {
            var newBook = new XLWorkbook(dialog.FileName);
            var info = DetectMaster(newBook);
            _masterBook?.Dispose();
            _masterBook = newBook;
            _masterPath = dialog.FileName;
            _masterInfo = info;
            MasterPath.Text = _masterPath;
            _rows.Clear();
            _analyzed = false;

            StatusText.Text = $"Maestro cargado: {info.StudentRows.Count} estudiantes. Hoja: {info.Worksheet.Name}.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"No se pudo leer el Excel maestro.\n\n{ex.Message}",
                "Error al cargar", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ChooseSources_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Archivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls",
            Multiselect = true
        };
        if (dialog.ShowDialog() != true) return;

        int added = 0;
        foreach (var path in dialog.FileNames)
        {
            if (_sources.Any(s => string.Equals(s.Path, path, StringComparison.OrdinalIgnoreCase)))
                continue;

            _sources.Add(new SourceItem
            {
                Path = path,
                FileName = Path.GetFileName(path),
                SheetName = "Pendiente de analizar",
                MatterName = "Automática",
                DetectionStatus = "En espera"
            });
            added++;
        }

        _analyzed = false;
        StatusText.Text = $"Se agregaron {added} archivo(s). Pulsa 'Analizar materias'.";
    }

    private void RemoveSources_Click(object sender, RoutedEventArgs e)
    {
        var selected = SourcesList.SelectedItems.Cast<SourceItem>().ToList();
        foreach (var item in selected) _sources.Remove(item);
        _rows.Clear();
        _analyzed = false;
        StatusText.Text = $"Archivos retirados: {selected.Count}.";
    }

    private void Analyze_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_masterBook == null || _masterInfo == null)
                throw new InvalidOperationException("Primero selecciona el Excel maestro.");
            if (_sources.Count == 0)
                throw new InvalidOperationException("Agrega al menos un Excel de materia.");

            _rows.Clear();
            _analyzed = false;
            int totalMatches = 0, totalIssues = 0;

            foreach (var source in _sources)
            {
                try
                {
                    var detected = AnalyzeSource(source);
                    source.SheetName = detected.SheetName;
                    source.MatterName = detected.MatterName;
                    source.DetectionStatus = $"Detectado · {detected.Matches} coincidencias";

                    foreach (var row in detected.Rows) _rows.Add(row);
                    totalMatches += detected.Matches;
                    totalIssues += detected.Rows.Count(r =>
                        !r.Status.StartsWith("COINCIDENCIA", StringComparison.Ordinal));

                }
                catch (Exception ex)
                {
                    source.DetectionStatus = "Error: " + ex.Message;
                    _rows.Add(new MatchRow
                    {
                        SourceFile = source.FileName,
                        Matter = "—",
                        Status = "ERROR DE LECTURA: " + ex.Message
                    });
                    totalIssues++;
                }
            }

            SourcesList.Items.Refresh();
            _analyzed = true;
            StatusText.Text = $"Análisis terminado. Coincidencias: {totalMatches}. Observaciones: {totalIssues}. Revisa antes de aplicar notas.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "No se pudo analizar",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private DetectedSource AnalyzeSource(SourceItem source)
    {
        using var workbook = OpenSource(source.Path);
        var masterGroups = _masterInfo!.StudentRows
            .GroupBy(x => FirstSurnameKey(x.Name), StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);

        var best = DetectSourceSheet(workbook, masterGroups.Keys.ToHashSet(StringComparer.Ordinal));

        // No reutilizar configuraciones antiguas de columnas: pudieron guardar
        // la columna "No." como si fuera la nota. Detectar siempre desde el Excel actual.
        var matter = FindDestinationMatter(Path.GetFileNameWithoutExtension(source.Path));

        if (matter.TargetColumn <= 0)
            throw new InvalidOperationException("No se pudo identificar la materia destino en el maestro. Revisa el nombre del archivo y los encabezados.");

        var matchedMasterRows = new HashSet<int>();
        var resultRows = new List<MatchRow>();
        var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
        var formatter = new DataFormatter(CultureInfo.InvariantCulture);

        int matches = 0;
        for (int r = best.StartRow; r <= best.Sheet.LastRowNum; r++)
        {
            var sourceName = GetCellText(best.Sheet.GetRow(r)?.GetCell(best.NameColumn));
            if (string.IsNullOrWhiteSpace(sourceName)) continue;

            string surnameKey = FirstSurnameKey(sourceName);
            if (string.IsNullOrWhiteSpace(surnameKey)) continue;

            var gradeCell = best.Sheet.GetRow(r)?.GetCell(best.GradeColumn);
            string grade = ReadCell(gradeCell, formatter, evaluator);

            if (!masterGroups.TryGetValue(surnameKey, out var candidates))
            {
                resultRows.Add(new MatchRow
                {
                    SourceFile = source.FileName,
                    Matter = matter.Name,
                    SourceName = sourceName,
                    Grade = grade,
                    Status = "ADICIONAL / PRIMER APELLIDO NO ENCONTRADO"
                });
                continue;
            }

            var available = candidates.Where(c => !matchedMasterRows.Contains(c.Row)).ToList();
            var chosen = ResolveStudent(sourceName, available, out string matchStatus);

            if (chosen == null)
            {
                resultRows.Add(new MatchRow
                {
                    SourceFile = source.FileName,
                    Matter = matter.Name,
                    SourceName = sourceName,
                    Grade = grade,
                    Student = string.Join(" / ", available.Select(x => x.Name)),
                    Status = available.Count == 0
                        ? "DUPLICADO EN ARCHIVO DE MATERIA"
                        : "COINCIDENCIA AMBIGUA · REVISAR"
                });
                continue;
            }

            matchedMasterRows.Add(chosen.Row);
            var targetCell = _masterInfo.Worksheet.Cell(chosen.Row, matter.TargetColumn);

            resultRows.Add(new MatchRow
            {
                SourceFile = source.FileName,
                Matter = matter.Name,
                Student = chosen.Name,
                SourceName = sourceName,
                Grade = grade,
                Existing = targetCell.Value.ToString(),
                MasterRow = chosen.Row,
                TargetColumn = matter.TargetColumn,
                Status = matchStatus
            });
            matches++;
        }

        foreach (var student in _masterInfo.StudentRows.Where(s => !matchedMasterRows.Contains(s.Row)))
        {
            resultRows.Add(new MatchRow
            {
                SourceFile = source.FileName,
                Matter = matter.Name,
                Student = student.Name,
                Status = "FALTANTE EN MATERIA",
                MasterRow = student.Row,
                TargetColumn = matter.TargetColumn
            });
        }

        return new DetectedSource(source.Path, best.Sheet.SheetName, matter.Name,
            matter.TargetColumn, best.NameColumn + 1, best.GradeColumn + 1,
            best.StartRow + 1, resultRows, matches);
    }

    // Primero exige el mismo primer apellido. Si ese apellido está repetido,
    // compara el nombre completo y solo acepta una coincidencia inequívoca.
    private static StudentRow? ResolveStudent(string sourceName, List<StudentRow> candidates, out string status)
    {
        status = "COINCIDENCIA POR PRIMER APELLIDO";

        if (candidates.Count == 0) return null;
        if (candidates.Count == 1)
        {
            status = NormalizeIdentity(candidates[0].Name) == NormalizeIdentity(sourceName)
                ? "COINCIDENCIA EXACTA"
                : "COINCIDENCIA POR PRIMER APELLIDO · REVISAR NOMBRE";
            return candidates[0];
        }

        string sourceFull = NormalizeIdentity(sourceName);
        var exact = candidates.Where(c => NormalizeIdentity(c.Name) == sourceFull).ToList();
        if (exact.Count == 1)
        {
            status = "COINCIDENCIA EXACTA · APELLIDO REPETIDO";
            return exact[0];
        }

        var ranked = candidates
            .Select(c => new { Student = c, Score = NameSimilarity(sourceFull, NormalizeIdentity(c.Name)) })
            .OrderByDescending(x => x.Score)
            .ToList();

        if (ranked.Count == 1 || (ranked[0].Score >= 0.90 &&
            ranked[0].Score - ranked[1].Score >= 0.12))
        {
            status = "COINCIDENCIA PROBABLE · CONFIRMAR";
            return ranked[0].Student;
        }

        return null;
    }

    private static double NameSimilarity(string a, string b)
    {
        if (a.Length == 0 || b.Length == 0) return 0;
        int distance = EditDistance(a, b);
        return 1.0 - (double)distance / Math.Max(a.Length, b.Length);
    }

    private SourceDetection DetectSourceSheet(IWorkbook workbook, HashSet<string> masterKeys)
    {
        SourceDetection? best = null;

        foreach (var sheet in workbook)
        {
            int headerRow = FindHeaderRow(sheet, "APELLIDOS");
            int nameCol = headerRow >= 0
                ? FindHeaderColumn(sheet.GetRow(headerRow),
                    new[] { "APELLIDOS Y NOMBRES", "NOMBRES Y APELLIDOS", "ESTUDIANTE", "ALUMNO" })
                : -1;

            if (nameCol < 0)
                nameCol = FindBestNameColumn(sheet, masterKeys, out headerRow);
            if (nameCol < 0) continue;

            int start = Math.Max(headerRow + 1, 0);
            while (start <= sheet.LastRowNum &&
                   !masterKeys.Contains(FirstSurnameKey(GetCellText(sheet.GetRow(start)?.GetCell(nameCol)))))
                start++;

            if (start > sheet.LastRowNum) continue;

            int matchCount = 0;
            for (int r = start; r <= sheet.LastRowNum; r++)
            {
                if (masterKeys.Contains(FirstSurnameKey(GetCellText(sheet.GetRow(r)?.GetCell(nameCol)))))
                    matchCount++;
            }

            int gradeCol = FindGradeColumn(sheet, nameCol, start);
            if (gradeCol < 0) continue;

            var candidate = new SourceDetection(sheet, nameCol, gradeCol, start, matchCount);
            if (best == null || candidate.MatchCount > best.MatchCount)
                best = candidate;
        }

        return best ?? throw new InvalidOperationException("No se detectaron columnas de estudiantes y calificaciones.");
    }

    private int FindBestNameColumn(ISheet sheet, HashSet<string> masterKeys, out int headerRow)
    {
        headerRow = 0;
        int bestCol = -1, bestCount = 0, bestStart = 0;
        int maxCol = Math.Min((int)(sheet.GetRow(0)?.LastCellNum ?? 12), 30);

        for (int c = 0; c < maxCol; c++)
        {
            int first = -1, count = 0;
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                string value = GetCellText(sheet.GetRow(r)?.GetCell(c));
                if (masterKeys.Contains(FirstSurnameKey(value)))
                {
                    if (first < 0) first = r;
                    count++;
                }
            }

            if (count > bestCount)
            {
                bestCount = count;
                bestCol = c;
                bestStart = first;
            }
        }

        headerRow = Math.Max(0, bestStart - 1);
        return bestCol;
    }

    // En los formatos de materias usados por la institución, la nota de la asignatura
    // está en la primera columna numérica situada A LA DERECHA del nombre.
    // La columna No. está a la izquierda y las columnas CUANTITATIVA/CUALITATIVA
    // suelen contener fórmulas de resumen, por lo que nunca deben usarse como nota.
    private static int FindGradeColumn(ISheet sheet, int nameCol, int startRow)
    {
        int lastCol = 0;
        for (int r = startRow; r <= Math.Min(sheet.LastRowNum, startRow + 80); r++)
        {
            var row = sheet.GetRow(r);
            if (row != null) lastCol = Math.Max(lastCol, row.LastCellNum);
        }

        if (lastCol <= nameCol + 1) return -1;

        int bestCol = -1;
        double bestScore = double.MinValue;

        // Buscar únicamente después de la columna de nombres. Así se excluye
        // expresamente la columna de numeración de estudiantes.
        for (int c = nameCol + 1; c < Math.Min(lastCol, nameCol + 8); c++)
        {
            int numericCount = 0;
            int inRangeCount = 0;
            int nonEmptyCount = 0;
            int formulaCount = 0;

            for (int r = startRow; r <= Math.Min(sheet.LastRowNum, startRow + 80); r++)
            {
                var cell = sheet.GetRow(r)?.GetCell(c);
                if (cell == null || cell.CellType == CellType.Blank) continue;

                // No contar fórmulas de resumen como notas de la asignatura.
                if (cell.CellType == CellType.Formula)
                {
                    formulaCount++;
                    continue;
                }

                if (cell.CellType == CellType.String &&
                    string.IsNullOrWhiteSpace(cell.StringCellValue))
                    continue;

                nonEmptyCount++;
                if (TryReadNumericCell(cell, out double value))
                {
                    numericCount++;
                    if (value >= 0 && value <= 10) inRangeCount++;
                }
            }

            if (numericCount < 3 || nonEmptyCount == 0) continue;

            double inRangeRatio = (double)inRangeCount / numericCount;
            if (inRangeRatio < 0.80) continue;

            // Prioriza la columna más cercana al nombre; penaliza fórmulas.
            double score = (inRangeRatio * 100.0) + (numericCount * 0.1)
                           - ((c - nameCol - 1) * 15.0) - (formulaCount * 2.0);

            if (score > bestScore)
            {
                bestScore = score;
                bestCol = c;
            }
        }

        return bestCol;
    }

    private MatterDestination FindDestinationMatter(string fileBaseName)
    {
        var ws = _masterInfo!.Worksheet;
        string target = NormalizeIdentity(fileBaseName);
        var candidates = new List<(string Name, int Col, int Score)>();

        for (int r = _masterInfo.HeaderRow;
             r <= Math.Min(_masterInfo.HeaderRow + 3, ws.LastRowUsed()?.RowNumber() ?? _masterInfo.HeaderRow);
             r++)
        {
            int lastColumn = ws.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (int c = 1; c <= lastColumn; c++)
            {
                string value = ws.Cell(r, c).GetString().Trim();
                string normalized = NormalizeIdentity(value);
                if (normalized.Length < 3 || LooksLikeHeaderOrName(normalized)) continue;

                int distance = EditDistance(normalized, target);
                int score = Math.Max(normalized.Length, target.Length) - distance;
                if (normalized.Contains(target) || target.Contains(normalized)) score += 4;
                candidates.Add((value, c, score));
            }
        }

        var best = candidates.OrderByDescending(x => x.Score).FirstOrDefault();
        if (best.Score <= 0 || best.Col == _masterInfo.NameColumn ||
            EditDistance(NormalizeIdentity(best.Name), target) > 2)
            return new MatterDestination(fileBaseName, -1);

        return new MatterDestination(best.Name, best.Col);
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        if (!_analyzed || _masterBook == null || _masterInfo == null)
        {
            MessageBox.Show("Primero analiza los archivos.");
            return;
        }

        int changed = 0, skipped = 0;

        foreach (var item in _rows.Where(x => x.MasterRow > 0 && x.TargetColumn > 0))
        {
            // Nunca aplicar filas ambiguas, adicionales, faltantes ni duplicadas.
            if (item.Status.Contains("AMBIGUA", StringComparison.OrdinalIgnoreCase) ||
                !item.Status.StartsWith("COINCIDENCIA", StringComparison.Ordinal))
            {
                skipped++;
                continue;
            }

            // Si la coincidencia es probable o el nombre difiere, pedir aprobación explícita.
            if (item.Status.Contains("REVISAR", StringComparison.OrdinalIgnoreCase) ||
                item.Status.Contains("CONFIRMAR", StringComparison.OrdinalIgnoreCase))
            {
                var confirm = MessageBox.Show(
                    $"Confirma que ambos registros corresponden al mismo estudiante.\n\n" +
                    $"Nombre del maestro: {item.Student}\n" +
                    $"Nombre en la materia: {item.SourceName}\n" +
                    $"Nota detectada: {item.Grade}\n\n" +
                    "El nombre del maestro no se modificará. ¿Deseas aplicar esta nota?",
                    "Confirmar coincidencia", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);

                if (confirm == MessageBoxResult.Cancel) break;
                if (confirm == MessageBoxResult.No) { skipped++; continue; }
            }

            if (!TryParseGrade(item.Grade, out double grade) || grade < 0 || grade > 10)
            {
                skipped++;
                item.Status = "NOTA INVÁLIDA (DEBE ESTAR ENTRE 0 Y 10)";
                continue;
            }

            var cell = _masterInfo.Worksheet.Cell(item.MasterRow, item.TargetColumn);
            string old = cell.Value.ToString();

            if (!string.IsNullOrWhiteSpace(old))
            {
                var answer = MessageBox.Show(
                    $"Materia: {item.Matter}\nEstudiante: {item.Student}\nNota actual: {old}\nNueva nota: {grade}\n\n¿Deseas reemplazarla?",
                    "Nota existente", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                if (answer == MessageBoxResult.Cancel) break;
                if (answer == MessageBoxResult.No) { skipped++; continue; }
            }

            cell.Value = grade;
            changed++;
            item.Existing = grade.ToString(CultureInfo.InvariantCulture);
            item.Status = "NOTA APLICADA";
        }

        ResultsGrid.Items.Refresh();
        StatusText.Text = $"Notas aplicadas: {changed}. Omitidas: {skipped}. Guarda el maestro para conservar los cambios.";
    }

    private void SortMaster_Click(object sender, RoutedEventArgs e)
    {
        if (_masterBook == null || _masterInfo == null)
        {
            MessageBox.Show("Carga primero el Excel maestro.");
            return;
        }

        var current = _masterInfo.StudentRows.OrderBy(x => x.Row).ToList();
        var sorted = current.OrderBy(x => x.Name, SpanishStudentNameComparer.Instance).ToList();

        bool alreadySorted = current.Select(x => x.Row)
            .SequenceEqual(sorted.Select(x => x.Row));

        if (alreadySorted)
        {
            MessageBox.Show("La nómina ya está ordenada alfabéticamente por primer apellido y, en caso de empate, por los nombres y apellidos restantes.",
                "Orden alfabético", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var newPositions = sorted.Select((student, index) => new { student.Row, Position = index + 1 })
            .ToDictionary(x => x.Row, x => x.Position);

        var differences = current.Select((student, index) => new
        {
            Student = student,
            CurrentPosition = index + 1,
            NewPosition = newPositions[student.Row]
        }).Where(x => x.CurrentPosition != x.NewPosition).ToList();

        string detail = string.Join(Environment.NewLine, differences.Take(30)
            .Select(x => $"Fila Excel {x.Student.Row}: {x.Student.Name} | posición actual {x.CurrentPosition} → propuesta {x.NewPosition}"));

        if (differences.Count > 30)
            detail += Environment.NewLine + $"... y {differences.Count - 30} estudiante(s) más.";

        string message =
            $"Se encontraron {differences.Count} estudiante(s) cuya posición cambiaría.{Environment.NewLine}" +
            "Orden: primer apellido; en caso de empate, los demás apellidos y nombres. Se respeta la Ñ española." +
            Environment.NewLine + Environment.NewLine + detail +
            Environment.NewLine + Environment.NewLine +
            "Se moverán las filas completas dentro del rango de estudiantes. Los nombres no se corregirán." +
            Environment.NewLine + "¿Apruebas el ordenamiento?";

        if (MessageBox.Show(message, "Revisión del orden alfabético",
            MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;

        try
        {
            ReorderStudentRows(_masterInfo, sorted);
            _masterInfo = DetectMaster(_masterBook);
            _analyzed = false;
            _rows.Clear();
            StatusText.Text = "Nómina ordenada. Vuelve a analizar las materias antes de aplicar notas.";
            MessageBox.Show("Ordenamiento aplicado. Los nombres originales se conservaron.",
                "Completado", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("No se pudo ordenar la nómina: " + ex.Message,
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static void ReorderStudentRows(MasterInfo info, List<StudentRow> sorted)
    {
        var ws = info.Worksheet;
        int firstCol = 1;
        int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? info.NameColumn;
        var rowSlots = info.StudentRows.Select(x => x.Row).OrderBy(x => x).ToList();

        // Primero se copian todas las filas a una zona temporal de la MISMA hoja.
        // CopyTo conserva valores, fórmulas y formatos; al copiar una fórmula a otra
        // fila, ClosedXML ajusta sus referencias relativas como hace Excel al moverla.
        int tempStart = Math.Max(ws.LastRowUsed()?.RowNumber() ?? rowSlots.Max(), rowSlots.Max()) + 2;
        var tempRows = new List<int>(sorted.Count);
        var rowHeights = new List<double>(sorted.Count);

        for (int i = 0; i < sorted.Count; i++)
        {
            int tempRow = tempStart + i;
            int sourceRow = sorted[i].Row;
            ws.Range(sourceRow, firstCol, sourceRow, lastCol)
                .CopyTo(ws.Range(tempRow, firstCol, tempRow, lastCol));
            ws.Row(tempRow).Height = ws.Row(sourceRow).Height;
            tempRows.Add(tempRow);
            rowHeights.Add(ws.Row(sourceRow).Height);
        }

        // Ya existe una copia íntegra de todas las filas originales, por lo que ahora
        // se pueden reemplazar los espacios de estudiantes sin perder datos ni fórmulas.
        for (int i = 0; i < rowSlots.Count; i++)
        {
            int destinationRow = rowSlots[i];
            int stagedRow = tempRows[i];
            ws.Range(stagedRow, firstCol, stagedRow, lastCol)
                .CopyTo(ws.Range(destinationRow, firstCol, destinationRow, lastCol));
            ws.Row(destinationRow).Height = rowHeights[i];
        }

        // Limpiar únicamente la zona temporal creada por este procedimiento.
        for (int i = 0; i < tempRows.Count; i++)
            ws.Range(tempRows[i], firstCol, tempRows[i], lastCol).Clear(XLClearOptions.All);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (_masterBook == null)
        {
            MessageBox.Show("Carga primero el Excel maestro.");
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx",
            FileName = "Maestro_actualizado.xlsx"
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            // Guardar un único archivo: no crear copias _respaldo.xlsx automáticamente.
            _masterBook.SaveAs(dialog.FileName);
            StatusText.Text = "Maestro guardado.";
            MessageBox.Show("Archivo guardado correctamente.", "Finalizado",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error al guardar",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private MasterInfo DetectMaster(XLWorkbook book)
    {
        foreach (var ws in book.Worksheets)
        {
            int headerRow = 0, nameCol = 0;
            int maxRow = Math.Min(ws.LastRowUsed()?.RowNumber() ?? 1, 30);
            int maxCol = Math.Min(ws.LastColumnUsed()?.ColumnNumber() ?? 1, 30);

            for (int r = 1; r <= maxRow; r++)
            {
                for (int c = 1; c <= maxCol; c++)
                {
                    if (NormalizeIdentity(ws.Cell(r, c).GetString()).Contains("APELLIDOSYNOMBRES"))
                    {
                        headerRow = r;
                        nameCol = c;
                        break;
                    }
                }
                if (headerRow > 0) break;
            }

            if (headerRow == 0) continue;

            var students = new List<StudentRow>();
            for (int r = headerRow + 1; r <= (ws.LastRowUsed()?.RowNumber() ?? headerRow); r++)
            {
                string name = ws.Cell(r, nameCol).GetString().Trim();
                if (LooksLikeStudentName(name))
                    students.Add(new StudentRow(r, name));
            }

            if (students.Count > 0)
                return new MasterInfo(ws, headerRow, nameCol, students);
        }

        throw new InvalidOperationException("No se encontró el encabezado 'APELLIDOS Y NOMBRES' en el Excel maestro.");
    }

    private static int FindHeaderRow(ISheet sheet, string keyword)
    {
        for (int r = 0; r <= Math.Min(sheet.LastRowNum, 20); r++)
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;

            for (int c = 0; c < Math.Min((int)row.LastCellNum, 30); c++)
            {
                if (NormalizeIdentity(GetCellText(row.GetCell(c)))
                    .Contains(NormalizeIdentity(keyword)))
                    return r;
            }
        }
        return -1;
    }

    private static int FindHeaderColumn(IRow? row, string[] options)
    {
        if (row == null) return -1;
        for (int c = 0; c < row.LastCellNum; c++)
        {
            string value = NormalizeIdentity(GetCellText(row.GetCell(c)));
            if (options.Any(o => value.Contains(NormalizeIdentity(o))))
                return c;
        }
        return -1;
    }

    private static bool LooksLikeStudentName(string value) =>
        value.Trim().Length >= 5 && value.Any(char.IsLetter) &&
        !NormalizeIdentity(value).Contains("APELLIDOSYNOMBRES");

    private static bool LooksLikeHeaderOrName(string normalized) =>
        normalized.Contains("APELLIDOS") || normalized.Contains("CALIFICACION") ||
        normalized.Contains("CUANTITATIVA") || normalized.Contains("CUALITATIVA") ||
        normalized.Contains("NOMBRES");

    private static string GetCellText(ICell? cell) => cell?.ToString()?.Trim() ?? "";

    private static string ReadCell(ICell? cell, DataFormatter formatter, IFormulaEvaluator evaluator)
    {
        if (cell == null || cell.CellType == CellType.Blank) return "";
        try
        {
            return formatter.FormatCellValue(cell, evaluator).Trim();
        }
        catch
        {
            return cell.ToString()?.Trim() ?? "";
        }
    }

    private static bool TryReadNumericCell(ICell cell, out double value)
    {
        value = 0;
        try
        {
            if (cell.CellType == CellType.Numeric)
            {
                value = cell.NumericCellValue;
                return true;
            }
            if (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric)
            {
                value = cell.NumericCellValue;
                return true;
            }
            return double.TryParse(cell.ToString()?.Replace(',', '.'),
                NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }
        catch { return false; }
    }

    private static bool TryParseGrade(string text, out double grade) =>
        double.TryParse(text.Replace(',', '.'), NumberStyles.Any,
            CultureInfo.InvariantCulture, out grade);

    private static IWorkbook OpenSource(string path)
    {
        using var stream = File.OpenRead(path);
        return Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase)
            ? new HSSFWorkbook(stream)
            : new XSSFWorkbook(stream);
    }

    private static string FirstSurnameKey(string fullName)
    {
        if (string.IsNullOrWhiteSpace(fullName)) return "";
        string firstToken = fullName.Trim()
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault() ?? "";
        return NormalizeIdentity(firstToken);
    }

    // Para identidad se ignoran tildes y diferencias de espacios.
    private static string NormalizeIdentity(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        string decomposed = value.Trim().Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder();

        foreach (char c in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark)
                continue;
            if (!char.IsWhiteSpace(c))
                builder.Append(char.ToUpperInvariant(c));
        }

        return builder.ToString().Normalize(NormalizationForm.FormC);
    }

    // Para ordenar se eliminan tildes de vocales, pero se conserva Ñ como letra propia.
    private static string NormalizeForSpanishSort(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        string decomposed = value.Trim().Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder();

        for (int i = 0; i < decomposed.Length; i++)
        {
            char c = decomposed[i];
            if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark)
            {
                // La virgulilla de la Ñ se conserva; los demás acentos se ignoran.
                if (builder.Length > 0 && builder[builder.Length - 1] == 'N' && c == '\u0303')
                    builder.Append('\u0303');
                continue;
            }
            if (!char.IsWhiteSpace(c))
                builder.Append(char.ToUpperInvariant(c));
        }

        return builder.ToString().Normalize(NormalizationForm.FormC);
    }

    private static int SpanishCharRank(char c)
    {
        if (c >= 'A' && c <= 'N') return c - 'A';
        if (c == 'Ñ') return 14;
        if (c >= 'O' && c <= 'Z') return c - 'A' + 1;
        return 100 + c;
    }

    private sealed class SpanishStudentNameComparer : IComparer<string>
    {
        public static readonly SpanishStudentNameComparer Instance = new();

        public int Compare(string? x, string? y)
        {
            string[] a = NormalizeForSpanishSort(x ?? "")
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);
            string[] b = NormalizeForSpanishSort(y ?? "")
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);

            int tokenCount = Math.Min(a.Length, b.Length);
            for (int t = 0; t < tokenCount; t++)
            {
                int charCount = Math.Min(a[t].Length, b[t].Length);
                for (int i = 0; i < charCount; i++)
                {
                    int cmp = SpanishCharRank(a[t][i]).CompareTo(SpanishCharRank(b[t][i]));
                    if (cmp != 0) return cmp;
                }

                int lengthCompare = a[t].Length.CompareTo(b[t].Length);
                if (lengthCompare != 0) return lengthCompare;
            }

            return a.Length.CompareTo(b.Length);
        }
    }

    private static int EditDistance(string a, string b)
    {
        int[,] d = new int[a.Length + 1, b.Length + 1];
        for (int i = 0; i <= a.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= b.Length; j++) d[0, j] = j;

        for (int i = 1; i <= a.Length; i++)
        {
            for (int j = 1; j <= b.Length; j++)
            {
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1));
            }
        }

        return d[a.Length, b.Length];
    }

    private sealed record MasterInfo(IXLWorksheet Worksheet, int HeaderRow, int NameColumn, List<StudentRow> StudentRows);
    private sealed record StudentRow(int Row, string Name);
    private sealed record SourceDetection(ISheet Sheet, int NameColumn, int GradeColumn, int StartRow, int MatchCount);
    private sealed record MatterDestination(string Name, int TargetColumn);
    private sealed record DetectedSource(string Path, string SheetName, string MatterName, int TargetColumn,
        int NameColumn, int GradeColumn, int StartRow, List<MatchRow> Rows, int Matches);

    public sealed class SourceItem
    {
        public string Path { get; set; } = "";
        public string FileName { get; set; } = "";
        public string SheetName { get; set; } = "";
        public string MatterName { get; set; } = "";
        public string DetectionStatus { get; set; } = "";
    }

    public sealed class MatchRow
    {
        public string SourceFile { get; set; } = "";
        public string Matter { get; set; } = "";
        public string Student { get; set; } = "";
        public string SourceName { get; set; } = "";
        public string Grade { get; set; } = "";
        public string Existing { get; set; } = "";
        public string Status { get; set; } = "";
        public int MasterRow { get; set; }
        public int TargetColumn { get; set; }
    }
}
