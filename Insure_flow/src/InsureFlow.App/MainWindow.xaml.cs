using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using InsureFlow.App.ViewModels;
using InsureFlow.Core.Model;

namespace InsureFlow.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm = new();

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _vm;
        BuildPreviewColumns();
        // 파일을 exe에 끌어다 놓거나 인자로 넘기면 종류를 인식해 자동으로 채운다.
        Loaded += (_, _) => _vm.HandleDrop(Environment.GetCommandLineArgs().Skip(1).Where(File.Exists));
        DragOver += (_, e) =>
        {
            e.Effects = _vm.IsStep0 && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        };
        Drop += (_, e) =>
        {
            if (_vm.IsStep0 && e.Data.GetData(DataFormats.FileDrop) is string[] files) _vm.HandleDrop(files);
        };
    }

    /// <summary>미리보기 표: 상태 + 사번 + 성명 + 8개 입력 칸(상태별 색상, 음수 빨간색, 우측 정렬).</summary>
    private void BuildPreviewColumns()
    {
        var g = PreviewGrid;
        g.Columns.Add(new DataGridTemplateColumn
        {
            Header = "상태", Width = new DataGridLength(128), CellTemplate = (DataTemplate)FindResource("StatusChipTemplate"),
        });
        g.Columns.Add(new DataGridTextColumn { Header = "사번", Binding = new Binding("EmpNo"), Width = new DataGridLength(100), ElementStyle = TextStyle(null, null) });
        g.Columns.Add(new DataGridTextColumn { Header = "성명", Binding = new Binding("Name"), Width = new DataGridLength(76), ElementStyle = TextStyle(null, null) });

        var baseCell = (Style)FindResource(typeof(DataGridCell));
        for (var i = 0; i < FieldInfo.All.Length; i++)
        {
            var path = $"Cells[{i}]";
            var cell = new Style(typeof(DataGridCell), baseCell);
            cell.Triggers.Add(StateTrigger(path + ".State", "NeedsCheck", (Brush)FindResource("SupportSoftBrush")));
            cell.Triggers.Add(StateTrigger(path + ".State", "Empty", (Brush)FindResource("NeutralCellBrush")));
            g.Columns.Add(new DataGridTextColumn
            {
                Header = HeaderText(FieldInfo.Label(FieldInfo.All[i])),
                Binding = new Binding(path + ".Text"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                MinWidth = 76,
                CellStyle = cell,
                ElementStyle = TextStyle(path, null),
            });
        }
    }

    /// <summary>열 머리글이 좁을 때 잘리지 않도록 '…정산'은 두 줄로 나눈다.</summary>
    private static string HeaderText(string label) =>
        label.EndsWith("정산") ? label[..^2] + "\n정산" : label;

    private static DataTrigger StateTrigger(string binding, string value, Brush bg)
    {
        var t = new DataTrigger { Binding = new Binding(binding), Value = value };
        t.Setters.Add(new Setter(Control.BackgroundProperty, bg));
        return t;
    }

    private Style TextStyle(string? cellPath, Action<Style>? extra)
    {
        var s = new Style(typeof(TextBlock));
        s.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(6, 0, 6, 0)));
        if (cellPath != null)
        {
            s.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Right));
            var neg = new DataTrigger { Binding = new Binding(cellPath + ".IsNegative"), Value = true };
            neg.Setters.Add(new Setter(TextBlock.ForegroundProperty, (Brush)FindResource("DangerDarkBrush")));
            s.Triggers.Add(neg);
        }
        extra?.Invoke(s);
        return s;
    }
}
