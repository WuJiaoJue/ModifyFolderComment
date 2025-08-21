using System.Diagnostics.CodeAnalysis;

namespace ModifyFolderComment;

/// <summary>
/// 文件夹注释修改窗体
/// </summary>
public class FolderCommentForm : Form
{
    private readonly string _folderPath;
    private Label? _label;
    private TextBox? _textBox;
    private Button? _buttonOk;
    private Button? _buttonCancel;

    /// <summary>
    /// 初始化文件夹注释修改窗体
    /// </summary>
    /// <param name="folderPath">要修改注释的文件夹路径</param>
    public FolderCommentForm(string folderPath)
    {
        _folderPath = folderPath;
        Text = "修改备注";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(400, 140);

        InitComponents();
    }

    [AllowNull] public sealed override string Text
    {
        get => base.Text;
        set => base.Text = value;
    }

    private void InitComponents()
    {
        const int margin = 12;
        const int spacing = 8;

        _label = new Label
        {
            Text = "请输入备注：",
            Location = new Point(margin, margin),
            AutoSize = true
        };

        _textBox = new TextBox
        {
            Location = new Point(margin, _label.Bottom + spacing),
            Width = ClientSize.Width - margin * 2,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            TabIndex = 0
        };

        // 预加载 desktop.ini 中的 InfoTip
        var existing = DesktopIniManager.GetInfoTip(_folderPath);
        if (!string.IsNullOrWhiteSpace(existing))
        {
            _textBox.Text = existing;
        }

        _buttonOk = new Button
        {
            Text = "确定",
            AutoSize = true,
            UseVisualStyleBackColor = true,
            DialogResult = DialogResult.OK,
            TabIndex = 1
        };

        _buttonCancel = new Button
        {
            Text = "取消",
            AutoSize = true,
            UseVisualStyleBackColor = true,
            DialogResult = DialogResult.Cancel,
            TabIndex = 2
        };

        // 在OnLoad中设置按钮位置
        _buttonOk.Click += (_, _) =>
        {
            var comment = _textBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(comment))
            {
                DesktopIniManager.SetInfoTip(_folderPath, comment);
            }
            Close();
        };

        _buttonCancel.Click += (_, _) => Close();

        Controls.AddRange([_label, _textBox, _buttonOk, _buttonCancel]);
        AcceptButton = _buttonOk;
        CancelButton = _buttonCancel;
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        
        const int margin = 12;
        const int spacing = 8;
        
        // 设置按钮位置
        var buttonY = _textBox!.Bottom + spacing;
        
        _buttonCancel!.Location = new Point(ClientSize.Width - margin - _buttonCancel.Width, buttonY);
        _buttonOk!.Location = new Point(_buttonCancel.Left - spacing - _buttonOk.Width, buttonY);
        
        _textBox.Focus();
        _textBox.SelectAll();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        
        if (_textBox != null && _buttonOk != null && _buttonCancel != null)
        {
            const int margin = 12;
            const int spacing = 8;
            
            // 更新输入框宽度
            _textBox.Width = ClientSize.Width - margin * 2;
            
            // 更新按钮位置
            var buttonY = _textBox.Bottom + spacing;
            _buttonCancel.Location = new Point(ClientSize.Width - margin - _buttonCancel.Width, buttonY);
            _buttonOk.Location = new Point(_buttonCancel.Left - spacing - _buttonOk.Width, buttonY);
        }
    }
}