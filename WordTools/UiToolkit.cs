using System.Drawing;
using System.Windows.Forms;

namespace WordTools
{
    public static class UiToolkit
    {
        public enum ButtonStyle
        {
            Primary,
            Success,
            Default
        }

        public static Button CreateButton(string text, ButtonStyle style = ButtonStyle.Default)
        {
            var btn = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                Font = Theme.Fonts.Bold,
                Cursor = Cursors.Hand
            };

            ApplyButtonStyle(btn, style);

            btn.EnabledChanged += (sender, e) =>
            {
                var button = sender as Button;
                if (button != null)
                {
                    if (button.Enabled)
                    {
                        ApplyButtonStyle(button, style);
                        button.Cursor = Cursors.Hand;
                    }
                    else
                    {
                        button.BackColor = Theme.Colors.ButtonDisabled;
                        button.ForeColor = Theme.Colors.TextDisabled;
                        button.FlatAppearance.BorderColor = Theme.Colors.ButtonDisabledBorder;
                        button.FlatAppearance.BorderSize = 1;
                        button.Cursor = Cursors.Default;
                    }
                }
            };

            return btn;
        }

        private static void ApplyButtonStyle(Button btn, ButtonStyle style)
        {
            switch (style)
            {
                case ButtonStyle.Primary:
                    btn.BackColor = Theme.Colors.Primary;
                    btn.ForeColor = Color.White;
                    btn.FlatAppearance.BorderSize = 0;
                    break;
                case ButtonStyle.Success:
                    btn.BackColor = Theme.Colors.Success;
                    btn.ForeColor = Color.White;
                    btn.FlatAppearance.BorderSize = 0;
                    break;
                default:
                    btn.BackColor = Theme.Colors.ButtonDefault;
                    btn.ForeColor = Theme.Colors.Text;
                    btn.FlatAppearance.BorderColor = Theme.Colors.Border;
                    btn.FlatAppearance.BorderSize = 1;
                    break;
            }
        }

        public static Label CreateDivider(int width)
        {
            return new Label
            {
                Size = new Size(width, 1),
                BackColor = Theme.Colors.Border,
                BorderStyle = BorderStyle.None
            };
        }
    }
}
