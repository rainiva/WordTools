namespace WordTools.Services
{
    public enum ProgressButtonAction
    {
        None,
        CancelRequested,
        CloseRequested
    }

    public sealed class ProgressFormStateController
    {
        public bool IsCancelled { get; private set; }
        public bool IsCompleted { get; private set; }
        public bool IsButtonEnabled { get; private set; } = true;
        public string ButtonText { get; private set; } = "取消";

        public ProgressButtonAction HandleButtonClick()
        {
            if (IsCompleted)
            {
                return ProgressButtonAction.CloseRequested;
            }

            if (IsCancelled)
            {
                return ProgressButtonAction.None;
            }

            IsCancelled = true;
            IsButtonEnabled = false;
            ButtonText = "正在取消...";
            return ProgressButtonAction.CancelRequested;
        }

        public void MarkCompleted()
        {
            IsCompleted = true;
            IsButtonEnabled = true;
            ButtonText = "关闭";
        }
    }
}
