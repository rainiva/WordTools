namespace WordTools.Forms
{
    public enum InsertPhotosRequestMode
    {
        Folder,
        SelectedFiles,
    }

    public sealed class InsertPhotosRequest
    {
        public InsertPhotosRequestMode Mode { get; set; }
        public string FolderPath { get; set; }
        public string[] SelectedFiles { get; set; }
        public float MinHeight { get; set; }
        public bool NeedDescription { get; set; }
        public bool UseFileNameAsDescription { get; set; }
        public bool UseFolderNameAsDescription { get; set; }
        public bool IncludeRootImages { get; set; }
        public bool IncludeSubFolderImages { get; set; }
        public bool NeedAutoNumbering { get; set; }
        public int NumberAlignment { get; set; }
        public int NumberPosition { get; set; }
    }
}
