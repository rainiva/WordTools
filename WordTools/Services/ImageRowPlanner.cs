using System.Collections.Generic;

namespace WordTools.Services
{
    public enum ImageCellAvailability
    {
        Available,
        OverwriteImage,
        OverwriteText,
        Blocked,
        Merged
    }

    public sealed class ImageRowAvailability
    {
        public ImageRowAvailability(int rowIndex, ImageCellAvailability leftCell, ImageCellAvailability rightCell)
        {
            RowIndex = rowIndex;
            LeftCell = leftCell;
            RightCell = rightCell;
        }

        public int RowIndex { get; }

        public ImageCellAvailability LeftCell { get; }

        public ImageCellAvailability RightCell { get; }

        public bool HasMergedCell
        {
            get
            {
                return LeftCell == ImageCellAvailability.Merged || RightCell == ImageCellAvailability.Merged;
            }
        }
    }

    public static class ImageRowPlanner
    {
        public static bool CanHostSingleImage(ImageCellAvailability availability)
        {
            return availability == ImageCellAvailability.Available
                || availability == ImageCellAvailability.OverwriteImage
                || availability == ImageCellAvailability.OverwriteText;
        }

        public static bool RequiresOverwriteWarning(ImageCellAvailability availability)
        {
            return availability == ImageCellAvailability.OverwriteImage
                || availability == ImageCellAvailability.OverwriteText;
        }

        public static bool ShouldRetryCurrentImage(ImageCellAvailability availability)
        {
            return availability == ImageCellAvailability.Merged
                || availability == ImageCellAvailability.Blocked;
        }

        public static bool CanHostImagePair(ImageRowAvailability row)
        {
            return row != null
                && CanHostSingleImage(row.LeftCell)
                && CanHostSingleImage(row.RightCell);
        }

        public static int FindPreferredPairRow(
            ImageRowAvailability currentRow,
            IEnumerable<ImageRowAvailability> fallbackRows,
            int notFoundRow = -1)
        {
            if (CanHostImagePair(currentRow))
            {
                return currentRow.RowIndex;
            }

            if (fallbackRows == null)
            {
                return notFoundRow;
            }

            foreach (var row in fallbackRows)
            {
                if (CanHostImagePair(row))
                {
                    return row.RowIndex;
                }
            }

            return notFoundRow;
        }

        public static int FindNextPairRow(IReadOnlyList<ImageRowAvailability> rows, int startIndex = 0)
        {
            if (rows == null || startIndex < 0)
            {
                return -1;
            }

            for (int i = startIndex; i < rows.Count; i++)
            {
                if (CanHostImagePair(rows[i]))
                {
                    return i;
                }
            }

            return -1;
        }
    }
}
