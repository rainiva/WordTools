using System;
using System.Collections.Generic;

namespace WordTools.Services
{
    public sealed class FloatingShapeAnchor
    {
        public FloatingShapeAnchor(int start, int end)
        {
            Start = start;
            End = end;
        }

        public int Start { get; }

        public int End { get; }
    }

    public sealed class FloatingShapeIndex
    {
        private readonly FloatingShapeAnchor[] _anchors;

        private FloatingShapeIndex(FloatingShapeAnchor[] anchors)
        {
            _anchors = anchors ?? new FloatingShapeAnchor[0];
        }

        public static FloatingShapeIndex Create(IEnumerable<FloatingShapeAnchor> anchors)
        {
            var list = anchors != null
                ? new List<FloatingShapeAnchor>(anchors)
                : new List<FloatingShapeAnchor>();

            list.Sort((left, right) => left.Start.CompareTo(right.Start));
            return new FloatingShapeIndex(list.ToArray());
        }

        public bool HasShapeInRange(int rangeStart, int rangeEnd)
        {
            if (_anchors.Length == 0 || rangeEnd < rangeStart)
            {
                return false;
            }

            int index = FindFirstCandidateIndex(rangeStart);
            for (int i = index; i < _anchors.Length; i++)
            {
                FloatingShapeAnchor anchor = _anchors[i];
                if (anchor.Start > rangeEnd)
                {
                    break;
                }

                if (anchor.Start >= rangeStart && anchor.End <= rangeEnd)
                {
                    return true;
                }
            }

            return false;
        }

        private int FindFirstCandidateIndex(int rangeStart)
        {
            int low = 0;
            int high = _anchors.Length - 1;
            int result = _anchors.Length;

            while (low <= high)
            {
                int mid = low + ((high - low) / 2);
                if (_anchors[mid].Start >= rangeStart)
                {
                    result = mid;
                    high = mid - 1;
                }
                else
                {
                    low = mid + 1;
                }
            }

            return result;
        }
    }

    public sealed class ImageInsertionBatchContext
    {
        private readonly Dictionary<int, ImageRowAvailability> _rowAvailabilityCache = new Dictionary<int, ImageRowAvailability>();
        private FloatingShapeIndex _floatingShapeIndex;

        public ImageInsertionBatchContext(InsertionPerformanceDiagnostics diagnostics = null)
        {
            Diagnostics = diagnostics ?? new InsertionPerformanceDiagnostics();
        }

        public InsertionPerformanceDiagnostics Diagnostics { get; }

        public bool TryGetCachedRowAvailability(int rowIndex, out ImageRowAvailability rowAvailability)
        {
            return _rowAvailabilityCache.TryGetValue(rowIndex, out rowAvailability);
        }

        public void CacheRowAvailability(ImageRowAvailability rowAvailability)
        {
            if (rowAvailability == null)
            {
                return;
            }

            _rowAvailabilityCache[rowAvailability.RowIndex] = rowAvailability;
        }

        public void ClearRowAvailability()
        {
            _rowAvailabilityCache.Clear();
        }

        public FloatingShapeIndex GetOrCreateFloatingShapeIndex(Func<FloatingShapeIndex> factory)
        {
            if (_floatingShapeIndex == null)
            {
                _floatingShapeIndex = factory != null
                    ? factory()
                    : FloatingShapeIndex.Create(null);
            }

            return _floatingShapeIndex;
        }

        public void InvalidateFloatingShapeIndex()
        {
            _floatingShapeIndex = null;
        }
    }
}
