using NPOI.SS.Formula;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.UserModel.Helpers
{
    /// <summary>
    /// Helper for Shifting columns left or right
    /// This abstract class exists to consolidate duplicated code between 
    /// XSSFColumnShifter and HSSFColumnShifter
    /// </summary>
    public abstract class ColumnShifter : BaseRowColShifter
    {
        protected ISheet sheet;

        public ColumnShifter(ISheet sh)
        {
            sheet = sh;
        }

        /// <summary>
        /// Shifts, grows, or shrinks the merged regions due to a column Shift.
        /// Merged regions that are completely overlaid by Shifting will 
        /// be deleted.
        /// </summary>
        /// <param name="startColumn">the column to start Shifting</param>
        /// <param name="endColumn">the column to end Shifting</param>
        /// <param name="n">the number of columns to shift</param>
        /// <returns>an array of affected merged regions, doesn't contain 
        /// deleted ones</returns>
        public override List<CellRangeAddress> ShiftMergedRegions(
            int startColumn,
            int endColumn,
            int n)
        {
            List<CellRangeAddress> shiftedRegions = [];
            HashSet<int> removedIndices = [];
            //move merged regions completely if they fall within the new region
            //boundaries when they are Shifted
            int size = sheet.NumMergedRegions;

            for(int i = 0; i < size; i++)
            {
                CellRangeAddress merged = sheet.GetMergedRegion(i);

                // remove merged region that overlaps Shifting
                if(RemovalNeeded(merged, startColumn, endColumn, n))
                {
                    _ = removedIndices.Add(i);
                    continue;
                }

                bool inStart = merged.FirstColumn >= startColumn
                    || merged.LastColumn >= startColumn;
                bool inEnd = merged.FirstColumn <= endColumn ||
                    merged.LastColumn <= endColumn;

                //don't check if it's not within the Shifted area
                if(!inStart || !inEnd)
                {
                    continue;
                }

                //only shift if the region outside the Shifted columns is not
                //merged too
                if(!merged.ContainsColumn(startColumn - 1)
                    && !merged.ContainsColumn(endColumn + 1))
                {
                    merged.FirstColumn += n;
                    merged.LastColumn += n;
                    //have to Remove/add it back
                    shiftedRegions.Add(merged);
                    _ = removedIndices.Add(i);
                }
            }

            if(removedIndices.Count != 0)
            {
                sheet.RemoveMergedRegions(removedIndices.ToList());
            }

            //read so it doesn't Get Shifted again
            foreach(CellRangeAddress region in shiftedRegions)
            {
                _ = sheet.AddMergedRegion(region);
            }

            return shiftedRegions;
        }

        // Keep in sync with {@link ColumnShifter#removalNeeded}
        private bool RemovalNeeded(
            CellRangeAddress merged,
            int startColumn,
            int endColumn,
            int n)
        {
            int firstColumn = startColumn + n;
            int lastColumn = endColumn + n;
            CellRangeAddress overwrite =
                new CellRangeAddress(0, sheet.LastRowNum, firstColumn, lastColumn);

            // if the merged-region and the overwritten area intersect,
            // we need to remove it
            return merged.Intersects(overwrite);
        }

        
        public void ShiftColumns(int firstShiftColumnIndex, int lastShiftColumnIndex, int step)
        {
            if(step > 0)
            {
                foreach(IRow row in sheet)
                    if(row != null)
                        row.ShiftCellsRight(firstShiftColumnIndex, lastShiftColumnIndex, step);
            }
            else if(step < 0)
            {
                foreach(IRow row in sheet)
                    if(row != null)
                        row.ShiftCellsLeft(firstShiftColumnIndex, lastShiftColumnIndex, -step);
            }
            //else step == 0 => nothing to shift
        }
    }
}
