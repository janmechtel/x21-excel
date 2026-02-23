using System;

namespace X21.Excel.Events
{
    public class SelectionChangedEventArgs : EventArgs
    {
        public int CellCount { get; }
        public string Address { get; }

        public SelectionChangedEventArgs(int cellCount, string address)
        {
            CellCount = cellCount;
            Address = address;
        }
    }
}
