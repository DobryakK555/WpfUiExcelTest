using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WpfUiExcelTest.Models
{
    public class CalculatedRate : INotifyPropertyChanged
    {
        public string Name { get; set; }

        public DateTime ArrivalDate { get; set; }

        public DateTime DepartmentDate { get; set; }

        public DateTime CalculationStart { get; set; }

        public DateTime CalculationEnd { get; set; }

        public int StorageDays { get; set; }

        public double Rate { get; set; }

        public string AdditionalInfo { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
