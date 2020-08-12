using NPOI.SS.UserModel;
using System;
using System.Linq;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Documents;
using System.Collections.Generic;

namespace WpfUiExcelTest.Models
{
    public class Tariff : INotifyPropertyChanged
    {
        private double _id;
        private int _periodStart;
        private int _periodEnd;
        private double _rate;

        public Tariff(IRow sheetRow)
        {
            Id = (int)sheetRow.GetCell(0).NumericCellValue;
            PeriodStart = (int)sheetRow.GetCell(1).NumericCellValue;
            PeriodEnd = (int)sheetRow.GetCell(2).NumericCellValue;
            Rate = (int)sheetRow.GetCell(3).NumericCellValue;
        }

        public double Id
        {
            get { return _id; }
            set
            {
                if (value == _id)
                    return;
                _id = value;
                OnPropertyChanged();
            }
        }

        public int PeriodStart
        {
            get { return _periodStart; }
            set
            {
                _periodStart = value;
                OnPropertyChanged();
            }
        }

        public int PeriodEnd
        {
            get { return _periodEnd; }
            set
            {
                _periodEnd = value;
                OnPropertyChanged();
            }
        }

        public double Rate
        {
            get { return _rate; }
            set
            {
                _rate = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
