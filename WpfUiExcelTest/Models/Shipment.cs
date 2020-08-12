using NPOI.SS.UserModel;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WpfUiExcelTest.Models
{
    public class Shipment : INotifyPropertyChanged
    {
        private string _name;
        private DateTime _arrivalDate;
        private DateTime? _departmentDate;

        public string Name
        {
            get => _name;
            set
            {
                if (value == _name)
                    return;
                _name = value;
                OnPropertyChanged();
            }
        }

        public DateTime ArrivalDate
        {
            get => _arrivalDate;
            set
            {
                if (value == _arrivalDate)
                    return;
                _arrivalDate = value;
                OnPropertyChanged();
            }
        }

        public DateTime? DepartmentDate
        {
            get => _departmentDate;
            set
            {
                if (value == _departmentDate)
                    return;
                _departmentDate = value;
                OnPropertyChanged();
            }
        }

        public Shipment(IRow sheetRow)
        {
            Name = sheetRow.GetCell(0).StringCellValue;
            ArrivalDate = sheetRow.GetCell(1).DateCellValue;
            DepartmentDate = !String.IsNullOrEmpty(sheetRow.GetCell(2).DateCellValue.ToString()) ? sheetRow.GetCell(2).DateCellValue : (DateTime?)null;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
