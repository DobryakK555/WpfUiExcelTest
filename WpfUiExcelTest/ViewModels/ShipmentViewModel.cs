using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfUiExcelTest.Models;

namespace WpfUiExcelTest.ViewModels
{
    class ShipmentViewModel : BaseViewModel
    {
        public ObservableCollection<Shipment> Shipments { get; set; }

        public ShipmentViewModel()
        {

        }
    }
}
