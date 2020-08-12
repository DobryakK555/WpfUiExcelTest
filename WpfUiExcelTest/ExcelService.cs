using System.IO;
using NPOI.SS.UserModel;
using System.Data;
using WpfUiExcelTest.Models;
using System.Collections.Generic;
using System;
using System.Reflection;

namespace WpfUiExcelTest
{
    public static class ExcelService
    {
        public static string FileName { get; set; }

        public static DateTime FirstDate { get; set; }

        public static DateTime SecondDate { get; set; }

        internal static bool CheckDates(DateTime? firstDate, DateTime? secondDate)
        {
            if ((firstDate != null) && (secondDate != null) && (firstDate.Value <= secondDate.Value))
            {
                FirstDate = (DateTime)firstDate;
                SecondDate = (DateTime)secondDate;
                return true;
            }
            else
            {
                return false;
            }
        }

        internal static List<int> GetFilteredSet()
        {
            var filteredRows = new List<int>();

            if (String.IsNullOrEmpty(FileName))
                return filteredRows;

            var column1 = 1;
            var column2 = 2;

            using (FileStream file = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                var workbook = WorkbookFactory.Create(file);
                var sheet = workbook.GetSheetAt(0);


                for (int i = 1; i < sheet.PhysicalNumberOfRows; i++)
                {
                    var sheetRow = sheet.GetRow(i);

                    var arrivalDate = sheetRow.GetCell(column1).DateCellValue;
                    var departmentDate = sheetRow.GetCell(column2, MissingCellPolicy.CREATE_NULL_AS_BLANK).DateCellValue;

                    if ((SecondDate < arrivalDate) || (FirstDate > departmentDate))
                        continue;

                    filteredRows.Add(i);
                }
            }
            return filteredRows;
        }

        internal static DataTable DoWork(List<int> shipmentsIndex)
        {
            using (FileStream file = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                var workbook = WorkbookFactory.Create(file);
                var shipmentSheet = workbook.GetSheetAt(0);
                var shipments = new List<IRow>();

                var allRates = new List<CalculatedRate>();

                var tariffs = ReadTariffSheet(workbook);

                foreach (var index in shipmentsIndex)
                {
                    shipments.Add(shipmentSheet.GetRow(index));
                }

                foreach (var row in shipments)
                {
                    var rates = DoRates(row, tariffs);
                    allRates.AddRange(rates);
                }
                return CreateDataTable(allRates);
            }
        }

        private static DataTable CreateDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

        private static List<CalculatedRate> DoRates(IRow row, List<Tariff> tariffs)
        {
            var arrivalDate = row.GetCell(1).DateCellValue;
            var departmentDate = row.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK).DateCellValue;
            var result = new List<CalculatedRate>();

            var currentDate = FirstDate >= arrivalDate ? FirstDate : arrivalDate;
            var days = (SecondDate - currentDate).Days;

            var index = 0;

            while (days > 0)
            {
                var rate = new CalculatedRate();
                
                var tariff = tariffs[index];

                rate.ArrivalDate = arrivalDate;
                rate.DepartmentDate = departmentDate;
                rate.Name = row.GetCell(0).StringCellValue;

                rate.CalculationStart = currentDate;
                if (tariff.PeriodEnd == 0)
                    currentDate = SecondDate;
                else
                    currentDate += TimeSpan.FromDays(tariff.PeriodEnd - tariff.PeriodStart);

                if (currentDate >= SecondDate)
                {
                    rate.CalculationEnd = SecondDate;
                }
                else
                {
                    rate.CalculationEnd = currentDate;
                }

                rate.Rate = tariff.Rate;
                rate.StorageDays = (rate.CalculationEnd - rate.CalculationStart).Days;
                rate.AdditionalInfo = $"Период №{index + 1}";

                days = days - rate.StorageDays;
                index += 1;
                result.Add(rate);
            }

            return result;
        }

        private static List<Tariff> ReadTariffSheet(IWorkbook workbook)
        {
            var sheet = workbook.GetSheetAt(1);

            var tariffs = new List<Tariff>();

            for (int i = 1; i < sheet.PhysicalNumberOfRows; i++)
            {
                var sheetRow = sheet.GetRow(i);
                var tariff = new Tariff(sheetRow);

                tariffs.Add(tariff);
            }
            return tariffs;
        }

    }
}