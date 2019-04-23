using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PayStubCreator
{
    public class Employee
    {
        private const decimal BASE_HOUR_VALUE = 40;
        private const decimal PUBLIC_TRANSIT_TICKET = 150;
        private const decimal COMP_PER_KM = 3;
        private const decimal PERSONAL_WRITEOFF = 750;
        private const decimal WRITEOFF_PER_SUPPORTED_MEMBER = 500;
        private const decimal OVERTIME_MULTIPLICATOR = 1.3M;
        private const decimal SICK_LEAVE_MULTIPLICATOR = 0.7M;
        private const decimal TAX_BRACKET_1 = 1500;
        private const decimal TAX_BRACKET_1_RATE = 0.2M;
        private const decimal TAX_BRACKET_2 = 3000;
        private const decimal TAX_BRACKET_2_RATE = 0.4M;
        public const decimal RETIREMENT_CONTRIBUTION_RATE = 0.15M;

        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int EmployeeID { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public int PostCode { get; set; }
        public string Title { get; set; }
        public decimal Coefficient { get; set; }
        public decimal DistanceFromWork { get; set; }
        public int? NumberOfDependents { get; set; }
        public int HoursWorked { get; set; }
        public int HoursOvertime { get; set; }
        public int HoursVacation { get; set; }
        public int SickLeave { get; set; }
        public int DaysWorked { get; set; }
        public string FullName
        {
            get
            {
                return string.Format("{0} {1}", FirstName, LastName);
            }
        }
        public Employee()
        {
            DaysWorked = 0;
        }

        public string AddressHeader()
        {
            return String.Format("{0}\n{1}\n{2} {3}\nWork title: {4}", FullName, Address, PostCode, City, Title);
        }

        public decimal TravelExpense()
        {
            decimal travelExpense;

            if (DistanceFromWork <= 10)
                travelExpense = PUBLIC_TRANSIT_TICKET;
            else
                travelExpense = COMP_PER_KM * DistanceFromWork;
            if (DaysWorked == 14)
            {
                return travelExpense;
            }
            else
            {
                return travelExpense * DaysWorked / 14;
            }
        }

        public decimal TaxBreak()
        {
            if (NumberOfDependents == 0 || NumberOfDependents == null)
            {
                return PERSONAL_WRITEOFF;
            }
            else
            {
                return (int)NumberOfDependents * WRITEOFF_PER_SUPPORTED_MEMBER;
            }
        }

        public decimal TotalPay()
        {
            return BASE_HOUR_VALUE * (HoursWorked + HoursVacation + OVERTIME_MULTIPLICATOR * HoursOvertime + SickLeave * SICK_LEAVE_MULTIPLICATOR);
        }
        public decimal TaxesOwed()
        {
            decimal taxBreak = TaxBreak();
            decimal totalPay = TotalPay();

            if (totalPay <= taxBreak || totalPay <= TAX_BRACKET_1)
            {
                return 0;
            }
            else
            {
                if (totalPay > TAX_BRACKET_1 && totalPay < TAX_BRACKET_2)
                {
                    return (totalPay - taxBreak) * TAX_BRACKET_1_RATE;
                }
                else
                {
                    return (TAX_BRACKET_1 - taxBreak) * TAX_BRACKET_1_RATE + (totalPay - TAX_BRACKET_1) * TAX_BRACKET_2_RATE;
                }
            }
        }
        public decimal NetPay()
        {
            return TotalPay() - TaxesOwed() - RetirementDeduction() + TravelExpense();
        }
        public decimal RetirementDeduction()
        {
            return TotalPay() * RETIREMENT_CONTRIBUTION_RATE;
        }
    }
}
