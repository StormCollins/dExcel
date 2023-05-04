using ExcelDna.Integration;
using QL = QuantLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using dExcel.Dates;
using dExcel.Utilities;

namespace dExcel.InterestRates
{
    using Utilities;

    public static class Instruments
    {
        public static object CreateInterestRateSwap(
            string swapType,
            DateTime startDate,
            DateTime endDate,
            string frequency,
            string calendars,
            string businessDayConvention,
            string generationRule)
        {
            QL.Swap.Type qlSwapType = (swapType.ToUpper() == "Payer")? QL.Swap.Type.Payer : QL.Swap.Type.Receiver;
            DateUtils.GenerateSchedule(startDate, endDate, frequency, calendars, businessDayConvention, generationRule);
            (QL.BusinessDayConvention? qlBusinessDayConvention, string? businessDayConventionErrorMessage) = 
                DateParserUtils.ParseBusinessDayConvention(businessDayConvention);
            (QL.Calendar? calendar, string? calendarErrorMessage) = DateParserUtils.ParseCalendars(calendars);
            QL.Schedule schedule = 
                new(startDate.ToQuantLibDate(), 
                    endDate.ToQuantLibDate(),
                    new QL.Period(frequency),
                    calendar,
                    (QL.BusinessDayConvention)qlBusinessDayConvention,
                    (QL.BusinessDayConvention)qlBusinessDayConvention,
                    QL.DateGeneration.Rule.Backward,
                    false);
            QL.VanillaSwap vanillaSwap =
                new(
                    type: QL.Swap.Type.Payer, 
                    nominal: 1000000, 
                    fixedSchedule: schedule, 
                    fixedRate: 0.1, 
                    fixedDayCount: new QL.Actual365Fixed(), 
                    floatSchedule: schedule, 
                    index: new QL.Jibar(new QL.Period("3m")), 
                    spread: 0,
                    floatingDayCount: new QL.Actual365Fixed());
           
            // QL.GaussianMultiPathGenerator multiPathGenerator = new()
            return vanillaSwap.NPV();
        }
    }
}
