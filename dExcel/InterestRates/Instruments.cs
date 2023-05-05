using QL = QuantLib;
using dExcel.Dates;

namespace dExcel.InterestRates
{
    /// <summary>
    /// Common interest rate instruments.
    /// </summary>
    public static class Instruments
    {
        /// <summary>
        /// Used to create a fixed-for-floating interest rate swap object in Excel.
        /// </summary>
        /// <param name="swapType">'Payer' or 'Receiver'.</param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="frequency"></param>
        /// <param name="calendars"></param>
        /// <param name="businessDayConvention"></param>
        /// <param name="generationRule"></param>
        /// <returns></returns>
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
                DateUtils.ParseBusinessDayConvention(businessDayConvention);
            (QL.Calendar? calendar, string? calendarErrorMessage) = DateUtils.ParseCalendars(calendars);
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
