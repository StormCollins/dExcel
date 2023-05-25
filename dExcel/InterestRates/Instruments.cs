using QL = QuantLib;
using dExcel.Dates;
using dExcel.Utilities;

namespace dExcel.InterestRates
{
    using ExcelDna.Integration;

    /// <summary>
    /// Common interest rate instruments.
    /// </summary>
    public static class Instruments
    {
        /// <summary>
        /// Used to create a fixed-for-floating interest rate swap object in Excel.
        /// </summary>
        /// <param name="swapType">'Payer' or 'Receiver'.</param>
        /// <param name="startDate">Start date.</param>
        /// <param name="endDate">End date.</param>
        /// <param name="frequency">Payment/receive frequency (assumed to be the same).</param>
        /// <param name="calendars">Calendars.</param>
        /// <param name="businessDayConvention">Business day convention.</param>
        /// <param name="generationRule">Date generation rule: 'Forward', 'Backward', or 'IMM'.
        /// 'Forward' => Dates are generated starting at the start date and moving towards the end date.
        /// 'Backward' => Dates are generated starting at the end date and moving towards the start date.
        /// 'IMM' => IMM convention i.e., business date on or closest to 20th of Mar, Jun, Sep, Dec.</param>
        /// <returns></returns>
        [ExcelFunction(
            Name = "d.IR_CreateInterestRateSwap",
            Description = "Creates a fixed-for-floating interest rate swap object in Excel",
            Category = "∂Excel: Interest Rates")]
        public static object CreateInterestRateSwap(
            string swapType,
            DateTime startDate,
            DateTime endDate,
            string frequency,
            string calendars,
            string businessDayConvention,
            string generationRule)
        {
            if (!ParserUtils.TryParseQuantLibSwapType(
                    swapType, 
                    out QL.Swap.Type? qlSwapType,
                    out string? swapTypeErrorMessage))
            {
                return swapTypeErrorMessage;
            }
            
            (QL.BusinessDayConvention? qlBusinessDayConvention, string? businessDayConventionErrorMessage) = 
                DateUtils.ParseBusinessDayConvention(businessDayConvention);
            (QL.Calendar? calendar, string? calendarErrorMessage) = DateUtils.ParseCalendars(calendars);
            
            QL.Schedule schedule = 
                new(effectiveDate: startDate.ToQuantLibDate(), 
                    terminationDate: endDate.ToQuantLibDate(),
                    tenor: new QL.Period(frequency),
                    calendar: calendar,
                    convention: (QL.BusinessDayConvention)qlBusinessDayConvention,
                    terminationDateConvention: (QL.BusinessDayConvention)qlBusinessDayConvention,
                    rule: QL.DateGeneration.Rule.Backward,
                    endOfMonth: false);
            
            QL.VanillaSwap vanillaSwap =
                new(
                    type: (QL.Swap.Type)qlSwapType, 
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
