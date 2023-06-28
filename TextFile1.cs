/*
 Copyright (C) 2005 Dominic Thuillier

 This file is part of QuantLib, a free-software/open-source library
 for financial quantitative analysts and developers - http://quantlib.org/

 QuantLib is free software: you can redistribute it and/or modify it
 under the terms of the QuantLib license.  You should have received a
 copy of the license along with this program; if not, please email
 <quantlib-dev@lists.sf.net>. The license is also available online at
 <http://quantlib.org/license.shtml>.

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

using System;
using QuantLib;

namespace BermudanSwaption
{
	class Run
	{


		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static double Main()
		{
			DateTime startTime = DateTime.Now;

			Date todaysDate = new Date(15, Month.February, 2002);
			Calendar calendar = new TARGET();
			Date settlementDate = new Date(19, Month.February, 2002);
			Settings.instance().setEvaluationDate(todaysDate);

			// flat yield term structure impling 1x5 swap at 5%
			Quote flatRate = new SimpleQuote(0.04875825);
			FlatForward myTermStructure = new FlatForward(
				settlementDate,
				new QuoteHandle(flatRate),
				new Actual365Fixed());
			RelinkableYieldTermStructureHandle rhTermStructure =
				new RelinkableYieldTermStructureHandle();
			rhTermStructure.linkTo(myTermStructure);

			// Define the ATM/OTM/ITM swaps
			Period fixedLegTenor = new Period(1, TimeUnit.Years);
			BusinessDayConvention fixedLegConvention =
				BusinessDayConvention.Unadjusted;
			BusinessDayConvention floatingLegConvention =
				BusinessDayConvention.ModifiedFollowing;
			DayCounter fixedLegDayCounter =
				new Thirty360(Thirty360.Convention.European);
			Period floatingLegTenor = new Period(6, TimeUnit.Months);
			double dummyFixedRate = 0.03;
			IborIndex indexSixMonths = new Euribor6M(rhTermStructure);

			Date startDate = calendar.advance(settlementDate, 1, TimeUnit.Years,
				floatingLegConvention);
			Date maturity = calendar.advance(startDate, 5, TimeUnit.Years,
				floatingLegConvention);
			Schedule fixedSchedule = new Schedule(startDate, maturity,
				fixedLegTenor, calendar, fixedLegConvention, fixedLegConvention,
				DateGeneration.Rule.Forward, false);
			Schedule floatSchedule = new Schedule(startDate, maturity,
				floatingLegTenor, calendar, floatingLegConvention,
				floatingLegConvention, DateGeneration.Rule.Forward, false);
			VanillaSwap swap = new VanillaSwap(
					   Swap.Type.Payer, 1000.0,
					   fixedSchedule, dummyFixedRate, fixedLegDayCounter,
					   floatSchedule, indexSixMonths, 0.0,
					   indexSixMonths.dayCounter());
			DiscountingSwapEngine swapEngine =
				new DiscountingSwapEngine(rhTermStructure);
			swap.setPricingEngine(swapEngine);
			double fixedATMRate = swap.fairRate();

			return fixedATMRate;
		}
	}
}
