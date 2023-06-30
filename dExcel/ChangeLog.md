# Changelog
All notable changes to ∂Excel will be documented in this file.

## [0.2.3] - 2023-06-29
### Features
- ``USDZAR`` FX basis adjusted curve and ``SOFR`` curve can now be pulled directly form Omicron.
- Added function ``d.Stats_GBM_Create``.
- Added function ``d.Stats_GBM_GetPaths``.
- Interest rate swap pricing features added.

## [0.2.2] - 2023-6-15
- Support added for manual ``SOFR`` and tenor basis bootstrapping.

## [0.2.1] - 2023-06-09
### Features
- Excel table column and row headers support spaces in names (improves readability).
- Added ``Insert Enums`` to ribbon for inserting drop down menus of common types e.g., rate indices.
- Interpolation for bootstrapping is now clearly state as being on discount factors, zero rates, or forward rates.
- Renamed ``d.Curve_Create`` to ``d.Curve_CreateFromDiscountFactors``.
- Added function ``d.Curve_CreateFromZeroRates``.
- Bootstrapping supports multiple instrument names e.g., ``Interest Rate Swaps`` or ``IRSs``, ``Depos`` or ``Deposits`` etc.
- Bootstrapping FX basis adjusted curves added.

## [0.2.0] - 2023-05-03
### Features
- Added multi-curve bootstrapping support in ``d.Curve_Boostrap``.
- Converted from QLNet to QuantLib-SWIG.

## [0.1.4] - 2023-02-21
### Features
- Added ``d.Math_InterpolateChosenColumn``.
- Enabled ``Run All Tests`` in ``dexcel-testing.xlsm``.
- Added tests for ``Hyperlink Utils`` in ``dexcel-testing.xlsm``.

### Bugs
- Fixed ``Fix EMS Links`` functionality.
- Fixed single curve bootstrapping.
- Made ``Rate Tenor`` last parameter in ``d.IR_ConvertInterestRate`` and marked it as optional.

## [0.1.3] - 2023-01-20 
### Features
- Enabled ExcelDnaIntelliSense.
- Added day count convention functions (e.g., ``d.Dates_Act365`` etc.).
- Added function ``d.Stats_CorrelatedNormalRandomNumbers``.
- Improved holiday parsing in DateUtils so that multiple sets of holidays can be parsed simultaneously.
- Added function ``d.IR_BlackForwardOptionPricer``.
- ``d.Dates_FolDay``, ``d.Dates_ModFolDay``, ``d.Dates_PrevDay`` et. al. can parse multiple lists of holidays i.e., a 2D list of holidays.
- ``d.Equity_BlackScholesSpotOptionPricer`` now provides verbose output.

### Refactoring
- Changed ``d.Math_Cholesky`` to return an upper triangular rather than a lower triangular matrix.

### Bugs
- Fixed issue where ``d.Dates_FolDay``, ``d.Dates_ModFolDay``, ``d.Dates_PrevDay`` calendars was not updating correctly.

## [0.1.1] - 2022-11-25
### Features
- Initial release.

## [0.0] - 2022-08-24
### Features:
- All AQS toolbox functions replaced.
- Non-database (i.e. in sheet) single curve bootstrapping functional.
