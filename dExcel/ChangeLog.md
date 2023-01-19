# Changelog
All notable changes to ∂Excel will be documented in this file.

## [0.1.3] - 2023-01-19 
### Features
- Added function ``d.Stats_CorrelatedNormalRandomNumbers``.
- Improved holiday parsing in DateUtils so that multiple sets of holidays can be parsed simultaneously.
- Added function ``d.IR_BlackForwardOptionPricer``.
- ``d.Dates_FolDay``, ``d.Dates_ModFolDay``, ``d.Dates_PrevDay`` et. al. can parse multiple lists of holidays i.e., a 2D list of holidays.
- ``d.Equity_BlackScholesSpotOptionPricer`` now provides verbose output.

### Refactoring
- Changed ``d.Math_Cholesky`` to return an upper triangular rather than a lower triangular matrix.

### Bugs
- Fixed issue where ``d.Dates_Fol``, ``d.Dates_ModFol``, ``d.Dates_PrevDay`` calendars was not updating correctly.

## [0.1.2] - 2022-12-12
### Features
- Enabled ExcelDnaIntelliSense.
- Added day count convention functions (e.g., ``d.Dates_Act365`` etc.).

## [0.1.1] - 2022-11-25
### Features
- Initial release.

## [0.0] - 2022-08-24
### Features:
- All AQS toolbox functions replaced.
- Non-database (i.e. in sheet) single curve bootstrapping functional.
