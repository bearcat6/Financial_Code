# VBA Rates Models

VBA prototypes for JPY rates analytics, including OIS curve construction, cap volatility bootstrapping, swaption volatility interpolation, and CMS-related valuation utilities.

This repository is for educational and prototyping purposes only. All sample data are fictional. No proprietary or confidential data are included.

## Design Policy

This repository is designed for practical Excel/VBA prototyping of JPY rates models.

The code is organized with the following principles:

- Keep financial model responsibilities separated.
- Keep small numerical components self-contained where portability is important.
- Avoid excessive abstraction so that the code remains understandable in Excel/VBA.
- Use clear naming rules suitable for VBA projects.
- Prefer explicit assumptions over hidden behavior.
- Design current components so that they can be extended later to OIS curves, cap volatility term structures, swaption volatility surfaces, and CMS-related valuation utilities.

## Naming Rules

The project follows these naming rules:

- Class modules start with `cls`.
- Standard modules start with `mdl_`.
- Function arguments start with `in_`.
- Object variables and object arguments use the prefix `c`.
- Internal class member variables use the prefix `m`.

Examples:

```text
clsHolidayCalendar
clsOISStepForwardCurve
clsSplineCurve

mdl_BusinessDay
mdl_DayCount
mdl_CapFormula

in_TargetDate
in_cCurve
mPointCount
