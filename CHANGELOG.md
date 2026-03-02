# Changelog

All notable changes to Canvas Tools for Google Sheets will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.1] - 2026-03-02

### Fixed
- Removed duplicate `getCanvasDomain()` declaration that caused a runtime error
- Fixed `ReferenceError` caused by undefined `CANVAS_DOMAIN` variable in grade upload functions; replaced with `getCanvasDomain()` calls
- Removed duplicate `fetchAllCanvasUsers_` declaration across script files
- Removed duplicate `parseLinkHeader_` declarations; consolidated into a single utility function
- Removed duplicate `columnLetterToNumber_`/`columnToNumber_` declarations; unified all call sites to use `columnLetterToNumber_`

## [1.0.0] - 2025-05-17

### Added
- Initial release of Canvas Tools for Google Sheets
- Fetch course users from Canvas into Google Sheets
- Fetch individual assignment grades or complete gradebook
- Upload individual grades, grade ranges, or complete gradebook to Canvas
- Configuration through spreadsheet (Course ID, Canvas Domain, API Key)
- Secure API key handling with Script Properties
- Progress tracking with toast notifications
- Extensive error handling and user feedback
- Comprehensive documentation and setup guide

[1.0.1]: https://github.com/jlouisbru/grade-tracking-Canvas-API/releases/tag/v1.0.1
[1.0.0]: https://github.com/jlouisbru/canvas-tools-for-sheets/releases/tag/v1.0.0
