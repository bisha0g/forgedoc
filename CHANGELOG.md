# Changelog

## [1.1.2] - 2025-03-07

### Changed
- Updated package version from 1.1.1 to 1.1.2
- Maintained DocumentFormat.OpenXml package at version 3.2.0 for compatibility

### Fixed
- Implemented ReplaceLoop method for list iteration in Word templates
- Enhanced ReplaceTable method with improved table handling:
  - Added support for table markers in paragraphs before tables
  - Added support for table markers inside table cells
  - Improved table formatting and content replacement
  - Added automatic header row generation
- Added helper method ProcessTableData for better table processing
- Added UpdateCellContent method for cleaner cell content management

## [1.1.1] - 2024-03-21

### Added
- Improved table search functionality in Word documents
- Table formatting preservation when replacing content

## [1.1.0] - 2024-03-21

### Added
- Word template processor for generating documents from .docx templates
- Support for simple variable replacements using {{variable}} syntax
- Support for list iterations using {%for item in items%} syntax

## [1.0.0] - 2024-03-20

### Added
- Initial release
