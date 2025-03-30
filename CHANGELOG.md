# Changelog

All notable changes to the ExcelSync package will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.1] - 2024-05-11

### Added
- Support for custom header row positions
- Added parameter `header_row` to the constructor of `ExcelSync` class
- Added optional parameter `header_row` to methods that interact with Excel structure
- Improved tests to cover custom header row functionality
- Updated documentation with usage examples for custom header row feature

### Changed
- Enhanced structure validation to include header row information
- Added header row information to the structure's file_properties section
- Updated README with examples of using custom header rows

## [0.1.0] - 2024-05-10

### Added
- Initial release of ExcelSync
- Core functionality for Excel structure validation and comparison
- Support for extracting Excel structure
- Export to JSON and YAML formats
- Schema validation for Excel structures 