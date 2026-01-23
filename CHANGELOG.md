# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0] - 2026-01-23

### Added

- Created dedicated `CHANGELOG.md` file
- Added `CONSTANTS.py` for centralized configuration
- Added `UTILS.py` for helper functions

### Security

- Migrated `SECRET_KEY` and `DEBUG` settings to environment variables (`.env`)

## [1.1.0] - 2025-06-23

### Added

- Added `django_extensions` to installed apps
- Added Graphviz configuration for model visualization

### Changed

- Refactored project structure to remove unused session keys
- Updated comments to correctly reference `site_id` instead of `plant_code`

### Fixed

- Fixed import issues in viewsw

## [1.0.0] - 2025-03-07

### Added

- **Major UI Update**: Redesigned interface with Tailwind CSS and Flowbite
- Added comprehensive error handling pages (404, 500)
- Implemented `UploadFileForm` with extensive validation logic

### Fixed

- Resolved Excel data processing issues for large files
- Fixed "Plant Unit Mismatch" validation logic

## [0.9.0] - 2025-02-25

### Added

- Initial project structure setup with Django 5.1
- Basic file upload functionality
- Integration with Pandas for Excel processing
- Background task setup using `django-background-tasks`
