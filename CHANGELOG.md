# Change Log

All notable changes to this project will be documented in this file. This project adheres to [Semantic Versioning](https://semver.org/).

## [v16.4.0] - 2023-08-30

### Added

- Support for SVG image processing
- New Retro page template with title background image; replicates **OneNote 2010** look

### Changed

- Migrated codebase from .NET Framework to .NET 6 ([Long Term Support](https://learn.microsoft.com/en-us/lifecycle/products/microsoft-net-and-net-core ".NET 6.0 (LTS) End Date: 11/12/2024"))
- No longer need to run Visual Studio as Administrator to compile the code
- Uses [RegSvr32] now to register/unregister the generated COM Host file instead of [RegAsm] for the application DLL file
- Many code quality improvements

### Removed

- **OneNote 2013** is no longer supported
- Rule Set files were deprecated in favor of the EditorConfig file

## [v16.3.0] - 2021-12-05

### Added

- Feature request: configure the notebook used &bull; Issue [**#3**] &bull; released from beta

### Changed

- Upgraded OneNote XML schema from 2010 to 2013 version

### Removed

- **OneNote 2010** is no longer supported

## [v16.2.1-beta] - 2021-04-07

### Added

- Support for storage of 'My Journal' notebook on OneDrive - Issue [**#3**]

## [v16.2.0] - 2020-09-01

### Added

- Support for 64-bit versions of OneNote

## [v16.1.0] - 2020-02-02

### Added

- Paper size user preference; defaults to Auto (automatic) - Issue [**#1**]

## [v16.0.0] - 2019-09-03

### Added

- Initial release

[**#1**]:https://github.com/atrenton/MyJournal.Notebook/issues/1
[**#3**]:https://github.com/atrenton/MyJournal.Notebook/issues/3
[RegAsm]:https://learn.microsoft.com/en-us/previous-versions/dotnet/netframework-4.0/tzat5yw6(v=vs.100) 'Regasm.exe (Assembly Registration Tool) | Microsoft Learn'
[RegSvr32]:https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/regsvr32 'regsvr32 | Microsoft Learn'
[v16.0.0]:https://github.com/atrenton/MyJournal.Notebook/tree/v16.0.0
[v16.1.0]:https://github.com/atrenton/MyJournal.Notebook/tree/v16.1.0
[v16.2.0]:https://github.com/atrenton/MyJournal.Notebook/tree/v16.2.0
[v16.2.1-beta]:https://github.com/atrenton/MyJournal.Notebook/tree/v16.2.1-beta
[v16.3.0]:https://github.com/atrenton/MyJournal.Notebook/tree/v16.3.0
[v16.4.0]:https://github.com/atrenton/MyJournal.Notebook/tree/v16.4.0
