# *MyJournal.Notebook*

:pushpin: _***MyJournal.Notebook*** makes journaling with OneNote as simple as possible, but not simpler!_  

![screenshot](docs/README-screenshot.png)

<div align="center">

<a href="">[![GitHub release (latest by date including pre-releases)](https://img.shields.io/github/v/release/atrenton/MyJournal.Notebook?color=blue&include_prereleases&logo=github)](https://github.com/atrenton/MyJournal.Notebook/releases)&emsp;&emsp;</a><a href="">[![GitHub Discussions](https://img.shields.io/github/discussions/atrenton/MyJournal.Notebook?color=green&logo=github)](https://github.com/atrenton/MyJournal.Notebook/discussions)&emsp;&emsp;</a><a href="">[![GitHub issues](https://img.shields.io/github/issues/atrenton/MyJournal.Notebook?logo=github)](https://github.com/atrenton/MyJournal.Notebook/issues)&emsp;&emsp;</a><a href="">[![tweets](https://img.shields.io/badge/twitter-545454.svg?logo=twitter)](https://twitter.com/ArtTrenton)</a>

</div>

## About

Record your daily interactions, ideas and inspirations with this add-in for Microsoft® OneNote® Windows desktop software.

<div align="center">
<table hspace="25">
  <tr>
    <th scope="row">
      <img src="docs/journal.png" alt="journal" />
    </th>
    <td>With one click of a button, this add-in
    <br />creates a notebook organized by
    <br />year, month, and day.</td>
  </tr>
  <tr />
  <tr>
    <th scope="row">User
    <br />Configurable
    <br />Settings</th>
    <td>
      <ul>
        <li>OneDrive storage (personal accounts only)</li>
        <li>Page color</li>
        <li>Page title date format</li>
        <li>Page rule lines</li>
        <li>Page template</li>
        <li>Paper size</li>
      </ul>
    </td>
  </tr>
  <tr />
  <tr>
    <th scope="row">Language</th>
    <td>C #</td>
  </tr>
  <tr />
  <tr>
    <th scope="row">License</th>
    <td>
      <a href="LICENSE.txt">Microsoft Public License (MS-PL)</a>
    </td>
  </tr>
  <tr />
  <tr>
    <th scope="row">Disclaimer</th>
    <td><b><i>MyJournal.Notebook</i></b> software is not developed by or affiliated with the Microsoft Corporation.</td>
  </tr>
  <tr />
  <tr>
    <th scope="row">Trademarks</th>
    <td>Microsoft and OneNote are registered trademarks of Microsoft Corporation.</td>
  </tr>
</table>
</div>

For additional information, check out the [**Wiki**](https://github.com/atrenton/MyJournal.Notebook/wiki).

## Prerequisites

- Microsoft [OneNote for Windows] desktop application software (2016 or later)<br />
- Microsoft [.NET 6 Supported Windows OS]<br />
- Microsoft [.NET 6 Windows Desktop Runtime], version 6.0.21 or later (x86 for 32-bit OneNote; x64 for 64-bit OneNote)<br />
- Microsoft Visual Studio 2022 version 17.3 or later (developers only)<br />

## Installation

- To use this add-in, you must have a Windows desktop version of OneNote installed as part of [Microsoft 365](https://www.microsoft.com/en-us/microsoft-365) or Office. See the **OneNote from Microsoft 365 now in the Microsoft Store** section of the following document for more information:
    - [Making it easier to get to the OneNote app on Windows](https://techcommunity.microsoft.com/t5/microsoft-365-blog/making-it-easier-to-get-to-the-onenote-app-on-windows/ba-p/3642219)

&NewLine;

- Supported versions of OneNote on devices running Windows:
    - OneNote for Microsoft 365
    - OneNote for Office 2016 or later

&NewLine;

- For supported versions of OneNote, download and install the [latest release] of the **MyJournal.Notebook.Setup** program.

&NewLine;

- Unsupported versions of OneNote:
    - [OneNote for Windows 10]
    - [OneNote 2013](https://learn.microsoft.com/en-us/lifecycle/products/microsoft-onenote-2013 "Microsoft Lifecycle Extended End Date: 04/11/2023") &mdash; if you are still using it, install [release 16.3.0].
    - [OneNote 2010](https://learn.microsoft.com/en-us/lifecycle/products/microsoft-onenote-2010 "Microsoft Lifecycle Extended End Date: 10/13/2020") &mdash; if you are still using it, install [release 16.2.0].

## Usage

- [Select journal page template](docs/HowTo-Select-Journal-Page-Template.md)
- [Select journal paper size](docs/HowTo-Select-Journal-Paper-Size.md)
- [Create journal page](docs/HowTo-Create-Journal-Page.md)
- [Select journal page color](docs/HowTo-Select-Page-Color.md)
- [Select journal page title date format](docs/HowTo-Select-Page-Title.md)
- [Select journal page rule lines format](docs/HowTo-Select-Rule-Lines.md)
- [Select OneDrive storage account](docs/HowTo-Select-OneDrive-Storage-Account.md)

[latest release]:https://github.com/atrenton/MyJournal.Notebook/releases "latest by date including pre-releases"
[release 16.2.0]:https://github.com/atrenton/MyJournal.Notebook/releases/tag/v16.2.0
[release 16.3.0]:https://github.com/atrenton/MyJournal.Notebook/releases/tag/v16.3.0

[.NET 6 Supported Windows OS]:https://github.com/dotnet/core/blob/main/release-notes/6.0/supported-os.md#windows "core/supported-os.md at main · dotnet/core · GitHub"
[.NET 6 Windows Desktop Runtime]:https://dotnet.microsoft.com/en-us/download/dotnet/6.0 "Download .NET Desktop Runtime"

[OneNote for Windows]:https://support.microsoft.com/en-us/office/what-s-the-difference-between-the-onenote-versions-a624e692-b78b-4c09-b07f-46181958118f#windows "What's the difference between the OneNote versions?"
[OneNote for Windows 10]:https://support.microsoft.com/en-us/office/what-s-the-difference-between-the-onenote-versions-a624e692-b78b-4c09-b07f-46181958118f#windows "What's the difference between the OneNote versions?"
