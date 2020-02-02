# Known Issues

This document describes known issues with the My Journal Add-In for Microsoft OneNote (32-bit).

## Paper Size

### Background:
This feature was added in release 16.1.0 to support a user's paper size preference
when printing a journal page. It defaults to the OneNote Auto (automatic) paper size.
This results in a single continuous page with an unlimited height. By OneNote design,
page breaks are not supported at this time.

### Issue: [Configurable page size #1](https://github.com/atrenton/MyJournal.Notebook/issues/1)
Setting the paper size to any value other than Automatic will result in a fixed
height page size. Content exceeding this height overflows the background page
color.

### Workaround:
- Leave the My Journal Notebook paper size option set to Automatic.
- Use **File \> Print \> Print Preview** to see where your page breaks are.
- Per [Microsoft Office Support](https://support.office.com/en-us/article/change-the-line-spacing-in-onenote-7de3c45a-5b5d-477b-b405-74877d8e18d1):

	> OneNote pages aren't like pages in Word. In OneNote, pages can go on and on.
	> And because OneNote is designed to capture your notes, not print out traditional
	> pages, you won't find a page break option in OneNote. Go to **File \> Print \>**
	> **Print Preview** to see how your pages will look when printed.
