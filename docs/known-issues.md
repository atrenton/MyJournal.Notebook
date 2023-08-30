# Known Issues

<sup>Last Updated August 30, 2023</sup>

The following is the list of Known Issues for the **My Journal Add-In for Microsoft OneNote**:

- [Paper Size](./known-issues.md#paper-size)
- [Rule Lines](./known-issues.md#rule-lines)

## Paper Size

### Background

This feature was added in release 16.1.0 to support a user's paper size preference
when printing a journal page. It defaults to the OneNote Auto (automatic) paper size.
This results in a single continuous page with an unlimited height. By OneNote design,
page breaks are not supported at this time.

### Issue: [Configurable page size #1](https://github.com/atrenton/MyJournal.Notebook/issues/1)

Setting the paper size to any value other than Automatic will result in a fixed
height page size. Content exceeding this height overflows the background page
color.

### Workaround

- Leave the My Journal Notebook paper size option set to Automatic.
- Use **File \> Print \> Print Preview** to see where your page breaks are.
- Per [Microsoft Office Support](https://support.office.com/en-us/article/change-the-line-spacing-in-onenote-7de3c45a-5b5d-477b-b405-74877d8e18d1):

    > OneNote pages aren't like pages in Word. In OneNote, pages can go on and on.
    > And because OneNote is designed to capture your notes, not print out traditional
    > pages, you won't find a page break option in OneNote. Go to **File \> Print \>**
    > **Print Preview** to see how your pages will look when printed.

<br/>

## Rule Lines

### Background

The Add-In can be configured to create notebook pages with 4 different rule line
styles or none at all.

### Issue: Rule lines do not display in OneNote 2016 or later

- Rule lines are not displayed in journal pages created by the Add-In with OneNote
2016 or later; works fine in earlier versions of OneNote.

- **UPDATE:** This issue has been resolved in [release 16.4.0].

### Workaround

- In OneNote, select **File \> Options \> Display** to change how OneNote looks.

- Check the following **Display** option:
    :white_check_mark: Create all new pages with rule lines

- If you have set the **Rule Lines** style to `None` in the Add-In's **Page Settings**,
you can still set rule lines for the current page via the OneNote **View \> Rule Lines**
menu option.

[release 16.4.0]:https://github.com/atrenton/MyJournal.Notebook/releases/tag/v16.4.0
