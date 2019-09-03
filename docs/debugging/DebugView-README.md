# About the DebugView-Filters.ini file

The **`DebugView-Filters.ini`** file is a [Sysinternals DebugView](https://docs.microsoft.com/en-us/sysinternals/downloads/debugview) filters file for debugging the **MyJournal.Notebook** add-in.

----------

### DebugView Menu Settings
- Capture
	- Select Capture Win32 (Ctrl+W)
	- Select Capture Events (Ctrl+E)

- Options
	- Deselect Win32 PIDs
	- Select Clock Time (Ctrl+T)
	- Select Show Milliseconds

- Computer
	- Select Connect Local

----------

### How to load the **`DebugView-Filters.ini`** file
1. From the DebugView menu, select Edit > Filter/Highlight... (Ctrl+L)
2. Click the Load button
3. Navigate to the repo `docs\debugging` subdirectory and select the `DebugView-Filters.ini` file
4. Click the Open button
5. Click the OK button
6. Save the filters configuration by closing and reopening DebugView

----------

### Debugging the MyJournal.Notebook Add-in
To debug the add-in, update the following ``appSettings`` values in ``App.config``:
Set ``key="Diagnostics.OutputWriter.Type.Name" value="TraceOutputWriter"``
Set ``key="Diagnostics.TraceSwitch.Level" value="Verbose"``

**NOTE:** When creating a ``Debug`` build, the ``App.config`` file will be automatically configured by the Microsoft VisualStudio [SlowCheetah](https://marketplace.visualstudio.com/items?itemName=vscps.SlowCheetah-XMLTransforms) package.