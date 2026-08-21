**[RU](README.md)**

# Zero-minute meeting reminders in Outlook classic

This VBA code automatically replaces the organizer-defined reminder in your
copy of a meeting with one built-in Outlook reminder at the meeting start time.
The code never calls `Send` and does not send meeting updates to attendees.

Only timed meetings in the current Outlook profile's primary calendar are
processed. Personal appointments without attendees, all-day events, canceled or
declined meetings, and meetings that have already started are left unchanged.

If reminders are disabled for an eligible meeting, the code enables them and
sets the value to `0 minutes`. This is intentional: the requirement to always
show a reminder at the meeting start takes precedence over a previous local
choice to disable the reminder.

## Requirements

- Windows and **Outlook classic**. New Outlook does not support VBA macros.
- Permission to run VBA macros in Outlook.
- Outlook reminders must be enabled. Windows notification settings must not
  block Outlook notifications either.

## Installation

1. Make sure you are using Outlook classic. If the **New Outlook** toggle is
   visible in the upper-right corner, turn it off.
2. Press `Alt+F11` in Outlook to open Microsoft Visual Basic for Applications.
3. In **Project Explorer**, expand the Outlook project (usually
   `Project1 (VbaProject.OTM)`) → `Microsoft Outlook Objects`, then double-click
   `ThisOutlookSession`.
4. If `ThisOutlookSession` already contains code, make a backup first. A module
   cannot contain two handlers for the same event: merge the existing contents
   of `Application_Startup`, `Application_MAPILogonComplete`, and
   `Application_Quit` into the corresponding handlers supplied here. Other
   existing procedures can remain alongside this code.
5. Copy the complete contents of `ThisOutlookSession.vba` into
   `ThisOutlookSession`.
6. In the VBA editor, select **Debug → Compile VBAProject**. Compilation must
   complete without errors.
7. Press `Ctrl+S`, close the VBA editor and Outlook, then start Outlook again.
   This one restart completes installation; after that, the code attaches
   automatically whenever Outlook starts.

## Macro security

Open:

`File → Options → Trust Center → Trust Center Settings → Macro Settings`.

Do not select **Enable all macros**. The preferred setup is:

1. Create or obtain a code-signing certificate. For a personal installation,
   you can use `SelfCert.exe`; it is a Microsoft Office component that may need
   to be added to some Office installations.
2. In the VBA editor, open **Tools → Digital Signature**, select the certificate,
   and save the project.
3. In Trust Center, select the option that permits digitally signed macros and
   disables unsigned macros.
4. At the first trust prompt, verify the certificate and trust the publisher.

If these settings are controlled by organizational policy, your Microsoft 365
or Windows administrator must provide the required signature or permission.

## Initial verification

1. Create a test meeting with an attendee 20–30 minutes in the future and give
   it a 15-minute reminder.
2. Press `Alt+F8`, select `ZeroReminder_RepairNow` (it may appear as
   `Project1.ThisOutlookSession.ZeroReminder_RepairNow`), and select **Run**.
   This manual command is for diagnostics only; normal operation does not depend
   on it.
3. Open the test meeting. Its reminder field should show `0 minutes`.
4. To inspect the diagnostic result, press `Alt+F11` and then `Ctrl+G`. The
   **Immediate** window will show `Changed`, `Skipped`, and `Errors` counters.
5. Close and reopen Outlook, create another test meeting with a 15-minute
   reminder, and confirm that the value changes to 0 automatically.
6. Wait for the test meeting to start and confirm that Outlook displays its
   built-in reminder.

## How it works

- `Application_Startup` attaches event handlers to the primary calendar whenever
  Outlook starts. `Application_MAPILogonComplete` retries automatically if the
  MAPI store was not ready during early startup.
- The startup repair uses separate restricted collections for future one-time
  meetings and recurring-series masters. If a store does not support the
  filter, the code automatically falls back to a safe full-calendar scan.
- `ItemAdd` fixes newly created and accepted meeting invitations.
- `ItemChange` restores the zero-minute value after an organizer updates a
  meeting.
- A disabled reminder on an eligible meeting is re-enabled and set to 0 minutes.
- Before saving, the code compares values and does not write an already-correct
  meeting again.
- For recurring meetings, the series master and existing modified occurrences
  are repaired. Deleted and canceled occurrences are skipped.
- A property-check or save error for one item is written to the **Immediate**
  window and does not stop other meetings from being processed. A COM collection
  enumeration error ends the current pass and is also logged.

## Limitations

- Outlook must be running when a meeting starts. VBA does not run while Outlook
  is closed.
- Windows **Do not disturb**, Focus Assist, or disabled notifications can hide
  the notification.
- A meeting must synchronize into the calendar before its start time.
- Outlook may not fire `ItemAdd` when many items are added in one operation.
  After a large initial synchronization, run `ZeroReminder_RepairNow` or restart
  Outlook; the startup repair will fix meetings that were missed.
- Only the current profile's primary calendar is modified. Shared calendars and
  primary calendars of other connected accounts are not processed.
- Editing VBA code invalidates its digital signature. Sign the project again
  after making changes.

## Show reminders above other windows

If the option is available in your Outlook version, open
`File → Options → Advanced → Reminders` and enable
**Show reminders on top of other windows**.

## Uninstallation

1. Press `Alt+F11` and open `ThisOutlookSession`.
2. Remove the added code or restore the backup created before installation.
3. Save the VBA project and restart Outlook.

Reminders already changed to 0 remain at 0; removing the code does not restore
the organizer's original values.
