# scrunas

**scrunas** is a Windows console (text-based, command-line) program that enables, disables, or displays the status of the **Run as administrator** setting for a Windows shortcut (.lnk) file.

## AUTHOR

Bill Stewart - bstewart at iname dot com

## LICENSE

**scrunas** is covered by the GNU Public License (GPL). See the file `LICENSE` for details.

## USAGE

Command-line parameters, except for the account name, are case-sensitive.

`scrunas` [[`--enable` | `--disable`] [`--force`]] _filename_

## PARAMETERS

Specify `--enable` (or `-e`) to enable the **Run as administrator** setting, or specify `--disable` (or `-d`) to disable the setting. The `--force` (or `-f`) option updates the shortcut (.lnk) file even if the requested setting is already configured.

Omit `--enable` (`-e`) or `--disable` (`-d`) to display the current state of the **Run as administrator** option for the shortcut (.lnk) file.

The `--quiet` (or `-q`) option prevents output.

## EXIT CODES

If you omit `--enable` (`-e`) or `--disable` (`-d`), the program will exit with an exit code of 1 if the **Run as administrator** setting is enabled for the shortcut file, or 0 otherwise. Any other exit code indicates an error.

If you specify either `--enable` (`-e`) or `--disable` (`-d`), the program will exit with an exit code of 0 for success, or non-zero for an error.

## TECHNICAL DETAILS

See Raymond Chen's blog posting:

How do I mark a shortcut file as requiring elevation? (https://devblogs.microsoft.com/oldnewthing/20071219-00/?p=24103)

At the top of that post, Raymond writes:

> Specifying whether elevation is required is typically something that is the responsibility of the program. This is done by adding a `requestedExecutionLevel` element to your manifest. ... But if the program you're running doesn't have such a manifest--maybe it's an old program that you don't have any control over--you can create a shortcut to the program and mark the shortcut as requiring elevation.

He's right, of course: The application should have a manifest that says elevation is required. But for those edge cases (such as an installer for an old program that creates shortcuts), you can use **scrunas** to set the **Run as administrator** option for shortcut (.lnk) files where needed.
