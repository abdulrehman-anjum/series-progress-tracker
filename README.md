# series-progress-tracker
A VBScript script that helps you track the progress of the TV series you have downloaded in a specific folder. It automatically renames the shortcuts of the series based on your viewing progress, allowing you to easily identify which episodes you have watched. 

This repository contains a VBScript script that helps you track the progress of the TV series you have downloaded in a specific folder. It automatically renames the shortcuts of the series based on your viewing progress, allowing you to easily identify which episodes you have watched.

## Prerequisites

To use this script, you need to have the following:

- Windows operating system.
- A folder containing the TV series episodes you have downloaded.
- Shortcut files (.lnk) for each TV series folder in the same directory.

## Usage

1. Download the `series_progress_tracker.vbs` file from this repository.
2. Open the `series_progress_tracker.vbs` file using a text editor.
3. Replace the value of the `objStartFolder` variable with the path to the folder where your TV series episodes and their corresponding shortcut files are located. For example:
   ```
   objStartFolder = "C:\Path\To\Your\Series\Folder"
   ```
4. Save the changes to the `series_progress_tracker.vbs` file.
5. Double-click the `series_progress_tracker.vbs` file to execute the script.
6. The script will start tracking the progress of your TV series episodes and updating the shortcut names accordingly.
7. View the updated shortcut names to determine your progress in each series.

**Note:** The script assumes a specific naming convention for your TV series episodes, where the file names follow the format "01x03 - Episode Name.mkv". The script extracts the season number and episode number from the file names to calculate the progress. Make sure your episode files are named correctly to ensure accurate tracking.

The script will run continuously and update the shortcut names every 10 seconds. It calculates the remaining episodes, season number, and progress percentage for each series folder and incorporates this information into the shortcut names.

## Example

Let's say you have a TV series called "Breaking Bad" in your folder, and you have watched 3 episodes of Season 1 (out of 10 total episodes). The original shortcut name is "Breaking Bad". After executing the script, the shortcut name will be updated to "(7) Breaking Bad [30%].lnk", indicating that there are 7 remaining episodes, it belongs to Season 1, and you have completed 30% of the season.

## Customization

If you want to modify the script or adapt it to your specific needs, you can open the `series_progress_tracker.vbs` file in a text editor and make the necessary changes. The script is well-commented, providing explanations for each section of the code.

## Contributions

Contributions to this project are welcome! If you have any suggestions, improvements, or bug fixes, feel free to submit a pull request.

## Disclaimer

The script provided in this repository is for personal use and should be used responsibly and in compliance with the terms and conditions of the content you have downloaded. The author is not responsible for any misuse or violation of copyright laws that may arise from the usage of this script.

## Support

For any questions, issues, or support regarding this script, please open an issue in the [issue tracker](https://github.com/abdulrehman-anjum/series-progress-tracker/issues).
