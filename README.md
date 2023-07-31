# Text 2 Video Narrator

This PowerShell script allows you to create videos from specially formatted text files using AWS Polly for text-to-speech conversion. The script reads the input text file, generates PowerPoint slides with appropriate content, and then exports them as video files. The process involves the use of AWS Polly for voice synthesis and the AWS CLI configured with a user having rights to Polly.

## Prerequisites

Before using this script, ensure the following prerequisites are met:

1. PowerShell v7.2 or later is installed on your system.
2. AWS CLI is installed and configured with a user having access to AWS Polly.
3. Microsoft Office PowerPoint 2012 or greater.
4. Drop ffmpeg.exe into the root folder of the script, or if already installed, adjust the script to reference it appropriately.
5. The required templates, audio files, and avatar images are available in the specified folders as defined in the script.

## Usage

1. Place the specially formatted text file in the designated `Text` folder (`C:\Script\Text`).
2. Make sure the necessary templates and audio files are available in the designated folders as defined in the script (`C:\Script\template`, `C:\Script\POLLY`, etc.).
3. Open PowerShell and navigate to the script directory (`C:\Script`).
4. Execute the script by running the PowerShell file:

```powershell
pwsh .\GenerateVideo.ps1
```

5. The script will read the input text file and generate PowerPoint slides accordingly.
6. A final PowerPoint presentation will be saved in the `ARCHIVE` folder (`C:\Script\ARCHIVE`) with the name `Story<UnixTime>.pptx`.
7. The script will export the video, thumbnail, and description text to the `VIDEOS` folder (`C:\Script\VIDEOS`) with appropriate filenames.

## Input File Format

The script expects the input text file to be in a specific format.  See the examples in the test file folder.  Please make sure the text file follows this format for the script to function correctly.

## Configuration

- To change the location of various folders and files, update the corresponding variables in the script.

## Note

- The script uses AWS Polly for text-to-speech synthesis. Ensure you have the necessary credentials and access to AWS Polly to use this functionality.  No other permissions are needed.

## License

This script is licensed under the MIT License. You are free to modify and distribute the script as per the terms of the MIT License.

## Disclaimer

This script is provided as-is and without any warranties. The author is not responsible for any damage or loss caused by using this script.
