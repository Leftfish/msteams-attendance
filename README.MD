# MS Teams attendance statistics generator

A very raw command-line tool that generates a CSV file that sums up multiple attendance reports exported from Microsoft Teams. It's currently only capable of parsing MS Teams with Polish locale and I only tested it under Windows 10 (probably won't work on MacOS or Linux because the paths are written in a Windows-specific way). I might try to improve it when I find some time to procrastrinate.

Sample input:

![A CSV report exported from MS Teams](sample_input.png "Sample input")

Sample output:

![A CSV report with attendance summary](sample_output.png "Sample output")