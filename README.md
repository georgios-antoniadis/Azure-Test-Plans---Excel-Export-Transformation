### Description
A tool to process the test plans export from Azure's Test Plan feature, and get it to a format that is comfortable to work with locally and/or offline.

*Requires*
- Exported csv from Azure Test Plans

*Features*
- Remove unecessary columns
- Add new columns for local use
    - Test result
    - Comments
- Apply stylying
    - Header
    - Title cells
    - Column width
    - Text wrapping

*Please note*
- You will not be able to re-upload the processed test plans to azure
- The tool accepts as input a single file
    - To include test plans from multiple test suites, combine the exported .csv into one file.