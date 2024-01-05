# ExcelParser

- Have Node.js installed on local machine
- https://nodejs.org/en/download
- Have an editor installed
- Clone from Github
- Run command "npm i" in terminal to install

TDV PARSER:
    Pre-requisites:
    - Have private data installed
    - Export into excel with Transmitter Data Viewer
    - Rename file into the following format: "MMDDYYYY_0000_MMDDYYYY_0000" (time is in 24hr format)

    How to run:
    - Place formated file into "tdvData" folder
    - In terminal run command "npm run tdv"

INQUISITO PARSER:
    Pre-requisites:
    - Login to inquisito
    - Search for account associated with the data to be installed
    - Select "Phone Error Logs"
    - Inquisito data installed
    - Rename file into the following format: "MMDDYYYY_0000_MMDDYYYY_0000" (time is in 24hr format)

    How to run:
    - Place formated file into "inquisitoData" folder
    - In terminal run command "npm run inq"

BOTH:
    - to run both TDV and Inqusito run "npm run build"