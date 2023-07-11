# Bank Statement Extractor

Bank Statement Extractor is a Python application that allows users to extract their bank statements from a specific website. It combines web scraping using `requests` and `selenium` libraries with a Streamlit frontend for user interaction.

## Features

- Scrapes bank statements from a specified website.
- Allows users to provide their login credentials (username and password) via an Excel file.
- Extracts statements based on the user-provided date range.
- Presents the extracted statements in a user-friendly format.

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your-username/bank-statement-extractor.git

2. Navigate to the project directory:

   ```bash
   cd bank-statement-extractor
3.Install the required dependencies:
  ```bash
    pip install -r requirements.txt
  ```

### Usage
1.Prepare an Excel file containing the login credentials. The file should have the following structure:
| Username    | Password    |
|-------------|-------------|
| your_username  | your_password  |

2.Launch the Streamlit application:
```
streamlit run main.py

```
3.Open the application in your web browser (the terminal will display the URL).

4.Use the provided user interface to upload the Excel file and specify the desired date range for statement extraction.

5.Click the "Extract Statements" button to initiate the extraction process.

6.Once the process is complete, the extracted statements will be displayed on the screen.
### Technologies Used
  * Python
  * Streamlit
  * Selenium
  * Requests
