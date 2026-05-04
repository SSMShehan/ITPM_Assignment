# ITPM Assignment 1 - Transliteration Testing (IT23836068)

This repository contains the Playwright test automation scripts tailored for IT23836068 for testing the accuracy of the Singlish-to-Sinhala chat translator.

## Prerequisites
- Python 3.11 or 3.12 installed
- Google Chrome installed

## Installation

1. Open a Command Prompt or Terminal in the `IT23836068_Playwright_Project` folder.
2. Install the required Python packages by running:
   ```bash
   pip install -U pip
   pip install playwright openpyxl
   ```
3. Install the required Playwright browsers:
   ```bash
   playwright install
   ```

## Creating Test Cases

Run the test case generator to create your 50 new negative test cases Excel file:
```bash
python generate_test_cases.py
```
This will create `IT23836068_Assignment 1 - Test cases.xlsx`.

## Running the Tests

To execute the test automation script, run the following command in the terminal:

```bash
python test_automation.py --excel "IT23836068_Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 5000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open
```

### What this script does:
1. It reads the test cases from `IT23836068_Assignment 1 - Test cases.xlsx`.
2. It launches a browser using Playwright and navigates to the Chat Translator.
3. It iterates through the 50 negative test cases.
4. For each test case, it inputs the Singlish phrase, waits for the result, and extracts the actual output.
5. It compares the actual output with the expected output and records the "Actual Output" and "Status" (PASS/FAIL) directly in the Excel file.

## Expected Results
Since these are negative test cases designed to test the limitations of the transliteration system, the script will record the actual outputs and mark the statuses. Many cases are expected to FAIL, highlighting areas for improvement in the transliteration model.

## Troubleshooting
If you encounter `UI Error` statuses due to slow loading or timeouts, run the retry script:
```bash
python retry_ui_errors.py
```
