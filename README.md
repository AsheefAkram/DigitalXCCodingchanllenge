# Group Occurrences Counter

## Overview
This Python script reads data from an Excel sheet and identifies groups mentioned under the "Additional comments" column. The script lists all unique groups and counts the number of times each group has been mentioned. The results are then saved to a new Excel file.

## Features
- Reads employee data from an Excel file.
- Extracts and lists unique groups mentioned in the "Additional comments" column.
- Counts the number of occurrences for each group.
- Saves the results to a new Excel file.

## Installation

### Prerequisites
- Python 3.6+
- pandas library
- openpyxl library

### Steps

1. Clone the repository:

    ```sh
    git clone https://github.com/AsheefAkram/DigitalXCCodingchanllenge.git
    cd group-occurrences-counter
    ```

2. Create and activate a virtual environment:

    - **Windows:**

    ```sh
    python -m venv myenv
    myenv\Scripts\activate
    ```

    - **macOS/Linux:**

    ```sh
    python -m venv myenv
    source myenv/bin/activate
    ```

3. Install required libraries:

    ```sh
    pip install pandas openpyxl
    ```

## Usage

1. Prepare the input Excel file (`coding challenge test.xlsx`) with the "Additional comments" column.

2. Update the `input_file` and `sheet_name` variables in the script:

    ```python
    input_file = "path_to_your_file/coding challenge test.xlsx"
    sheet_name = 'Input Data sheet'
    ```

3. Run the script:

    ```sh
    python group_occurrences_counter.py
    ```

4. Check the generated output file (`Group_occurrences.xlsx`) for the group occurrences.

## Code Explanation

### `group_occurrences_counter.py`

- **extract_groups Function**:
    - Uses a regular expression to find lines matching the pattern `Groups : [code]<I>XXXX</I>[/code]`.
    - Splits the groups if multiple groups are mentioned, separated by commas.

- **count_group_occurrences Function**:
    - Reads the Excel file using `pandas`.
    - Extracts comments from the "Additional comments" column.
    - Uses the `extract_groups` function to get a list of all groups mentioned.
    - Counts the occurrences of each group using `Counter` from the `collections` module.
    - Creates a DataFrame from the count data and sorts it by the number of occurrences.
    - Saves the results to a new Excel file.

### Example Input and Output

- **Input File (`coding challenge test.xlsx`)**:
    ```excel
    | Additional comments                                  |
    |------------------------------------------------------|
    | Groups : [code]<I>Group A, Group B</I>[/code]        |
    | Groups : [code]<I>Group C</I>[/code]                 |
    | Groups : [code]<I>Group A</I>[/code]                 |
    | Groups : [code]<I>Group B, Group C, Group D</I>[/code]|
    ```

- **Output File (`Group_occurrences.xlsx`)**:
    ```excel
    | Group name | Number of occurrences |
    |------------|------------------------|
    | Group A    | 2                      |
    | Group B    | 2                      |
    | Group C    | 2                      |
    | Group D    | 1                      |
    ```

## Error Handling

- The script includes error handling to manage potential exceptions, such as invalid input or file handling issues.

## Contribution

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -m 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Create a Pull Request.

