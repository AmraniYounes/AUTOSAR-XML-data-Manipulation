# Analysis of AUTOSAR-Related Data Stored in XML Files

## Description
This project involves manipulating AUTOSAR-related data stored in XML files. The main tasks include:
- Implementing a class to print the ports of a specific SWC (Software Component) to the CLI.
- Adding operator overloading to merge data from different XML files.
- Generating an Excel sheet for improved viewing and sorting of data.
- Creating a class to retrieve newly added ports from another XML or generated Excel file.
- Adding operator overloading to this class as well.
- Generating a report that highlights the differences between XML files.
- Implementing a class that converts XML files into a result file showing the differences via command-line args.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Example XML Structure](#example-xml-structure)

## Installation

### Prerequisites
- Python 3.9 or higher
- `pip` (Python package installer)

### Dependencies
- `pandas`
- `lxml` (for advanced XML parsing)

### Setup
1. Extract the folder from the archive you received.
2. Navigate to the project directory:
    ```bash
    cd path/to/extracted-folder/
    ```
3. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

**Note:** When running the scripts, you should only use the file names and not their full paths, provided you are in the project directory.

### Running the Project
- Ensure you're in the `src` directory before running the scripts.

### Examples
- To view SWC ports:
    ```python
    obj = SwcPortViewer(xml_file)
    df = obj.get_all_ports_of_swc('Diag')
    print(df)
    ```
    ```bash
    # This command runs the script to view ports for a specific SWC in the XML file.
    python SwcPortViewer.py
    ```
- To generate a resulting difference between XML files:
    ```python
    obj = CLIConverter(xml1, xml2)
    obj.generate_excel_from_CLI('Diag_Eol')
    ```
    ```bash
    # This command compares two XML files and generates an Excel file with the differences.
    python CLIConverter.py -i xml1_file.xml xml2_file.xml -o result.xlsx
    ```

## Example XML Structure

Below is an example of the XML structure used in this project:

```xml
<ROOT>
    <ITEM>
        <ID>123</ID>
        <PORTS>
            <PORT>
                <NAME>PortA</NAME>
                <SWC>Diag</SWC>
            </PORT>
            <PORT>
                <NAME>PortB</NAME>
                <SWC>Com</SWC>
            </PORT>
        </PORTS>
    </ITEM>
    <ITEM>
        <ID>456</ID>
        <PORTS>
            <PORT>
                <NAME>PortC</NAME>
                <SWC>Diag</SWC>
            </PORT>
        </PORTS>
    </ITEM>
</ROOT>
