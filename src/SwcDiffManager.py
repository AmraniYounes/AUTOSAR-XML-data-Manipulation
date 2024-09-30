import os
import platform
import pandas as pd
import xml.etree.ElementTree as ET
from SwcPortViewer import SwcPortViewer as SPV

# XML files directory
xml_dir = "../xml_files/"
xml_file_1_path = f"{xml_dir}dids_swc_mapping.xml"
xml_file_2_path = f"{xml_dir}rids_swc_mapping.xml"

# Excel files directories
excel_dir = "../excel_files/"
excel_file_1_path = f"{excel_dir}dids_swc_mapping.xlsx"
excel_file_2_path = f"{excel_dir}rids_swc_mapping.xlsx"

class SwcDiffManager(SPV):
    '''
    Manage the differences between either two XML files or their two generated Excel files.

    This class compares two files (either XML or Excel), and inherits methods from `SwcPortViewer` 
    to facilitate the implementation of comparison features. The files can be of different types
    (e.g., one XML and one Excel), and the class handles the processing of these files accordingly.

    Attributes:
        file1 (str): Name of the first file (XML, Excel).
        file2 (str): Name of the second file (XML, Excel).
        dataframe1 (pd.DataFrame): DataFrame created from the first file.
        dataframe2 (pd.DataFrame): DataFrame created from the second file.
        file1_type (str): File type of the first file.
        file2_type (str): File type of the second file.
        is_merged (bool): Indicates whether the instance is an overloaded one.
    '''

    
    def __init__(self, file1, file2):
        '''
        Initializes a SwcDiffManager instance.
        
        Parameters:
            file1 (str): Name of the first file (XML or Excel).
            file2 (str): Name of the second file (XML or Excel).
        Raises:
            FileNotFoundError: If either files doesn't exist.
        '''
        self.file1 = file1
        self.file2 = file2

        # Check if both files exist
        if not os.path.exists(file1):
            raise FileNotFoundError(f"The file '{file1}' does not exist.")
        if not os.path.exists(file2):
            raise FileNotFoundError(f"The file '{file2}' does not exist.")

        # Determine the type of the file (xml, excel or other)
        self.file1_type = file1.split('.')[-1]
        self.file2_type = file2.split('.')[-1]

        # Load DataFrames for both files
        self.dataframe1 = self.load_file(self.file1, self.file1_type)
        self.dataframe2 = self.load_file(self.file2, self.file2_type)
        self.is_merged = False


    def __add__(self, other):
        '''
        Enables operator overloading for the '+' sign to merge two SwcDiffManager instances.

        This method allows combining data from two SwcDiffManager instances into a new instance.

        Args:
            other (SwcDiffManager): An instance of the SwcDiffManager class to merge with the current instance.
        Returns:
            merged_instance (SwcDiffManager): A new instance of SwcDiffManager with merged data from both instances.
        Raises:
            TypeError: If the other operand is not an instance of SwcDiffManager.
        '''
        if not isinstance(other, SwcDiffManager):
            raise TypeError("Operands must be instances of SwcDiffManager")
        
        # Merging the dataframes from both objects
        merged_df1 = pd.concat([self.dataframe1, other.dataframe1]).drop_duplicates().reset_index(drop=True)
        merged_df2 = pd.concat([self.dataframe2, other.dataframe2]).drop_duplicates().reset_index(drop=True)
        
        # Create a new SwcDiffManager instance
        merged_instance = SwcDiffManager.__new__(SwcDiffManager)
        
        # Set merged data
        merged_instance.dataframe1 = merged_df1
        merged_instance.dataframe2 = merged_df2

        # Set attributes for the new instance
        merged_instance.file1 = None
        merged_instance.file2 = None
        merged_instance.file1_type = None
        merged_instance.file2_type = None
        merged_instance.is_merged = True  

        # Returning a new SwcDiffManager object with the merged data
        return merged_instance


    def load_file(self, file, file_type):
        '''
        Loads a file into a DataFrame based on its type.

        This method reads data from a file (either XML or Excel) and converts it into a DataFrame.
        
        Args:
            file (str): The name to the file to load.
            file_type (str): The type of the file ('.xml' or '.xlsx').
        Returns:
            df (DataFrame): A DataFrame containing the data from the file.
        Raises:
            ValueError: If the file type is unsupported.
        '''
        if file_type == 'xml':
            # Parse the XML file and convert it to a DataFrame
            self.tree = ET.parse(file)  
            self.root = self.tree.getroot()
            df = self.convert_xml_to_df()

        elif file_type == 'xlsx':
            # Read the Excel file into a DataFrame
            df = pd.read_excel(file)

        else:
            # Raise an error for unsupported file types
            raise ValueError(f"Unsupported file type: {file_type}")
        
        return df
        

    def convert_xml_to_df(self):
        '''
        Converts the XML data from the loaded file into a DataFrame.

        This method calls the parent class's `convert_xml_to_df` method to perform the conversion and it is used in
        the load_file(file, file_type) method.

        Returns:
            df (DataFrame): A DataFrame containing the data extracted from the XML file.
        '''
        return super().convert_xml_to_df()
    

    def get_new_ports_in_swc(self, swc):
        '''
        Retrieves new ports for a specified SWC that have been added in the second file compared to the first.

        Args:
            swc (str): The SWC for which to get new ports.

        Returns:
            new_ports (DataFrame): A DataFrame containing the new ports in the specified SWC.
            If no new ports are found, an empty DataFrame is returned.
        '''
        df1_swc = self.dataframe1[self.dataframe1['SWC'] == swc]
        df2_swc = self.dataframe2[self.dataframe2['SWC'] == swc]
        
        if not df1_swc.empty and not df2_swc.empty:
            # SWC is in both files, return new ports in file2
            new_ports = df2_swc[~df2_swc.isin(df1_swc)].dropna()
            if new_ports.empty:
                print(f'\nNo new ports for {swc}')
                return pd.DataFrame()

        elif df1_swc.empty and not df2_swc.empty:
            # SWC only in file2, return all ports from file2
            new_ports = df2_swc

        elif not df1_swc.empty and df2_swc.empty:
            # SWC only in file1, return empty DataFrame
            print(f'\nNo new ports for {swc}')
            return pd.DataFrame()
        else:
            # SWC does not exist in either file
            print(f'\nSWC {swc} does not exist in either file')
            return pd.DataFrame()
        
        # Reset index for the result
        new_ports = new_ports.reset_index(drop=True)
        new_ports.index += 1

        # Set max display settings
        pd.set_option('display.max_colwidth', None)
        pd.set_option('display.max_rows', None)
        
        return new_ports
    

    def generate_report(self):
        '''
        Generates a report of new ports in second file and saves it to an Excel file with visual charts.

        Returns:
            None.
        '''
        try:
            # Get unique SWCs from both DataFrames in a numpy array
            swcs = pd.concat([self.dataframe1['SWC'], self.dataframe2['SWC']]).unique()

            report_data = []
            for swc in swcs:
                new_ports_df = self.get_new_ports_in_swc(swc)

                # Collect data if there are new ports    
                new_ports = len(new_ports_df)
                if new_ports > 0:
                    report_data.append({'SWC': swc, 'New Ports': new_ports})

            if not report_data:
                print("\nNo changes detected between the two files.")
                return

            # Create a DataFrame for the report
            df = pd.DataFrame(report_data)
            output_file = os.path.abspath('../excel_files/report.xlsx')
            output_dir = os.path.dirname(output_file)

            # Ensure the directory exists
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Write DataFrame to Excel and create visual charts
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Report', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Report']

                # Apply formatting to the header
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'center',
                    'fg_color': '#F10909',  # Red background
                    'font_color': '#FFFFFF',  # White font color
                    'border': 1
                })
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)

                # Apply formatting to the data rows
                alternating_row_format1 = workbook.add_format({
                    'bg_color': '#F2F2F2',  # Light grey background
                    'border': 1,
                    'valign': 'center',
                    'align': 'center'
                })
                alternating_row_format2 = workbook.add_format({
                    'bg_color': '#FFFFFF',  # White background
                    'border': 1,
                    'valign': 'center',
                    'align': 'center'
                })

                # Alternate the formatting for rows
                for row_num in range(1, len(df) + 1):
                    format_to_use = alternating_row_format1 if row_num % 2 == 0 else alternating_row_format2
                    for col_num in range(len(df.columns)):
                        worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], format_to_use)

                # Autofit columns
                for col_num, col_data in enumerate(df.columns):
                    max_length = max(df[col_data].astype(str).map(len).max(), len(col_data))
                    worksheet.set_column(col_num, col_num, max_length + 2)

                # Create and configure the column chart
                chart1 = workbook.add_chart({'type': 'column'})
                chart1.add_series({
                    'name': 'New Ports',
                    'categories': '=Report!$A$2:$A${}'.format(len(df) + 1),
                    'values': '=Report!$B$2:$B${}'.format(len(df) + 1),
                })
                chart1.set_title({'name': 'New Ports per SWC'})
                chart1.set_x_axis({'name': 'SWC'})
                chart1.set_y_axis({'name': 'Number of New Ports'})
                worksheet.insert_chart('E9', chart1)

                # Create and configure a pie chart for the percentage of new ports by SWC
                chart2 = workbook.add_chart({'type': 'pie'})
                chart2.add_series({
                    'name': 'New Ports Distribution',
                    'categories': '=Report!$A$2:$A${}'.format(len(df) + 1),
                    'values': '=Report!$B$2:$B${}'.format(len(df) + 1),
                })
                chart2.set_title({'name': 'New Ports Distribution per SWC'})
                worksheet.insert_chart('N9', chart2)

            print(f'\n=> Report generated: {output_file}')

            # Open the Excel file on Windows
            if platform.system().lower() == 'windows':
                os.startfile(output_file)

        # Handle the case when the file is opened and the user tries to rerun the script
        except PermissionError:
            print(f"\nError: The file '{output_file}' is currently open. Please close it and try again.")
        
        except Exception as e:
            print(f"\nError: {e}")


# Test code 
if __name__ == "__main__":    
    
    # obj1 = SwcDiffManager(xml_file_1_path, xml_file_2_path)
    # obj2 = SwcDiffManager(xml_file_1_path, xml_file_2_path)
    # obj3 = obj1 + obj2

    # df = obj1.get_new_ports_in_swc('jdd_2010')
    # if not df.empty:
    #     print(df)
    
    # df = obj3.get_new_ports_in_swc('Diag_Eol')
    # if not df.empty:
    #     print(df)
    # obj3.generate_report()

    print()