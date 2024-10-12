import os
import platform
import pandas as pd
import xml.etree.ElementTree as ET

# XML files directory
xml_dir = "../xml_files/"
xml_file_1_path = f"{xml_dir}dids_swc_mapping.xml"
xml_file_2_path = f"{xml_dir}rids_swc_mapping.xml"

class SwcPortViewer:
    '''
    A class for analyzing AUTOSAR-related XML files.
    
    Each instance of this class represents an XML file related to AUTOSAR data.
    The class provides functionality to analyze XML data related to Software Components (SWCs) and
    generate excel files on demand.

    Attributes:
        file (str): The path to the XML file to use.
        tree (xml.etree.ElementTree.ElementTree): The parsed XML tree.
        root (xml.etree.ElementTree.Element): The root element of the XML tree.
    '''


    def __init__(self, file):
        '''
        Initializes a SwcPortViewer instance.
        
        Parameters:
            file (str): The file path of the xml file to use.
        Raises:
            ValueError: If the file type is not XML.
        '''

        # Check if the XML file exists
        if not os.path.exists(file):
            raise FileNotFoundError(f"The file '{file}' does not exist.")
        
        self.file = file

        # Parse and get the root of the XML if the file type is 'xml'.
        self.file_type = file.split('.')[-1]
        if self.file_type == 'xml':
            self.tree = ET.parse(file)  
            self.root = self.tree.getroot()
        else:
            raise ValueError(f"unsupported file type {self.file}")
        

    def __add__(self, other):
        '''
        Enables operator overloading for the '+' sign to merge two SwcPortViewer instances.

        This method allows combining data from two SwcPortViewer instances into a new instance.

        Args:
            other (SwcPortViewer): An instance of the SwcPortViewer class to merge with the current instance.
        Returns:
            merged_instance (SwcPortViewer): A new instance of SwcPortViewer with merged data from both instances.
        Raises:
            TypeError: If the other operand is not an instance of SwcPortViewer.
        '''
        if not isinstance(other, SwcPortViewer):
            raise TypeError("Operands must be instances of SwcPortViewer")
        
        df1 = self.convert_xml_to_df()
        df2 = other.convert_xml_to_df()

        # Concatenate the DataFrames and set index to start from 1
        combined_df = pd.concat([df1, df2]).reset_index(drop=True)
        combined_df.index += 1
        
        # Create a new SwcPortViewer instance with the combined data
        merged_instance = SwcPortViewer.__new__(SwcPortViewer)

        # Give new attribute to distinguish between original and merged DataFrames
        merged_instance.combined_df = combined_df

        # Initialize other attributes for safety
        merged_instance.file = None
        merged_instance.root = None
        merged_instance.tree = None
        return merged_instance


    def convert_xml_to_df(self):
        '''
        Converts the XML file associated with this instance into a DataFrame.

        If the instance has an overloaded DataFrame (from merging), it returns that. 
        Otherwise, it converts the XML file to create a DataFrame.
        
        Args:
            None.
        Returns:
            df (DataFrame): A DataFrame representation of the XML file's data, including ID, PORT, and SWC.
        '''
        data = [] # Initialize a list to hold the data
        if hasattr(self, 'combined_df'):
            return self.combined_df

        # Traverse the XML tree starting from ITEM (the subroot)
        for item in self.root.findall('ITEM'):
            id = item.find('ID').text if item.find('ID') is not None else None
            
            id = hex(int(id))[2: ].upper() # Format the ID to be in this format : 0x0FD1
            id = id.zfill(4)               # You can alse remove the .upper() for an ID like this : 0x0fd1 
            id= f"0x{id}"                  # 

            for port in item.find('PORTS').findall('PORT'):
                name = port.find('NAME').text if port.find('NAME') is not None else None
                swc = port.find('SWC').text if port.find('SWC') is not None else None
                
                # Append data to the list
                data.append({'ID': id, 'PORT': name, 'SWC': swc})

        # Return the DataFrame
        df = pd.DataFrame(data)
        return df
        

    def print_dataframe_of_xml(self):
        '''
        Prints the DataFrame created from the XML file associated with this instance.

        This method converts the XML file to a DataFrame and prints it to the CLI.
        It sets display options to ensure that the entire content of the DataFrame is shown.

        Args:
            None.
        '''

        # Set pandas options to display the full DataFrame
        pd.set_option('display.max_colwidth', None)
        pd.set_option('display.max_rows', None)
        
        # Convert XML to DataFrame
        df = self.convert_xml_to_df()
        
        # Print the DataFrame
        print(df)
        
        # Reset specific pandas options to their default values
        pd.reset_option('display.max_colwidth')
        pd.reset_option('display.max_rows')


    def get_all_ports_of_swc(self, swc):
        '''
        Returns the ports of the specified SWC.
        
        Locates and returns all ports corresponding to the given SWC.
        The index of the returned DataFrame is reset to start from 1 for optimized memory usage
        and easier readability. 
        
        Args:
            swc (str): The name of the software component (SWC) for which to retrieve the ports.
        Returns:
            swc_ports (DataFrame): A DataFrame containing the ports of the specified SWC, with the index
            to start from 1.
        Notes:
            - This method is case sensitive.
        '''

        # Check if we're dealing with a merged instance
        if hasattr(self, 'combined_df'):
            df = self.combined_df
        else:
            df = self.convert_xml_to_df()
        
        # Filter DataFrame and reset index and optimize memory usage and ease readability
        swc_ports = df.loc[df['SWC'] == swc].reset_index(drop=True)
        
        # Adjust index to start from 1
        swc_ports.index += 1

        # Early return if DataFrame is empty
        if swc_ports.empty:
            print(f'\nNo ports found for {swc}.')
            return
        
        # Set display option incase the df is printed
        pd.set_option('display.max_colwidth', None)
        pd.set_option('display.max_rows', None)

        return swc_ports
    
    
    def generate_excel(self):
        '''
        Generate an Excel file from a DataFrame and apply formatting using convert_to_excel().

        This method converts the internal Dataframe into an Excel sheet, applies specific 
        formatting to the headers, rows and columns, by calling convert_to_excel()
        and saves the result as an '.xlsx' file.

        Args:
            None.
        '''
        if hasattr(self, 'combined_df'):
            df = self.combined_df
            exl_file_name = 'dids_rids_merged_swc_mapping.xlsx'
        else:
            df = self.convert_xml_to_df()

            # Removing suffix and prefix of xml file name, result example : dids_swc_mapping
            xml_file_name = self.file.removeprefix(xml_dir).removesuffix('.xml') 

            # Adding the Excel extension
            exl_file_name = f'{xml_file_name}.xlsx'
        
        self.convert_to_excel(df, exl_file_name)


    def convert_to_excel(self, dataframe, file_name):
        '''
        Converts a DataFrame to an Excel file with custom formatting.

        This method creates an Excel file from the given DataFrame and applies formatting to improve readability.

        Args:
            dataframe (DataFrame): The DataFrame to convert and save as an Excel file.
            file_name (str): The name of the Excel file to create.
        '''
        try:
            # Create directory for excel files if it doesn't already exist
            excel_path = '../excel_files/'
            if not os.path.exists(excel_path):
                os.mkdir(excel_path)

            # Define full path for the Excel file
            full_path = os.path.abspath(os.path.join(excel_path, file_name))

            # Create an ExcelWriter object
            writer = pd.ExcelWriter(f'{excel_path}{file_name}', engine='xlsxwriter', mode='w')

            # Convert the DataFrame to an XlsxWriter Excel object
            dataframe.to_excel(writer, sheet_name='Sheet1', index=False)

            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # Define formatting styles
            header_bg_color = '#F10909'  # Red
            header_font_color = '#FFFFFF'  # White
            row1_bg_color = '#F2F2F2'  # Light grey
            row2_bg_color = '#FFFFFF'  # White
            border_color = '#000000'  # Black

            header_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 12,
                'bg_color': header_bg_color,
                'font_color': header_font_color,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'border_color': border_color
            })

            row1_format = workbook.add_format({
                'bg_color': row1_bg_color,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'border_color': border_color
            })

            row2_format = workbook.add_format({
                'bg_color': row2_bg_color,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'border_color': border_color
            })

            # Define the range of the table based on DataFrame shape
            (max_row, max_col) = dataframe.shape
            
            # Adjust max_row and max_col if necessary
            max_row = len(dataframe.dropna(how='all'))  # Only count rows with data
            max_col = len(dataframe.dropna(axis=1, how='all').columns)  # Only count columns with data

            table_range = f"A1:{chr(65 + max_col - 1)}{max_row + 1}"

            # Add a table to the worksheet with the DataFrame data
            worksheet.add_table(table_range, {
                'columns': [{'header': column} for column in dataframe.columns],
            })

            # Apply the header format to the header row
            worksheet.set_row(0, 25, header_format)
            
            # Apply alternating row formats
            for row in range(1, max_row + 1):
                if row % 2 == 0:
                    worksheet.set_row(row, None, row1_format)
                else:
                    worksheet.set_row(row, None, row2_format)

            # Adjust column width to fit content    
            worksheet.autofit()

            # Save the workbook
            writer.close()
            print(f"\n=> Excel file generated: {excel_path}{file_name}")

            # Open the Excel file on Windows
            if platform.system() == 'Windows':
                os.startfile(full_path)

        # Handle the case when the file is opened and the user tries to rerun the script
        except PermissionError:
            print(f"\nError: The file '{file_name}' is currently open. Please close it and try again.")
        
        except Exception as e:
            print(f"\nError: {e}")

# Test code 
if __name__ == "__main__":

    obj1 = SwcPortViewer(xml_file_1_path)
    obj2 = SwcPortViewer(xml_file_2_path)
    merged_obj = obj1+ obj2

    df = obj1.get_all_ports_of_swc('Diag')
    if not df.empty:
        print(df)
    # # obj1.get_all_ports_of_swc('Diag_Eol')
    # obj2.generate_excel()
    obj1.generate_excel()
    merged_obj.generate_excel()
    # merged_obj.print_dataframe_of_xml()
    print()
