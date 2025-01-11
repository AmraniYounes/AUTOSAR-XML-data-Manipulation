import os
import platform
import argparse
from SwcDiffManager import SwcDiffManager as SDM
from SwcPortViewer import SwcPortViewer as SPV

# XML files directory
xml_dir = "../xml_files/"
xml_file_1_path = f"{xml_dir}dids_swc_mapping.xml"
xml_file_2_path = f"{xml_dir}rids_swc_mapping.xml"


class CLIConverter(SDM, SPV):
    '''
    A class to handle the conversion and comparison of two XML files via command-line interface (CLI).

    This class inherits from both SwcPortViewer and SwcDiffManager to manage differences and perform operations on XML files.
    It initializes with two XML files, loads them into DataFrames, and generates a report Excel file.

    Attributes:
        xml1 (str): Path to the first XML file.
        xml2 (str): Path to the second XML file.
        file1_type (str): File type of the first file.
        file2_type (str): File type of the second file.
        dataframe1 (pd.DataFrame): DataFrame containing data from the first XML file.
        dataframe2 (pd.DataFrame): DataFrame containing data from the second XML file.
    '''


    def __init__(self, xml1, xml2):
        '''
        Initializes the CLIConverter instance with two XML files.

        Args:
            xml1 (str): Path to the first XML file.
            xml2 (str): Path to the second XML file.
        Raises:
            FileNotFoundError: If either files doesn't exist
            ValueError: If either file is not an XML file or XML files have the same name.
        '''
        self.xml1 = xml1
        self.xml2 = xml2

        # If files have same name raise ValueError
        if os.path.basename(xml1) == os.path.basename(xml2):
            raise ValueError('\nAttributes cannot have the same name')

        # Check if both files exist
        if not os.path.exists(xml1):
            raise FileNotFoundError(f"The file '{xml1}' does not exist.")
        if not os.path.exists(xml2):
            raise FileNotFoundError(f"The file '{xml2}' does not exist.")

        # Determine file types
        self.file1_type = xml1.split('.')[-1]
        self.file2_type = xml2.split('.')[-1]

        # Check if both files are XML
        if self.file1_type != 'xml' or self.file2_type != 'xml':
            raise ValueError('File type incompatible. Both files must be XML.')
        
        # Load XML files into DataFrames
        self.dataframe1 = self.load_file(xml1, self.file1_type)
        self.dataframe2 = self.load_file(xml2, self.file2_type)


    def load_file(self, file, file_type):
        '''
        Loads the specified file into a DataFrame using the superclass's load_file method.

        Args:
            file (str): The name to the file to be loaded.
            file_type (str): The type of the file (e.g., '.xml', '.xlsx').
        Returns:
            df (DataFrame): A DataFrame containing the data from the file.
        Raises:
            ValueError: If the file_type is unsupported by the superclass method.
        '''
        return super().load_file(file, file_type)
    

    def get_new_ports_in_swc(self, swc):
        '''
        Retrieves new ports for a specified SWC that have been added in the second XML file compared to the first.

        This method overrides the superclass method to ensure the data is processed according to this class's 
        specific requirements.

        Args:
            swc (str): The SWC for which to get new ports.

        Returns:
            new_ports (DataFrame): A DataFrame containing the new ports for the specified SWC.
            If no new ports are found, an empty DataFrame is returned.
        '''
        return super().get_new_ports_in_swc(swc) 
    

    def convert_to_excel(self, dataframe, file_name):
        '''
        Converts a DataFrame to an Excel file and applies custom formatting.

        This method overrides the superclass method to ensure compatibility with this class's requirements.

        Args:
            dataframe (DataFrame): The DataFrame to be converted to an Excel file.
            file_name (str): The name of the Excel file to save.
        '''
        return super().convert_to_excel(dataframe, file_name)


    def generate_excel_from_CLI(self, swc):
        '''
        Generates an Excel file from XML files specified via command-line arguments.

        This method uses the argparse library to parse command-line arguments for input XML files
        and an output Excel file path. It retrieves new ports for a specified SWC and saves the result
        to an Excel file.

        Args:
            swc (str): The SWC for which to get new ports and generate the Excel report.
        '''
        try:
            # Create an argument parser
            parser = argparse.ArgumentParser(description="Generate Excel file from XML files.")
            
            # Add arguments for two input XML files and output Excel file
            parser.add_argument('-i', '--input', nargs=2, required=True, help="Paths to the two input XML files.")
            parser.add_argument('-o', '--output', required=True, help="Path to the output Excel file.")
            
            # Parse arguments
            args = parser.parse_args()

            # Extract input and output file names
            xml1, xml2 = args.input
            output_excel = args.output

            # Define full path for the XML adn Excel files
            xml1 = f'../xml_files/{xml1}'
            xml2 = f'../xml_files/{xml2}'
            full_path = os.path.abspath(os.path.join('../excel_files/', output_excel))

            # Check if the output file is already open
            if os.path.exists(full_path):
                try:
                    # Try opening the file for writing to check if it's open
                    with open(full_path, 'a'):
                        pass
                except PermissionError:
                    print(f"\nError: The file '{output_excel}' is currently open. Please close it and try again.")
                    return

            # Initialize CLIConverter with the provided XML files
            converter = CLIConverter(xml1, xml2)

            # Get the new ports in the specified SWC
            result_df = converter.get_new_ports_in_swc(swc)

            if result_df.empty:
                print(f'\nNo new ports for {swc} to generate an Excel file for')
                return

            # Convert the DataFrame to an Excel file
            if converter.convert_to_excel(result_df, output_excel):
                print(f'\n=> Excel file generated: ../excel_files/{output_excel}')
            
            # Open the Excel file on Windows
            if platform.system() == 'Windows':
                os.startfile(full_path)
                
        # Handle the case when the file is opened and the user tries to rerun the script
        except PermissionError:
            print(f"\nError: The file '{output_excel}' is currently open. Please close it and try again.")

        except Exception as e:
            print(f"\nError: {e}")


# Test code 
if __name__ == "__main__":

    # obj = CLIConverter(xml_file_1_path, xml_file_2_path)
    # obj.generate_excel_from_CLI('Diag_SWC')
    print()
