import pypandoc
import os

def convert_doc(input_file, output_format):
    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"The file {input_file} does not exist.")
    
    try:
        output = pypandoc.convert_file(input_file, output_format)
        return output
    except Exception as e:
        print(f"Error during conversion: {e}")
        return None

# Example usage
input_file = r'C:\Users\Asus\Desktop\Code Files\Resume_Scorecard\extracted_files\20241212131520884283_Naukri_SANJAYKUMARGG[3y_0m](1) - Copy.doc'
output_format = 'plain'  # or 'pdf', 'html', etc.
converted_content = convert_doc(input_file, output_format)

if converted_content:
    print(converted_content)
