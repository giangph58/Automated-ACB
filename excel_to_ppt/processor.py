from .utils import load_excel, process_dataframe, extract_province, \
                   find_table, update_table_with_data, \
                   remove_all_pictures, update_table_with_images, \
                   write_district, write_period, extract_period

from pathlib import Path
from pptx import Presentation
import os

def generate_ppt(input_file, output_path, template_path, image_path):
    # Constants
    phrases = ['nắng', 'không nắng', 'có mưa', 'không mưa', 'mây', 'không mây', 'dông', 'không dông']
    image_mappings = {
            frozenset(['nắng', 'mây', 'không mưa']): 'hs_hc_nr_nt.png',
            frozenset(['nắng', 'mây', 'có mưa']): 'hs_hc_hr_nt.png',
            frozenset(['nắng', 'không mây', 'có mưa', 'dông']): 'hs_nc_hr_ht.png',
            frozenset(['nắng', 'không mây', 'có mưa']): 'hs_nc_hs_nt.png',
            frozenset(['nắng', 'không mây', 'không mưa']): 'hs_nc_nr_nt.png',
            frozenset(['không nắng', 'có mây', 'không mưa']): 'ns_hc_nr_nt.png',
            frozenset(['không nắng', 'không mây', 'có mưa']): 'ns_nc_hr_nt.png',
        }
    
    df = load_excel(input_file)
    processed_df = process_dataframe(df)
    districts = processed_df.iloc[:, 0].unique()
    output_files = []

    for district in districts:
        district_df = processed_df[processed_df.iloc[:, 0] == district]
        ppt = Presentation(template_path)
        slide = ppt.slides[0]
        table = find_table(slide)
        
        # Update table
        update_table_with_data(table, district_df)
        remove_all_pictures(slide)
        update_table_with_images(slide, table, phrases, image_mappings, image_path)
        
        # Update district and period
        district_name = district.split('-')[0].strip().title()
        write_district(slide, district_name)
        write_period(slide, extract_period(district_df))

        # Save the PowerPoint file for the current district
        province = extract_province(os.path.basename(input_file))
        output_file = Path(output_path) / f'ACB_{province}_{district_name}.pptx'
        output_files.append(output_file)
        ppt.save(output_file)
        print(f"Saved presentation for {district_name} at {output_file}")
    
    return output_files
