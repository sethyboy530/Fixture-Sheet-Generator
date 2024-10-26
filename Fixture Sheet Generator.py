import pandas as pd

# Function to get user input in a more batch-oriented way, similar to the GrandMA2 Patch function
def patch_fixtures():
    fixture_list = []
    
    while True:
        print("Patch new fixtures:")
        
        # User specifies the fixture details
        manufacture = input("Enter the Fixture Manufacture: ")
        model = input("Enter the Fixture Model: ")
        channels = int(input("Enter the number of DMX channels: "))
        universe = int(input("Enter the Universe: ") or 1)
        starting_address = int(input("Enter the Starting DMX Address: ") or 1)
        quantity = int(input("Enter the Number of Fixtures to Patch: "))
        fixture_number = int(input("Enter the Starting Fixture Number (or press Enter to skip): ") or 0)
        ma_channel = int(input("Enter the Starting MA Channel Number (or press Enter to skip): ") or 0)
        position = input("Enter the Position (e.g., Pipe 1, or press Enter to skip): ")
        notes = input("Enter any optional notes (or press Enter to skip): ")

        # Automatically assign DMX addresses and patch the fixtures
        valid_batch = True
        batch_fixtures = []
        for i in range(quantity):
            current_address = starting_address + (i * channels)
            if current_address + channels - 1 > 512:
                print(f"Error: Fixture {i + 1} in the batch exceeds the DMX limit of 512 channels in universe {universe}.")
                print("Please restart the batch with valid values. DMX Universe channel limit exceeded.")
                valid_batch = False
                break
            
            # Check for patch collisions
            for fixture in fixture_list:
                existing_universe, existing_address, _, _, existing_channels, *_ = fixture
                if existing_universe == universe and not (current_address + channels - 1 < existing_address or current_address > existing_address + existing_channels - 1):
                    print(f"Error: Fixture {i + 1} in the batch has a patch collision with an existing fixture in universe {universe}.")
                    print("Please restart the batch with valid values to avoid patch collisions.")
                    valid_batch = False
                    break
            
            if not valid_batch:
                break
            
            current_fixture = [
                universe, current_address, manufacture, model,
                channels, fixture_number + i if fixture_number != 0 else '', ma_channel + i if ma_channel != 0 else '',
                position, '', notes
            ]
            batch_fixtures.append(current_fixture)
        
        if valid_batch:
            fixture_list.extend(batch_fixtures)
        else:
            continue
        
        more_batches = input("Do you want to patch another batch of fixtures? (yes/no): ").strip().lower()
        if more_batches != 'yes':
            break

    return fixture_list

# Function to create and save the DMX spreadsheet
def create_dmx_spreadsheet(data):
    # Allow user to define custom columns
    use_custom_columns = input("Do you want to create your own columns? (yes/no): ").strip().lower()
    if use_custom_columns == 'yes':
        custom_columns = []
        while True:
            column_name = input("Enter a column name (or press Enter to finish): ")
            if not column_name:
                break
            custom_columns.append(column_name)
        if not custom_columns:
            columns = [
                "Universe", "DMX Address", "Manufacture", "Model", 
                "Mode (DMX Channels)", "MA Fixture #", "MA Channel #", 
                "Position", "Unit # on position", "Notes"
            ]
        else:
            columns = [
                "Universe", "DMX Address", "Manufacture", "Model", 
                "Mode (DMX Channels)", "MA Fixture #", "MA Channel #", 
                "Position"
            ] + custom_columns + ["Notes"]
    else:
        columns = [
            "Universe", "DMX Address", "Manufacture", "Model", 
            "Mode (DMX Channels)", "MA Fixture #", "MA Channel #", 
            "Position", "Unit # on position", "Notes"
        ]
    
    # Adjust data to match columns
    data = [fixture + [''] * (len(columns) - len(fixture)) for fixture in data]
    
    sheet_name = input("Enter the name for the Excel file and sheet: ")
    output_path = f'C:/Users/stmic/Downloads/{sheet_name}.xlsx'
    
    # Create a DataFrame from the user's input data
    df_new = pd.DataFrame(data, columns=columns)
    df_new.sort_values(by=['Universe', 'DMX Address'], inplace=True)
    df_new.reset_index(drop=True, inplace=True)
    
    # Add universe labeling rows
    labeled_rows = []
    for universe, group in df_new.groupby('Universe'):
        # Add label for the start of the universe
        labeled_rows.append([f'Universe {universe}'] + [None] * (len(columns) - 1))
        # Add the group of fixtures for the universe
        labeled_rows.extend(group.values.tolist())
        # Add label for the end of the universe
        labeled_rows.append([f'Universe {universe}'] + [None] * (len(columns) - 1))
    df_new = pd.DataFrame(labeled_rows, columns=columns)
    
    # Save the new data to an Excel file
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df_new.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({'bg_color': '#bdbdbd', 'align': 'center', 'border': 1, 'font_name': 'Helvetica', 'font_size': 12})
        label_format = workbook.add_format({'bg_color': '#bdbdbd', 'align': 'center', 'border': 1, 'font_name': 'Helvetica', 'font_size': 12})
        default_format = workbook.add_format({'font_name': 'Helvetica', 'font_size': 12, 'border': 7})
        alt_row_format = workbook.add_format({'bg_color': '#f1f1f1', 'font_name': 'Helvetica', 'font_size': 12, 'border': 7})
        
        # Format the header row
        for idx, col in enumerate(df_new.columns):
            max_len = max(
                df_new[col].astype(str).map(len).max(),  # Length of largest item
                len(col)  # Length of column name/header
            ) + 5  # Adding extra space for larger text
            worksheet.set_column(idx, idx, max_len)
        worksheet.write_row(0, 0, df_new.columns, header_format)
        worksheet.freeze_panes(1, 0)  # Freeze the header row
        
        # Merge all columns for universe label rows and apply formatting
        for row_num in range(1, len(df_new) + 1):
            if str(df_new.iloc[row_num - 1, 0]).startswith('Universe') or str(df_new.iloc[row_num - 1, 0]).startswith('End of Universe'):
                worksheet.merge_range(row_num, 0, row_num, len(columns) - 1, df_new.iloc[row_num - 1, 0], label_format)
            else:
                # Apply formatting only to columns 1 through 10
                row_format = alt_row_format if row_num % 2 == 0 else default_format
                for col_num in range(len(columns)):
                    worksheet.write(row_num, col_num, df_new.iloc[row_num - 1, col_num], row_format)
    
    return output_path

# Main function to drive the program
def main():
    print("Welcome to The Fixture Sheet Generator")
    
    # Get fixture data from the user in a batch process
    fixture_data = patch_fixtures()
    
    # Generate the spreadsheet
    output_file = create_dmx_spreadsheet(fixture_data)
    
    print(f"Spreadsheet created successfully! The file is saved as: {output_file}")
    print("Please fill in the 'Unit # on position' column manually after opening the spreadsheet.")

# Run the program
if __name__ == "__main__":
    main()
