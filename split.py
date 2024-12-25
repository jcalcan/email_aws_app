import csv
import os
from tkinter import filedialog, Tk


def split_csv(input_file, output_prefix, rows_per_file=500):
    with open(input_file, 'r', newline='') as csvfile:
        reader = csv.reader(csvfile)
        header = next(reader)
        file_number = 1
        row_count = 0
        current_output = None
        writer = None

        for row in reader:
            # Check if we need to create a new file
            if row_count % rows_per_file == 0:
                if current_output:
                    current_output.close()

                # Create a new output file name
                output_file = f"{output_prefix}_{file_number:03d}.csv"
                current_output = open(output_file, 'w', newline='')
                writer = csv.writer(current_output)
                writer.writerow(header)  # Write header to each new file
                print(f"Creating file: {output_file}")  # Debug print statement
                file_number += 1  # Increment file number after creating a new file

            # Write the current row to the current output file
            writer.writerow(row)
            row_count += 1

        # Close the last output file if it exists
        if current_output:
            current_output.close()

    print(f"Split complete. Created {file_number - 1} files.")


# Create a root window and hide it
root = Tk()
root.withdraw()

# Open file dialog for user to select input CSV file
input_file = filedialog.askopenfilename(title="Select CSV file to split", filetypes=[("CSV files", "*.csv")])

if input_file:
    # Get the directory and base name of the input file
    input_dir = os.path.dirname(input_file)
    input_base = os.path.splitext(os.path.basename(input_file))[0]

    # Create output prefix
    output_prefix = os.path.join(input_dir, f"{input_base}_split")

    # Call the split_csv function
    split_csv(input_file, output_prefix)
else:
    print("No file selected. Exiting.")