# Participant ID#: Change this and it will rename the folder and files with this ID#
participant_ID = "MRS_xx" # Whatever study fomat you use, just make sure it is the same as the grouping specification in the main processsing script.

# import OS module so python can interact with files 
import os

# Get the folder path where the script is located
script_folder_path = os.path.dirname(os.path.realpath(__file__))

# Define the path to the 'data' subfolder within the script's folder
folder_path = os.path.join(script_folder_path, "data")


# Create the subfolder (change ID# as needed)
subfolder_name = participant_ID
subfolder_path = os.path.join(folder_path, subfolder_name)


# Create the subfolder if it doesn't exist
os.makedirs(subfolder_path, exist_ok=True)

# List all files in the 'data' folder and skip any subfolders
files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

# Sort files by modification time (oldest first)
files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)))

# Create log file that will be saved in the new subfolder 
log_file = os.path.join(subfolder_path, "rename_log.txt")  # Log file path saved in the script's folder

# Open the log file for writing
with open(log_file, "w") as log:
    # Rename the files and move them to the subfolder
    for index, file in enumerate(files, start=1):
        old_path = os.path.join(folder_path, file)
        file_ext = os.path.splitext(file)[1]  # Keep original extension
        new_name = f"{participant_ID}_{index:02d}{file_ext}"  # Format as MRS_ID#_01, MRS_ID#_02, etc.
        new_path = os.path.join(subfolder_path, new_name)  # Save in the subfolder

        # Rename and move the file to the subfolder
        os.rename(old_path, new_path)

        # Write to the log file
        log.write(f"Renamed: {file} → {new_name}\n")

    log.write("Renaming complete!\n")

print("Renaming complete! Log saved to rename_log.txt")
