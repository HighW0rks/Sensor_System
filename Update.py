import os
import requests
import psutil
import zipfile
import sys
import shutil
from Config import text_config, readfile_value

# Probeer een tweede applicatie te maken die activeert wanneer een nieuwe versie beschikbaar is.
# Het zal de huidige applicatie verwijderen en de nieuwe applicatie neerzetten
# Als het om de config gaat zal hij de data van de version opslaan en writen wanneer er een nieuwe geplaatst is
# Doe dit niet via index maar in een "for config in read:" etc.



def update():
    terminate_existing_main_processes()
    response = requests.get("https://api.github.com/repos/HighW0rks/Sensor_System/releases/latest")
    if response.status_code == 200:
        latest_release = response.json()
    else:
        print("Failed to retrieve the latest release.")
        return
    if latest_release:
        response = requests.get("https://api.github.com/repos/HighW0rks/Sensor_System/tags")
        print(latest_release.get('body'))
        if response.status_code == 200:
            latest_tag = response.json()[0]['name']
            print(f"Latest tag version: {latest_tag}")
        else:
            print("Failed to retrieve the latest tag.")
            exit()

        if readfile_value(12) != latest_tag:
            for asset in latest_release.get('assets', []):
                if asset['name'] == 'Skalar_Saxon_Tester.zip':
                    asset_url = asset['browser_download_url']
                    break

            if asset_url:
                # Download asset
                response = requests.get(asset_url)
                if response.status_code == 200:
                    execute_path = os.path.abspath(sys.argv[0])
                    main_folder = os.path.dirname(execute_path)
                    extract_folder = fr"C:\Users\{os.getlogin()}\Extract"
                    os.makedirs(extract_folder, exist_ok=True)
                    path = extract_folder + r"\Skalar_Saxon_Tester.zip"
                    with open(path, 'wb') as f:
                        f.write(response.content)
                    print(f"Asset downloaded to: {path}")

                    # Unzip file
                    with zipfile.ZipFile(path, 'r') as zip_ref:
                        # Extract only the files mentioned in the body text
                        body_text = latest_release.get('body', '')
                        files_changed = parse_files_changed(body_text)
                        for file in files_changed:
                            zip_ref.extract(file, main_folder)
                    print(f"Files extracted to: {main_folder}")

                    # Remove zip file
                    shutil.rmtree(extract_folder)
                    print("Zip file deleted.")
                    text_config(12, latest_tag)
                else:
                    print(f"Failed to download asset. Status code: {response.status_code}")
            else:
                print(f"Asset '{asset_name}' not found in the latest release.")

        else:
            print("No update found")

def parse_files_changed(body_text):
    files_changed = []
    # Split the body text by lines
    lines = body_text.split('\n')
    # Find the line that starts with "Files changed:"
    start_index = -1
    for i, line in enumerate(lines):
        if line.strip().startswith("Files changed:"):
            start_index = i
            break
    if start_index != -1:
        # Add all lines after the "Files changed:" line until a blank line is encountered
        for line in lines[start_index + 1:]:
            if line.strip():
                files_changed.append(line.strip().lstrip('- '))
            else:
                break
    return files_changed


def terminate_existing_main_processes():
    # Function to terminate existing instances of the main application process
    for proc in psutil.process_iter(['pid', 'name']):
        # Iterate through all running processes
        if proc.info['name'] == 'Skalar Saxon Tester.exe':
            # Check if the process name matches the main application
            proc.terminate()  # Terminate the process
        if proc.info['name'] == 'Skalar Saxon Tester Console.exe':
            proc.terminate()

# Call the update function
update()
