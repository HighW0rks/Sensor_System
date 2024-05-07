import os
import requests
import psutil
import zipfile
import sys
import shutil
import threading
import logging
from Config import text_config, readfile_value


def log():
    while True:
        logging.basicConfig(filename='log.txt', level=logging.NOTSET)
        sys.stdout = sys.stderr = open('log.txt', 'a')


def update():
    terminate_existing_main_processes()

    # Replace 'YOUR_GITHUB_TOKEN' with your actual GitHub personal access token
    headers = {
        'Authorization': 'token ',
        'Accept': 'application/vnd.github.v3+json'
    }

    response = requests.get("https://api.github.com/repos/HighW0rks/Sensor_System/releases/latest", headers=headers)
    if response.status_code == 200:
        latest_release = response.json()
    else:
        print("Failed to retrieve the latest release.")
        return

    if latest_release:
        response = requests.get("https://api.github.com/repos/HighW0rks/Sensor_System/tags", headers=headers)
        if response.status_code == 200:
            latest_tag = response.json()[0]['name']
        else:
            print("Failed to retrieve the latest tag.")
            exit()

        if readfile_value(12) != latest_tag:
            for asset in latest_release.get('assets', []):
                if asset['name'] == 'Skalar_Saxon_Tester.zip':
                    asset_url = asset['browser_download_url']
                    break

            if asset_url:
                response = requests.get(asset_url, headers=headers)
                if response.status_code == 200:
                    execute_path = os.path.abspath(sys.argv[0])
                    main_folder = os.path.dirname(execute_path)
                    print("The main folder locations is: ", main_folder)
                    extract_folder = fr"C:\Users\{os.getlogin()}\Extract"
                    os.makedirs(extract_folder, exist_ok=True)
                    path = extract_folder + r"\Skalar_Saxon_Tester.zip"
                    with open(path, 'wb') as f:
                        f.write(response.content)

                    print(f"Asset downloaded to: {path}")

                    with zipfile.ZipFile(path, 'r') as zip_ref:
                        body_text = latest_release.get('body', '')
                        files_to_move = parse_files_changed(body_text)

                        # Extract all files - maintains folder structure
                        zip_ref.extractall(extract_folder)

                        # Move only files/folders in files_to_move
                        for root, _, files in os.walk(extract_folder):
                            for file in files:
                                if file in files_to_move:
                                    source = os.path.join(root, file)
                                    dest = os.path.join(main_folder, file)  # Maintain relative path
                                    shutil.move(source, dest)
                            for folder in os.listdir(root):
                                if folder in files_to_move:
                                    source = os.path.join(root, folder)
                                    dest = main_folder  # Maintain relative path
                                    if os.path.exists(dest):
                                        shutil.rmtree(fr"{dest}\{folder}")
                                    shutil.move(source, dest)

                    print(f"Extracted files: All Files")  # All files extracted
                    print(f"Moved files/folders: {', '.join(files_to_move)}")
                    print(f"Files located at: {main_folder}")

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
    lines = body_text.split('\n')
    start_index = -1
    for i, line in enumerate(lines):
        if line.strip().startswith("Files changed:"):
            start_index = i
            break
    if start_index != -1:
        for line in lines[start_index + 1:]:
            if line.strip():
                files_changed.append(line.strip().lstrip('- '))
            else:
                break
    return files_changed


def terminate_existing_main_processes():
    # Function to terminate existing instances of the main application process
    current_pid = os.getpid()  # Get the PID of the current process
    for proc in psutil.process_iter(['pid', 'name']):
        # Iterate through all running processes
        if proc.info['pid'] == current_pid:
            # Check if the process PID matches the PID of the current process
            os.system("Skalar Saxon Tester.exe")
            proc.terminate()  # Terminate the process


threading.Thread(target=log, daemon=True).start()
update()
