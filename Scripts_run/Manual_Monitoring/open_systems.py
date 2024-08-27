import datetime
import time
import subprocess
import sys
import os

current_directory = os.getcwd()

def main():
    start_time = time.time()
    script_path = os.path.abspath(os.path.join(current_directory, 'Files', 'Manual_Monitoring','open_systems.py'))
    subprocess.run(["python", script_path], check=True)
    end_time = time.time()
    running_time = end_time - start_time
    running_time = str(datetime.timedelta(seconds = int(end_time - start_time)))
 
    print (f"Total running time (hh:mm:ss format) : {running_time}")

if __name__ == "__main__":
    main()