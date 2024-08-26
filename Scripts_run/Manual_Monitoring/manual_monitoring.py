import subprocess
import sys
import datetime
import time
import os

current_directory = os.getcwd()

def run_scripts(script_list):
    try:
        for script in script_list:
            try:
                print(f"Running {script} script...")
                script_path = os.path.abspath(os.path.join(current_directory, 'Files', 'Manual_Monitoring',f'{script}_script.py'))
                subprocess.run(["python", script_path], check=True)
            except subprocess.CalledProcessError:
                print(f"{script.upper()} Error")
                sys.exit()

        print("All selected scripts executed successfully.")

    except subprocess.CalledProcessError as e:
        print(f"Error running script: {e}")

def main():
    start_time = time.time()
    option = int(input("Enter your choice (1 for monitoring, 2 for vertex): ").strip())

    if option == 1 :
        print("Available scripts:")
        print("1. PR3")
        print("2. PK9")
        print("3. PMP")
        print("4. PL5")
        print("5. PKX")
        print("6. PKS")
        print("7. CP5")
        print("8. PM6")
        print("9. PR5")
        print("10. Document Creation")

        selected_scripts = []
        while True:
            script_choices = input("Enter the script numbers to run (seperated by space), or 'all' to run all, or leave blank to finish: ").strip().lower()
            
            if script_choices == "":
                break
            elif script_choices == "all":
                selected_scripts = [
                    "pr3", "pk9", "pmp", "pl5", "pkx", "pks", "cp5", "pm6", "pr5", "doc_creation"
                ]
                break
            else:
                choices = script_choices.split()
                invalid_choices = False
                for choice in choices:
                    if choice.isdigit() and int(choice) in range(1, 11):
                        script_name = {
                            "1": "pr3", "2": "pk9", "3": "pmp", "4": "pl5", "5": "pkx", 
                            "6": "pks", "7": "cp5", "8": "pm6", "9": "pr5", "10": "doc_creation"
                        }[choice]
                        selected_scripts.append(script_name)
                    else:
                        invalid_choices = True
                        break
                
                if invalid_choices:
                    print("Invalid choice. Please enter valid script numbers (1 to 10) or 'all'.")
                    continue

        if selected_scripts:
            run_scripts(selected_scripts)
        else:
            print("No scripts selected.")

    elif option == 2:
        print("Running Vertex script...")

        vertex_script_path = os.path.abspath(os.path.join(current_directory, 'Files', 'Vertex','vertex_full.py'))
        subprocess.run(["python", vertex_script_path], check=True)
        print("Vertex script executed successfully.")

    else:
        print("Invalid option. Please choose either 'monitoring' or 'vertex'.")
    
    end_time = time.time()
    running_time = end_time - start_time
    running_time = str(datetime.timedelta(seconds = int(end_time - start_time)))
 
    print (f"Total running time (hh:mm:ss format) : {running_time}")

if __name__ == "__main__":
    main()
