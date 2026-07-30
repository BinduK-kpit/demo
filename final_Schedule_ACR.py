import os
import subprocess

# Folder containing your scripts
SCRIPT_DIR = r"C:\Users\S0QLTCQ\Downloads\Azure_WA_ACR_DR"

MAIN_SCRIPT = os.path.join(SCRIPT_DIR, "main_ACR.py")
HTML_PARSER_SCRIPT = os.path.join(SCRIPT_DIR, "html_ACR_parser_json.py")
TEST_ACR_SCRIPT = os.path.join(SCRIPT_DIR, "Scrape_APICall_ACR.py")


def run_script(script_path):
    print(f"Running {os.path.basename(script_path)}...")
    subprocess.run(
        ["python", script_path],
        check=True,
        cwd=SCRIPT_DIR
    )
    print(f"Completed: {os.path.basename(script_path)}")


def main():
    run_script(MAIN_SCRIPT)
    run_script(HTML_PARSER_SCRIPT)
    run_script(TEST_ACR_SCRIPT)

    print("\nAll scripts executed successfully.")
    print("Output files are stored locally in their respective output locations.")


if __name__ == "__main__":
    main()