import subprocess
import time
import sys


def run_window_layout_manager(
    exe_path, window_title, ini_file, timeout=300
):  # 5 minutes timeout
    command = [exe_path, "Gen", window_title, ini_file]

    print(f"Executing command: {' '.join(command)}")
    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    start_time = time.time()
    output_lines = []
    while True:
        output = process.stdout.readline()
        if output == "" and process.poll() is not None:
            break
        if output:
            output_line = output.strip()
            print(output_line)
            output_lines.append(output_line)

        # Check for timeout
        if time.time() - start_time > timeout:
            print(f"Timeout after {timeout} seconds. Terminating process.")
            process.terminate()
            return False, output_lines

    # Process has finished, check return code
    rc = process.poll()
    if rc == 0:
        print("Process completed successfully.")
        return True, output_lines
    else:
        error = process.stderr.read()
        print(f"Error occurred: {error}")
        return False, output_lines


def main():
    exe_path = ".\\WindowLayoutManager.exe"
    window_title = r"""2024 - [C:\Users\xorjf\OneDrive\바탕 화면\'3764-(에스비일렉트릭)경북 예천군 용궁면 덕계리 380-1, 381 도성기1~5호태양광발전소(축사위)-完 (1)'\02-(마이다스)구조계산\태양광\'[태양광]도성기1~5] - [MIDAS/Gen]"""
    ini_file = ".\\midas_gen.ini"

    print("Starting Window Layout Manager...")
    print("Waiting for restoration to complete. This may take a few moments...")
    success, output_lines = run_window_layout_manager(exe_path, window_title, ini_file)

    if success:
        print("Window layout restoration process completed successfully.")
    else:
        print("Window layout restoration failed or timed out.")

    print("\nFull output:")
    for line in output_lines:
        print(line)


if __name__ == "__main__":
    # Set console output encoding to UTF-8
    sys.stdout.reconfigure(encoding="utf-8")
    main()
