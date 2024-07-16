import subprocess
import os


def get_commits():
    result = subprocess.run(["git", "rev-list", "HEAD"], stdout=subprocess.PIPE)
    commits = result.stdout.decode("utf-8").split()
    return commits


def checkout_commit(commit):
    subprocess.run(["git", "checkout", commit])


def run_script():
    script_path = os.path.join(os.getcwd(), "AutoMouseTracker.py")
    if not os.path.isfile(script_path):
        print(f"Errors: {script_path} does not exist")
        return

    result = subprocess.run(
        ["python", script_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    print(result.stdout.decode("utf-8", errors="ignore"))
    if result.stderr:
        print("Errors:", result.stderr.decode("utf-8", errors="ignore"))


def get_current_commit():
    result = subprocess.run(
        ["git", "rev-parse", "--abbrev-ref", "HEAD"], stdout=subprocess.PIPE
    )
    current_commit = result.stdout.decode("utf-8").strip()
    return current_commit


def main():
    current_commit = get_current_commit()
    commits = get_commits()
    for commit in commits:
        print(f"Checking out commit: {commit}")
        checkout_commit(commit)
        run_script()

    print(f"Restoring the original commit: {current_commit}")
    checkout_commit(current_commit)


if __name__ == "__main__":
    main()
