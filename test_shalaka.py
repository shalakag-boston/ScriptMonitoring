"""
Dummy test script for ScriptMonitoringEnd.py

Run this locally (on Windows, with your .env filled in and
python-dotenv installed) to confirm the env-var-based credentials
work end to end.

Usage:
    python test_shalaka.py
"""

from ScriptMonitoringEnd import scriptmonitoring
import datetime

now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def run_case(failure_flag):
    """
    failure_flag == 0 -> treated as a Success run
    failure_flag == 1 -> treated as a Failure run
    Keeping result tied directly to failure_flag avoids the two
    ever getting out of sync (e.g. failure=1 but result="Success").
    """
    result = "Success" if failure_flag == 0 else "Failure"
    label = "SUCCESS" if failure_flag == 0 else "FAILURE"

    print(f"=== Running {label} case (failure={failure_flag}) ===")
    scriptmonitoring(
        script="shalaka_test",
        scriptstart=now,
        scriptfinish=now,
        scriptinput="test input",
        scriptoutput=f"test ran as a {label.lower()} case",
        result=result,
        failure=failure_flag,
    )
    print(f"{label} case done.")


run_case(0)  # failure == 0 -> Success
print("Check SQL table / CSV backup for a row with script='shalaka_test', result='Success'.")

run_case(1)  # failure == 1 -> Failure
print("Check for an Asana task assigned to ipoole@bu.edu, or a fallback email.")
