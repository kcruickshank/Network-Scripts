import subprocess

try:
    response = subprocess.check_output(
        ['ping', '-c', '3', '8.8.8.8'],
        stderr=subprocess.STDOUT,  # get all output
        universal_newlines=True  # return string not bytes
    )
    print(response)
except subprocess.CalledProcessError:
    response = None