
import requests
import json
import jsonpath

import os

from pathlib import Path
home = str(Path.home())
print(home)


# Get the current working
# directory (CWD)
cwd = os.getcwd()

# Print the current working
# directory (CWD)
print("Current working directory:"+cwd)
