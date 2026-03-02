from uvbekutils import pyautobek


def reminder():
    # Identify which VoterLetters files should be downloaded before starting
    if True:
        choice = pyautobek.confirm(
                      (f"\nDownload data from Sincere before running.\n\n"
                       f"1. Get addresses assigned in Sincere:\n\n"
                       f"   All Enterprise >\n"
                       f"   ROV >\n"
                       f"   Reports >\n"
                       f"   New Report >\n"
                       f"   Parent Campaign Address COUNTS\n\n"
                       f"   If 'Locked' is selected, make sure reports are not including erroneous campaigns.\n\n\n"
                       f"2. Get Master and County campaign information from Sincere:\n\n"
                       f"   Reports >\n"
                       f"   New Report >\n"
                       f"   All Parent Campaign Address REQUESTS\n\n"
                       f"   Dates 1/1/24 to prior Monday INCLUSIVE (includes to the day prior to specified).\n\n\n"
                       f"3. Get Sincere users (to assign google read permissions):\n\n"
                       f"   Reports >\n"
                       f"   New Report >\n"
                       f"   All USERS"),
                        "REMINDER",
                       buttons=["Ok", "Exit"],
                          )
        if choice == "exit":
            exit()
