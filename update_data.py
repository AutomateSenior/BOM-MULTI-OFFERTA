#!/usr/bin/env python3
"""
Script to update BOM8 library version in Google Apps Script manifest
Uses clasp to push only the appsscript.json file
"""

import subprocess
import sys
import os

def main():
    # Change to BOM directory
    bom_dir = "/Users/enricoerba/openprojects/automate/gestione interna/BOM"
    os.chdir(bom_dir)

    print("Updating appsscript.json to use BOM8 library version 16...")

    try:
        # Pull latest to sync
        print("Pulling latest files...")
        subprocess.run(["clasp", "pull"], check=True)

        # Now push with force to overwrite
        print("Pushing updated manifest...")
        result = subprocess.run(
            ["clasp", "push", "--force"],
            capture_output=True,
            text=True
        )

        if result.returncode == 0:
            print("✓ Successfully updated to BOM8 library version 16")
            print(result.stdout)
        else:
            print("✗ Push failed:", result.stderr)
            # Try alternative: push specific file only
            print("\nTrying alternative method...")
            subprocess.run(["clasp", "push", "--force", "appsscript.json"], check=True)
            print("✓ Successfully pushed appsscript.json")

    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
