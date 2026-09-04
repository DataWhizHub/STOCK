"""
Run locally to create the password_hash values needed in secrets.toml.

    python generate_password_hash.py

Then paste the printed hash into .streamlit/secrets.toml.
"""

import getpass
import bcrypt

if __name__ == "__main__":
    pwd = getpass.getpass("Enter the password to hash: ")
    hashed = bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    print("\npassword_hash =", f'"{hashed}"')
