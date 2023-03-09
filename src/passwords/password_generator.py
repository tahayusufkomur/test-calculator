import random
import string
import json


class PasswordGenerator:
    def __init__(self,
                 prefix,
                 count):
        self.prefix = prefix
        self.count = int(count)
        self.passwords = {}
        self.create_password()
        self.save_passwords()

    def create_password(self):

        all = string.ascii_lowercase + string.ascii_uppercase + string.digits

        for i in range(1, self.count+1):
            temp = random.sample(all, 4)
            password = "".join(temp)
            self.passwords[f"{self.prefix}_{password}"] = f'yedek{i}'
        regex = "|".join(self.passwords.keys())

        self.passwords['regex'] = regex

    def save_passwords(self):
        with open(f"{self.prefix}.txt", "w") as outfile:
            json.dump(self.passwords, outfile, indent=4)
