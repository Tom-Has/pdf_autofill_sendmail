import datetime
import hashlib

filename = "path_to_file/maildump.csv"
hashbase = "hashbase text"
header = "Name,Email,Country"
howmany = 100

with open(filename, "a") as file:
    file.write(header)

for x in range(howmany):
    tohash = hashbase + str(datetime.datetime.now()) + str(x)
    hashed = str(hashlib.md5(tohash.encode()).hexdigest())
    addy = "\n" + hashed + "," + hashed + "@mailinator.com" + ",Austria"
    with open(filename, "a") as file:
        file.write(addy)

