import base64

with open("UserForm1.frx", "rb") as f:
    data = base64.b64encode(f.read())

with open("UserForm1.txt", "wb") as f:
    f.write(data)


import base64

with open("UserForm1.txt", "rb") as f:
    data = base64.b64decode(f.read())

with open("UserForm1.frx", "wb") as f:
    f.write(data)
