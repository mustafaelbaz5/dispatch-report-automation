# make_icon.py
from PIL import Image
Image.open("car_logo.png").save("car_logo.ico", format="ICO", sizes=[(256,256)])

