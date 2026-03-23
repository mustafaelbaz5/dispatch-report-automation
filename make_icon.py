# make_icon.py
from PIL import Image
Image.open("car_logo.jpeg").save("car_logo.ico", format="ICO", sizes=[(256,256)])

