import io
import numpy as np
import cv2
from PIL import Image, ImageFile


# Erhöhen Sie das Limit für die Bildgröße
# Image.MAX_IMAGE_PIXELS = None
# ImageFile.LOAD_TRUNCATED_IMAGES = True


def check_orientation(img_url):
    with Image.open(img_url) as img:
        width, height = img.size
        if width > height:
            return "Landscape"
        else:
            return "Portrait"


def crop_image_percent(img_path, top, bottom, left, right):
    img = cv2.imread(img_path)
    if img is None:
        raise ValueError(f"Bild konnte nicht geladen werden: {img_path}")
    top_crop = int(img.shape[0] * top)
    bottom_crop = int(img.shape[0] * bottom)
    left_crop = int(img.shape[1] * left)
    right_crop = int(img.shape[1] * right)
    cropped_img = img[top_crop:bottom_crop, left_crop:right_crop]
    if cropped_img.size == 0:
        raise ValueError("Das zugeschnittene Bild ist leer. Überprüfe die Zuschneideparameter.")
    print(cropped_img)
    print(type(cropped_img))
    return cropped_img


def crop_image(img_path):
    img = cv2.imread(img_path)
    height, width, _ = img.shape
    cropped_img = img[height // 4: height - (height // 4), 0 : width]
    # cv2.imshow("cropped image", cropped_img)
    # cv2.waitKey(0)
    return cropped_img
