import cv2
import numpy as np

def analyze_image(image_path):
    image = cv2.imread(image_path)
    if image is None:
        raise ValueError("Image not found or unable to load")

    image = cv2.resize(image, (300, 300), interpolation=cv2.INTER_AREA)
    hsv_image = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    hue, saturation, value = cv2.split(hsv_image)
    saturation = saturation / 255.0
    value = value / 255.0

    warm_hue_mask = ((hue >= 0) & (hue <= 30)) | ((hue >= 150) & (hue <= 180))
    warm_hue_percentage = np.sum(warm_hue_mask) / (hue.shape[0] * hue.shape[1])

    avg_saturation = np.mean(saturation)
    avg_brightness = np.mean(value)
    contrast_brightness = np.std(value) / np.max(value)

    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    laplacian_var = cv2.Laplacian(gray_image, cv2.CV_64F).var()
    clarity_score = min(laplacian_var / 1000.0, 1.0) 

    results = {
        "warm_hue": round(warm_hue_percentage, 3),
        "saturation": round(avg_saturation, 3),
        "brightness": round(avg_brightness, 3),
        "contrast_brightness": round(contrast_brightness, 3),
        "image_clarity": round(clarity_score, 3),
    }

    return results

if __name__ == "__main__":
    image_path = "Los Angeles\\90002\\759 E 105th St\\genMid.CV24206046_3_0.jpg"
    x ='Los Angeles\\90002\\759 E 105th St\\genMid.CV24206046_2_0.jpg'
    metrics = analyze_image(image_path)
    y = analyze_image(x)
    print(metrics)
    print(y,)
