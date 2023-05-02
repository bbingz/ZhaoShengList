import subprocess

def recognize_text_from_image(image_path):
    try:
        result = subprocess.run(['SwiftOCR', image_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode == 0:
            return result.stdout
        else:
            print("Error:", result.stderr)
            return None
    except FileNotFoundError:
        print("Error: SwiftOCR not found")
        return None

image_path = '/Users/bing/Downloads/0.png'
recognized_text = recognize_text_from_image(image_path)
if recognized_text:
    print("Recognized text:", recognized_text)
