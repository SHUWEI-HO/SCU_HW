import cv2
import os
import argparse
import numpy as np
import matplotlib.pyplot as plt 

def get_parser():

    parser  = argparse.ArgumentParser()
    parser.add_argument('--input', default='images_data', help='Input Directory')
    parser.add_argument('--output', default='0314_outputs/', help='Output Directory')

    return parser.parse_args()

class Practice:
    def __init__(self, args):
        
        self.Input = args.input
        self.Output = args.output

    def main(self):

        images = [os.path.join(self.Input, img_name) for img_name in os.listdir(self.Input)]
        gamma = 0.125

        for idx, img_path in enumerate(images):

            img = cv2.imread(img_path)
            img = cv2.cvtColor(img,cv2.COLOR_BGR2RGB)

            # Calulate Images Histogram
            img_hist = cv2.calcHist([img], channels=[0], histSize=[256], range=[0, 256])

            # Image Blur
            img_blur_1 = cv2.pow(img, gamma)
            img_blur_2 = cv2.blur(img, ksize=(10,10))
            img_blur_2 = cv2.GaussianBlur(img, (5,5), 10)

            # Kernel Method
            kernels = np.ones(shape=(5, 5), dtype=np.float32) / 25
            filter = cv2.filter2D(img, -1, kernels)
            

if __name__ == '__main__':
    args = get_parser()
    Practice(args).main()