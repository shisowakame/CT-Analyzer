import os
import glob
import numpy as np
import pydicom as dicom
import matplotlib.pyplot as plt
#import itertools

fn = r"C:\Users\idy01\HirosakiUniv-fixed-test\testB\RT-CT_48_fixed\p48_81_out.dcm" #faile name of image or path to the image

#(y,x)
ROI = (230,180)
S = (10,10) # 10x10 ROI

img = dicom.dcmread(fn, force=True)
img.file_meta.TransferSyntaxUID = dicom.uid.ImplicitVRLittleEndian
dat = img.pixel_array.astype(np.float32) + img.RescaleIntercept # img.pixel._array is "numpy array" with the shape (512,512)
#dir_name = os.path.dirname(fn)
#print(dir_name)

rdat = dat[ROI[0]:(ROI[0]+S[0]), ROI[1]:(ROI[1]+S[1])]
rdat = rdat.flatten()

# mean and std in the ROI
hu_man = np.mean(rdat)
hu_std = np.std(rdat)

print(f"Mean HU in ROI: {hu_man:.2f}")
print(f"Standard Deviation of HU in ROI: {hu_std:.2f}")
