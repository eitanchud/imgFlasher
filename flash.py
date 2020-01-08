#!/usr/bin/env python3

import wmi as wmi
import time


import sys
import os
import subprocess
import win32file

CreateFile = win32file.CreateFile
CloseHandle = win32file.CloseHandle
DeviceIoControl = win32file.DeviceIoControl
ReadFile = win32file.ReadFile
WriteFile = win32file.WriteFile
w32GenericRead = win32file.GENERIC_READ
w32OpenExisting = win32file.OPEN_EXISTING


def lockWinDevice(physicalDevice, volumeGUID):

    hDevice = ""

    hVolumes = []
    for vguid in volumeGUID:
        # vguid need the last \ removed
        vguidc = vguid[:-1]
        try :
            # Open the volume
            hVolume = CreateFile(vguidc, 1|2, 1|2, None, 3, 0, None)
            # Lock the volume
            out = DeviceIoControl(hVolume, 589848, None, None, None)
            # Dismount the volume
            out = DeviceIoControl(hVolume, 589856, None, None, None)
            # append the hvolume to the array
            hVolumes.append(hVolume)
        except:
            # error
            print("Error Opening, locking or dismounting {}".format(vguidc))
            raise

    try:

        hDevice = CreateFile(physicalDevice, 1|2, 1|2, None, 3, 0, None)
        DeviceIoControl(hDevice, 589848, None, None, None)
        # Dismount the device
        DeviceIoControl(hDevice, 589856, None, None, None)
    except:
        # error
        print("Error Opening, locking or dismounting {}".format(physicalDevice))

    if hDevice and hVolume:
        return hDevice, hVolumes
    else:
        return False, False

def Devices():

    c = wmi.WMI()
    data = []

    for physical_disk in c.Win32_DiskDrive():
        if 7 in physical_disk.Capabilities:
            # is a removable media
            phy = physical_disk.DeviceID
            size = int(physical_disk.Size)
            desc = physical_disk.Caption
            drives = []
            for partition in physical_disk.associators ("Win32_DiskDriveToDiskPartition"):
                for logical_disk in partition.associators ("Win32_LogicalDiskToPartition"):
                    name = logical_disk.Name + "\\"
                    label = logical_disk.VolumeName
                    guid = win32file.GetVolumeNameForVolumeMountPoint(name)
                    drives.append((name, label, guid))

            data.append([phy, desc, size, drives])

    return data

def getphyguid(letter, phydata):

    if not "\\" in letter:
        letter = letter + "\\"

    for phyd in phydata:
        phy = phyd[0]
        size = int(phyd[2])
        guids = []
        isit = False
        for d in phyd[3]:
            guids.append(d[2])
            if d[0].lower() == letter.lower():
                isit = True
        
        if isit:
            return (phy, guids, size)

    return (False, False, False)

def myFmtCallback(command, modifier, arg):
    print(command)
    return 1    # TRUE

def start(image, drive, logfile):
    # main action goes here
    image = image
    fsize = 0
    logfile = logfile


    data = Devices()

    try:
        fsize = os.path.getsize(image)
    except FileNotFoundError:
        print("ERROR: File '{}' not found, please check that.".format(image))
        sys.exit()

    (physicalDevice, volumeGUID, dsize) = getphyguid(drive, data)
    if physicalDevice == False:
        print("ERROR: Can't find the specified drive '{}', please check that.".format(drive))
        sys.exit()

    #format disk
    sd_format_script_path = "C:\\Users\\Public\\sdformat.txt"
    if not os.path.exists(os.path.dirname(sd_format_script_path)):
        sd_format_script_path = "C:\\sdformat.txt"

    file = open(sd_format_script_path, "w")
    sd_format_cmd = "REM ****** CLEAN SDCARD AND REFORMAT ******\nselect disk {}\nclean\ncreate partition primary\nselect partition 1\nactive\nformat fs=fat32 quick\nassign letter {}\nonline disk\nexit".format(physicalDevice[-1], drive[0])
    file.write(sd_format_cmd)
    file.close()

    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    p = subprocess.call("diskpart /s {} > nul".format(sd_format_script_path), startupinfo=si)

    time.sleep(5)
    # open the logfile
    lf = open(logfile, mode='wt', buffering=1)

    # check sizes mismatch
    if fsize > dsize:
        print("ERROR: File '{}' is bigger than drive {}, aborting.".format(image, drive))
        sys.exit()

    # Receive the paths to the physical device device and logical volume and return a handler to each one
    hDevice, hVolumes = lockWinDevice(physicalDevice, volumeGUID)

    # Open the Skybian file using a PyHANDLE (a wrapper to a standard Win32 HANDLE)
    inputFileHandle = CreateFile(image, w32GenericRead, 0, None, w32OpenExisting, 0, None)

    # windows needs a write chunk to be multiple of the sector size, aka 512 bytes
    sectorSize = 512
    portionSize = sectorSize * 1000 # 1/2 MB chunks
    actualPosition = 0

    # build node loop
    while actualPosition < fsize:
        errorCode, data = ReadFile(inputFileHandle, portionSize)

        # returned error code
        if errorCode != 0:
            print("ERROR: Read from image file failed!")
            sys.exit()

        # avoid writes with no sectorsize length
        if len(data) != portionSize:
            # final part, fill with zeroes until 512 multiple
            diff = len(data) % sectorSize
            data += bytearray(sectorSize - diff)

        # actual write
        ws = WriteFile
        errorCode, ws = WriteFile(hDevice, data)


        if errorCode != 0:
            print("ERROR: Write to device failed!")
            sys.exit()

        if ws != len(data):
            print("ERROR: Write data != from read data!")
            sys.exit()

        actualPosition += portionSize
        per = actualPosition / fsize
        if per > 100:
            per = 100.0

        lf.write("{:.1%}\n".format(per))

    os.remove(sd_format_script_path)
    CloseHandle(inputFileHandle)
    CloseHandle(hDevice)
