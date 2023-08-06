     -----------------------------------------
        Microsoft Windows Millennium Edition
          Help for Emergency Startup Disk  
     -----------------------------------------

     (c) Copyright Microsoft Corporation, 2000


This document provides complementary or late-breaking 
information to supplement the Windows Millennium Edition 
(Windows ME) documentation.

To close this Help file, press ALT-F-X.

-------- 
CONTENTS
--------

SUMMARY
  If Windows ME Starts Only in Safe Mode

TYPES OF WINDOWS ME INSTALLATION MEDIA AND METHODS
  Windows Millennium Edition (Retail Version)
  Windows Millennium Edition (Upgrade Version)  
  How Windows ME Upgrade Setup Searches for Previous Versions
  Clean Installations
  OEM Versions of Windows ME

IF WINDOWS ME DOES NOT START IN SAFE MODE
  Using ScanDisk to check your hard disk
  Using Windows Registry Checker

IF SETUP STOPS AND WINDOWS ME WILL NOT START IN SAFE MODE
  A Note on Anti-Virus Software
  Setup Stops Responding During Hardware Detection

UNINSTALLING WINDOWS ME

INSTALLING WINDOWS ME WITH WINDOWS NT OR WINDOWS 2000

INSTALLING WINDOWS ME ON A NEW HARD DISK


=======
SUMMARY
=======

If you have problems setting up or starting Windows ME, try 
starting your computer in Safe Mode before trying the steps 
outlined in this document.

If you can start in Safe Mode, see the Windows Millennium Edition 
Safe Mode Troubleshooter on the Safe Mode warning page. In Safe 
Mode, you can use Windows Troubleshooting tools, such as System 
Restore, which are not available from this Startup disk.

To start Windows in Safe Mode:

1. Remove the Startup disk, and then restart your computer. 

2. As soon as the computer restarts, press and hold down the 
   CTRL key until the Microsoft Windows Millennium Edition 
   Startup Menu appears.

3. On the Startup menu, choose Safe Mode (option 3).

4. Follow the steps in the Safe Mode Troubleshooter.

If Windows ME Starts Only in Safe Mode 
--------------------------------------
If you have followed the steps outlined in the Safe Mode 
Troubleshooter and are still unable to start Windows normally, 
do the following:
 
1. Close this Help file.

2. At the command prompt, type:
     scanreg /restore 
   and then press ENTER.

3. Select the most recent backup date, and then choose Restore.


If you try these steps and still cannot start normally:

1. Restart your computer by using the Windows ME Startup disk, 
   select option 1 on the Startup menu, and then press ENTER.

2. Close the Help file.

3. Try to reinstall Windows ME by following the instructions on 
   your screen.

**********
IMPORTANT: If you have installed software that came with your 
hard disk, be sure to read the documentation that describes how 
to start your computer by using a floppy disk.
**********

==================================================
TYPES OF WINDOWS ME INSTALLATION MEDIA AND METHODS
==================================================

The following section describes the versions of Windows ME and 
their installation requirements.

Windows Millennium Edition (Retail Version)
-------------------------------------------
This does not require a prior version of Microsoft Windows.

Windows Millennium Edition (Upgrade Version)
--------------------------------------------
Setup requires an existing copy of Microsoft Windows 95 or 
Windows 98.

How Windows ME Upgrade Setup Searches for Previous Versions
-----------------------------------------------------------
Setup searches your computer to find a qualifying product to 
upgrade. If Setup is unable to find a previous version of 
Windows on your computer, Setup prompts you to insert the 
previous version's media to confirm your eligibility for the 
upgrade. This media can be a CD-ROM, floppy disks, or a folder 
on the hard disk that contains setup files. If you supply 
floppy disks, Setup may prompt you for multiple disks.

Clean Installations
-------------------
You can use the Windows ME Upgrade to install Windows ME on a 
computer that does not have a prior version of Windows installed, 
such as after formatting your hard disk. This is commonly called 
a "clean install."

Because the compliance check process is unable to find evidence 
of a prior version, Setup prompts you to insert your previous 
version media to confirm your eligibility for the upgrade. This 
media can be a CD-ROM, floppy disks, or a folder on the hard disk 
that contains setup files. If you supply floppy disks, Setup may 
prompt you for multiple disks.

For more information about how to install Windows ME on a 
computer that has no previous operating system, see "How To 
Install Windows Millennium Edition Onto a New Hard Disk" 
referenced at the end of this document.

OEM versions of Windows ME
--------------------------
An Original Equipment Manufacturer (OEM) version of Windows ME 
is a special release product made available for computer 
Manufacturers. OEMs can customize these versions of Windows ME 
specifically for their hardware and software.

The Setup procedures and requirements outlined in this document 
may be different if you have an OEM version of Windows ME. For 
more information, read the documentation that came with your 
computer or contact your computer manufacturer.


=========================================
IF WINDOWS ME DOES NOT START IN SAFE MODE
=========================================

Any of the following conditions can cause Windows ME not to start 
in Safe Mode:

- Your computer is infected with a virus. Run up-to-date anti-
  virus software to check for a virus and clean your computer if 
  necessary.   

- Your computer's CMOS settings are not correct. Check your 
  computer's CMOS settings to make sure they are correct. Note 
  that you may need to contact the computer manufacturer to 
  verify these settings.

- There is a hardware failure. Note that you may need to contact 
  the computer manufacturer for more information about your 
  hardware.      

- There is an error on your computer's hard disk.  See "Using 
  ScanDisk to check your hard disk."

- There is an error in the Windows registry. See "Using the 
  Windows Registry Checker."

Using ScanDisk to Check Your Hard Disk
--------------------------------------
If you suspect there may be file corruption or other problems 
with your hard disk(s), run ScanDisk to check for and repair 
errors.

To check all your hard disks for errors:

1. At the command prompt, type:
   
   scandisk /all 

2. Press ENTER.

To perform a full surface scan of your hard disk(s) for maximum 
protection against data loss:

1. At the command prompt, type:

   scandisk /all /surface
   
2. Press ENTER.

Using Windows Registry Checker
------------------------------
If you are still unable to start Windows ME in Safe Mode, run 
the Windows Registry Checker (Scanreg.exe) tool, as there may 
be a problem with the system registry. 

To start the Windows Registry Checker:

1. At the command prompt, type:
   
   scanreg /restore 

2. Press ENTER.  

Scanreg may not be available if Windows ME has not been 
successfully installed on your computer.


=========================================================
IF SETUP STOPS AND WINDOWS ME WILL NOT START IN SAFE MODE
=========================================================

The following section explains what you can do to recover from 
a failed Windows Setup. For more information about other Setup 
problems, see the Setup.txt file in the Win9X folder of your 
Windows ME CD.

If you encounter any of these error messages while running Setup:

- Invalid system disk 
 
- Incorrect MS-DOS version 

- Missing or corrupted Command.com 

it is likely that your computer's startup drive needs updated 
system files. Certain CMOS settings or anti-virus software can 
prevent Windows Setup from installing the correct system files 
on your computer.

To replace your system files:

1. Restart your computer by using the Windows ME Startup Disk, 
   selecting option 1 on the Startup menu, and then pressing 
   ENTER.

2. Close the Help file.

3. Follow the instructions on the screen to run Setup.

4. Choose "Use Safe Recovery" if prompted.

A Note on Anti-Virus Software
-----------------------------
If anti-virus programs are left running during Setup, they might 
prevent Setup from properly updating the system files. 

**********
IMPORTANT: You might receive a warning message after the first 
restart during Setup, informing you that the Master Boot Record 
or other files have changed. If you see such a message, you MUST 
accept these changes or Setup may fail to update critical files 
that Windows ME uses to start your computer. 
**********

Setup Stops Responding During Hardware Detection
------------------------------------------------
If Setup stops responding while it is detecting the hardware in 
your computer, turn your computer off and wait a few seconds, and 
then turn it back on. You may need to do this several times. 
Choose Use "Safe Recovery" when Setup restarts. Setup could stop 
responding during several different detection modules and will 
skip areas it fails to complete successfully.

**********
IMPORTANT: Use the power switch to turn your computer completely 
off. Do not use the reset button or press CTRL+ALT+DEL to restart 
your computer.
**********

If Setup still fails to complete successfully, it may be 
necessary to start your computer in Safe Mode so that you can 
view the Help topics associated with hardware detection.


=======================
UNINSTALLING WINDOWS ME
=======================

If the above steps do not work, you can try to uninstall 
Windows ME and return to your previous version of Windows.

To uninstall Windows ME, you must have chosen to save uninstall 
information during Windows ME Setup. The uninstall information 
is saved in the Winundo.dat and Winundo.ini files. If these 
files are deleted, you cannot uninstall Windows ME.

**********
IMPORTANT: You should not attempt to Uninstall Windows ME if 
the partition information for your hard disk or disks has 
changed since you last installed Windows ME successfully. 
Windows ME creates a backup copy of your partition information 
in the Suhdlog.dat file at the end of a successful installation, 
and Uninstal.exe restores the partition information listed in 
the Suhdlog.dat file to your hard disk during Uninstall. If the 
partition information on your hard disk or disks has changed 
since you last installed successfully (specifically, since the 
Suhdlog.dat file was created), then you might experience partial 
or complete data loss as a result of attempting to Uninstall 
Windows ME.
**********

To uninstall Windows ME:

1. Restart your computer using the Windows ME Startup Disk, 
   selecting option 1 on the Startup menu, and then pressing 
   ENTER.

2. Close the Help file.

3. At the command prompt, type each of the following lines, 
   pressing ENTER after each line.

   smartdrv
   <drive>: 
   cd\<windows>\command
   uninstal.exe

   NOTE: <drive> is the drive letter of the drive containing the 
   Windows folder (the default is C). <windows> is the name of 
   the Windows folder (the default is "Windows").

4. Type Y when you are prompted to continue. 
   This process can take several minutes.

NOTE: SmartDrive (Smartdrv.exe) is not required to run 
Uninstal.exe, but it can significantly speed up the process.


=====================================================
INSTALLING WINDOWS ME WITH WINDOWS NT OR WINDOWS 2000
=====================================================

You cannot install Windows ME over any version of Windows 2000 
or Windows NT, but they can exist together on a single system. 
However, for compatibility reasons, it is recommended that 
you install each to a separate hard disk or partition. If 
Windows NT is already installed, Windows ME Setup will add 
itself to the Windows NT boot menu to allow you to multi-boot 
between Windows ME and Windows NT.


========================================
INSTALLING WINDOWS ME ON A NEW HARD DISK
========================================

See the document "Installing Windows Millennium Edition to a 
New Hard Disk," located at <CD-ROM>:\Win9x\Cleanhd.Txt.

You might also find the Setup readme useful. It is located at 
<CD-ROM>:\Win9X\Setup.txt.
