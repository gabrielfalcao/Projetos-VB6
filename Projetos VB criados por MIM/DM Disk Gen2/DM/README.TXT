		    Welcome to DiscWizard!

DiscWizard is a Windows-based installation utility that makes
adding a new hard drive to your system easy.  To use DiscWizard,
follow these steps BEFORE attaching your new hard drive:

   1) Make a working copy of the DiscWizard diskette.
      Use the working copy to install your hard drive.
      Keep the original diskette in a safe place.  

   2) Boot your machine and start Windows.

   3) Insert the DiscWizard installation diskette into drive A:.

   4) For Windows 95:
	Select "Run..." from the Start menu.  Type a:\setup

      For Windows 3.1:
	Select "Run..." from the Program Manager File menu.  
	Type a:\setup 

   5) Setup will copy DiscWizard files to your hard drive.
      DiscWizard will then:
	-ask you how you wish to place your new drive in your 
	 system,
	-make custom installation instructions for jumpering and 
	 attaching your drive, and
	-prepare your new hard drive for use.

To install your hard disk using Disk Manager, simply follow these 
basic steps:

   1) Following the instructions that DiscWizard provides you,
      set your system Setup (CMOS), jumper your drive(s), and 
      physically install your hard disk into your computer.
      Disk Manager assumes power, cables, jumpers, etc.,
      are connected and properly installed.  

   2) Insert your DOS diskette in drive A: and boot the machine
      to DOS. 

   3) Insert your DiscWizard diskette into drive A: and
      type DM to invoke Disk Manager.

   4) When the installation has completed, remove the DiscWizard
      diskette from drive A: and boot from the hard disk.  

See the Disk Manager online manual for more information.

DiscWizard (TM) is a trademark of Seagate Technology, Inc.
Refer to the License Agreement at the beginning of the DiscWizard
program.

Disk Manager is copyright 1985-1998 by ONTRACK Data International, Inc.
Refer to the License Agreement in the Disk Manager online manual.

*************************************************************************
*************************************************************************

BIOS LIMITATIONS
----------------
Included here are brief explanations of a number of drive capacity 
limitations that exist in the computer industry.  The use of Disk Manager 
and its Dynamic Drive Overlay offers a solution to each of these problems.

*** 528 MB Limitation ***

Using the traditional IDE interface limits the system to a maximum drive 
capacity of 528 MB.  The cause of this limitation is Int 13h (BIOS) and 
IDE field sizes for the CHS (Cylinder, Head, and Sector) entries.

Because the system must perform a translation between the CHS parameters
recognized by the drive and those established in the Int 13h code, 
parameters are limited to the smaller of the field sizes allowed for 
each parameter by the BIOS and the IDE register set. The chart below
displays the BIOS, IDE, and limiting field size.

			BIOS                IDE             Limit
Sectors per Track         63                255               63
Number of Heads          255                 16               16
Number of Cylinders     1024              65536             1024
			------           --------           ------
Maximum Capacity        8.4 GB           136.9 GB           528 MB

The maximum system drive capacity in a combined BIOS/IDE setup is 
determined by the limiting field size -- 528 MB.  Currently, computers 
are being shipped with a BIOS that implements Extended Int 13h or 
"Logical Block Addressing" (LBA), both of which are solutions to the 
528 MB limitation.

*** 4096 Cylinder (2.1 GB) Limitation ***

Some computers have a BIOS that does not properly deal with the "13th
bit". The 13th bit is needed to provide support for a drive having 
4096 or more cylinders.  The chart below displays the corresponding
cylinder values in decimal, hex, and binary values.

	DECIMAL     HEX      BINARY      SIZE
	  1023  =   3FF  =  10 bits  =  528 MB
	  2047  =   7FF  =  11 bits  =  1.0 GB
	  4095  =   FFF  =  12 bits  =  2.1 GB
	  8191  =  1FFF  =  13 bits  =  4.2 GB
	 16383  =  3FFF  =  14 bits  =  8.4 GB

If you have added a new drive and your system locks up at boot time 
(right after turning power on) or during System Setup, there may be 
several causes.  Verify that the data cable is properly attached to your 
drive, pin 1 is correct, and the cable is not installed off a row of pins.  
If your new drive is larger than 2.1GB and your System Setup (CMOS) is 
set to "AUTO", you may have a BIOS with a 4096 or greater cylinder 
limitation.  In this case, power off your system, remove your new drive, 
and follow the instructions that DiscWizard provides.  When configuring 
System Setup (CMOS), DO NOT USE AUTO.  Rather, choose one of the 
following:
     - USER DEFINABLE set to 1024 cyls 16 hds 63 sects
     - Drive type 1.
Another option is to contact your computer manufacturer to get a BIOS 
upgrade that will support more than 4096 cylinders.

*** 6322 Cylinder (3.27 GB) Limitation ***

Some computers have a BIOS that does not properly handle a cylinder value
over 6322.  If you are in the CMOS Setup attempting to set the cylinder 
value higher than 6322 (for a 3.27 GB+ drive) and your computer hangs, 
your computer may have a BIOS with this limitation.  To by-pass this 
limitation, you have two options:
     - Set the cylinder value to 1024 or less and use Ontrack's Disk
       Manager to provide support for the whole drive.
     - Contact your computer manufacturer for a BIOS upgrade, if one is
       available.

*** Invalid BIOS information ***

Some computers have a BIOS that may display invalid information in the 
CMOS setup. This issue may show up in one of two ways:
     - The CMOS will display the drive parameters and capacity correctly.
       However, it is not translating the drive correctly.
     - The CMOS will display invalid drive parameters. However, the BIOS
       is translating the drive correctly.  
To ensure your drive is translated to its full capacity, you
will need to check the actual drive size.  This can be done when creating
partitions on the drive.

*** 8.4 Gigabyte limit ***
If your drive is larger than 8.4 gigabytes, the capacity may exceed the 
limits of your system BIOS and operating system. Most system BIOS cannot support  
ATA drives this large. DOS and Windows operating systems limit the drive 
capacity to 8.4 Gigabytes per physical drive and 2 Gigabytes per partition. 
Because of these limitations, a 32-bit file allocation table (FAT32) is required 
to acheive full capacity of your drive beyond 8.4 Gigabytes.
To acheive full capacity of your drive you need a Windows operating system
that supports FAT32 and BIOS support for drives greater than 8.4 Gigabytes, from 
one of the following:
Third party device driver, such as Disk Manager (Disk  Manager is provided on 
the DiscWizard diskette included with your drive), or An intelligent ATA Host 
Adapter, or A system BIOS upgrade.
