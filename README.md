# TIPPanel

This program is developed back in 2013 in Visual Basic 6. I have written this in VB6, because the task was to be suitable for old machines using Windows 95/98.
This whole thing took me about 6 months to get it working as it should. Over the years till 2016 i have made some updates for adding some features and due to bugs, that clients found.

It is in usage by now in several Concrete Plants (maybe about 8-10) here in Bulgaria.
First prototype is established for testing on a double concrete machine in Ruse (company name is Intis). 
PostgreSQL is used for local database server and OPC Omron is used as a connection with the Omron PLC's.
It is capable of exporting xls-files.

The sofware consist of three main parts, each of which has it special purpose.
The main part TIPDispatcher is the one for dipatching everything needed for the machine to operate and collecting back the data from the PLC itself.
The second TIPReporter is for the managers to check their database without leaving the office.
The third i wrote for testing. I could not afford to buy a PLC, so i needed something to do the work and send data for testing the program. This parts imitates the machine (PLC) work itself. Note that i had no access the PLC program, so i wrote this by studying how the machine works.

These were my very first steps in coding a real working program. The code works till today, although is poorly structured (no conventions at all) and maybe more than 15 000 lines as i remember. I am prety sure that 20 percent can be reduced. But no VB6 anymore :).

Now i am willing to revive this one is a new project in Java, hope i will have the time ...... :)
