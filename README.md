# NewUserPC.ps1
Powershell script for automating the on-boarding process for new user's machines and a new hire packet
Recommended Usage
Create a directory called OnBoarding. In that directory create the following folders
UserData
PCData
Templates

When configuring the script, the User info csv should be saved in the UserData folder, PC Info csv in the PCData folder, and New Hire Packet template in the Templates folder.

The packet generation uses search and replace to populate info in the packet. Your template will need the following tags in their correct place
<first> First Name
<last> Last Name
<start> Start Date
<username>
<email>
<phone> (if able to be reached externally)
<ext>
