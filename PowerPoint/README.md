##README: Installing the Custom Ribbon and Macros in PowerPoint

#Introduction

This guide explains how to install the custom ribbon and macros included in the provided .pptm file, ensuring they are available every time PowerPoint starts.

Steps to Install the Macros and Ribbon

#1. Enable Macros in PowerPoint

Before proceeding, ensure that PowerPoint is set to allow macros:

Open PowerPoint.

Go to File > Options.

Select Trust Center > Trust Center Settings.

Click on Macro Settings.

Choose Enable all macros (or Disable all macros with notification for security).

Click OK.

#2. Save the .pptm File as a PowerPoint Add-in

To make the macros and ribbon always available:

Open the .pptm file in PowerPoint.

Go to File > Save As.

Select PowerPoint Add-in (*.ppam) as the file type.

Save it in the default PowerPoint Add-ins folder (usually C:\Users\YourUsername\AppData\Roaming\Microsoft\AddIns\).

#3. Load the Add-in into PowerPoint

Open PowerPoint.

Go to File > Options.

Select Add-ins.

In the Manage dropdown at the bottom, select PowerPoint Add-ins and click Go....

Click Add New... and navigate to the saved .ppam file.

Select the file and click OK.

Ensure the add-in is checked in the list and click Close.

#4. Verify the Custom Ribbon

Restart PowerPoint.

Check if the new ribbon tab appears in the PowerPoint menu.

Test one of the macros to confirm functionality.

Troubleshooting

Ribbon or macros do not appear: Ensure the add-in is properly enabled in the PowerPoint Add-ins menu.

Security warnings: If prompted, enable macros or adjust security settings in the Trust Center.

File path issues: Make sure the add-in is stored in a trusted location.

#Conclusion

Following these steps ensures that the custom ribbon and macros are always available when using PowerPoint. If you encounter issues, check PowerPoint’s macro and add-in settings.
