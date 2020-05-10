Change to folder where the Provisioning Services Console is installed and run:

For 32-bit
%systemroot%\Microsoft.NET\Framework\v2.0.50727\installutil.exe McliPSSnapIn.dll

For 64-bit
%systemroot%\Microsoft.NET\Framework64\v2.0.50727\installutil.exe McliPSSnapIn.dll

If you are running 64-bit, you will need to run both commands so the snapin is registered for both
32-bit and 64-bit PowerShell.

Start a powershell session and change to the folder where you placed the script.

run Add-PSSnapin -Name McliPSSnapIn

To run the script:

.\pvs_inventory.ps1 | out-file .\PVSFarm.doc

Open the file in either Word or Wordpad.

I have tested this script with PVS 5.6 SP2, PVS 6.0 and PVS 6.1.

Please let me know what is missing or needs to be changed.

As of March 19, 2012, the script is as complete as I can get it based on my lab setup.  I will 
have to wait until I hear from others to know what else I need to add to the Status tab sections.

Thanks


Webster

