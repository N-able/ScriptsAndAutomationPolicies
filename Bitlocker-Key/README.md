# Retrieve / Store Bitlocker - Key 
The PS1 Script will pull the bitlocker key(for C: or D: drive) and display it if you run it as a script
The amp can be used to populate a custom property

## Bitlocker Status and Key
This is a simple solution to moving away from DEM
All this does is get the encryption status and the key then stores it in a custom device property

### GetEncryptionStatusAndKey.ps1
- This is the script that the AMP runs to get the information.

### Get Encryption Status and Key.amp
- An amp that we can run in N-Central to populate a CDP
