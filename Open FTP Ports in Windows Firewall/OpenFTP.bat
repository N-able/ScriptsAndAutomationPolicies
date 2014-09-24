rem Open FTP Ports in the Windows 2008/Windows 7/Windows Vista Firewall

rem This script will allow FTP to pass through on it's dynamic ports allowing PASV Mode to work.

rem  1) Open port 21 on the firewall

netsh advfirewall firewall add rule name="FTP (no SSL)" action=allow protocol=TCP dir=in localport=21

rem 2) Configure Dynamic Ports

netsh advfirewall set global StatefulFtp enable

