# Note that this was found at lockergnome.com
# Note that this is similiar to the one called renew wmi on n-able.com except for my start and stop commands
# Mike Scallion, Lead NOC Engineer, masterIT - July 2008 Updated


winmgmt /clearadap
winmgmt /kill
winmgmt /unregserver
winmgmt /regserver
winmgmt /resyncperf 

# Stop WMI, not always necessary, but you never know...
net stop "Windows Management Instrumentation"
# This is an old DOS BATCH TIME DELAY Trick
ping 127.0.0.1 -n 5 > NUL
# Now answer yes to the, "do you also want to stop these services" remember this is just WMI
y
# Now restart WMI, Note other stopped services with it will not restart automatically, again
# most are not necessary to function and if exchange, simply restart using generic service restart within n-able
# again, pause first
ping 127.0.0.1 -n 3 > NUL
net start "Windows Management Instrumentation"
# For Vista, restart Security Center and IP Helper
net start "Security Center"
net start "IP Helper"
# For Exchange, restart Exchange Monitoring Piece.
net start "Microsoft Exchange Management"
# Disclaimer, this has worked on my machines, note that it may not work on all, but has on mine
# Not tested on 2008