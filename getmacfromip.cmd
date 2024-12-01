:: Get MAC using IP address

@echo off
ping -n 1 %1 >nul
arp -a | find "%1"