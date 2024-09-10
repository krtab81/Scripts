:: Obtener la MAC de una IP determinada

@echo off
ping -n 1 %1 >nul
arp -a | find "%1"